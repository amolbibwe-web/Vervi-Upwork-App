#!/usr/bin/env python3
"""
upwork_to_journal.py
====================

Converts a raw Upwork account statement into double-entry accounting journal
entries, using an externally-maintained mapping workbook (Sheet2) so that new
freelancers, wallets or transaction types can be added WITHOUT editing code.

Pipeline
--------
    Sheet1 (statement)  ─┐
    Sheet2 (mapping)    ─┼─►  build journal lines  ─►  balance check  ─►  Excel + CSV
    FX rates            ─┘                                    │
                                                              └─►  exceptions report

Accounting treatment (all driven by Sheet2, see `parse_treatments`)
------------------------------------------------------------------
*   EXPENSE_RCM  (Service Fee, Connects, Subscription, Withdrawal Fee)
        Dr  <Expense ledger>      base
        Dr  RCM Input IGST        base x rate
            Cr  <Wallet GL>                 base
            Cr  RCM Output IGST             base x rate
        The two IGST legs are equal and opposite, so the reverse-charge
        liability nets to zero while both halves stay visible for GSTR filing.

*   SIMPLE       (WHT)
        Dr  TDS                   base
            Cr  <Wallet GL>                 base

*   INCOME       (Hourly, Fixed-price) -- two shapes, selected by --income-mode:
        combined (default):
            Dr  <Wallet GL>       base
                Cr  Revenue - TPT Export    base
        two-entry:
            Dr  AR <Client>       base   /  Cr  Revenue - TPT Export   base
            Dr  <Wallet GL>       base   /  Cr  AR <Client>            base

*   SKIP         (Withdrawal) -- no entry, logged to the skipped report.

Currency
--------
The statement is denominated in USD.  With --currency INR (the default) each
amount is converted at the RBI USD/INR reference rate *for its own transaction
date*, fetched and cached by rbi_fx.py -- see that module for the fetch, cache
and previous-day-fallback rules.  Precedence per date:

    --fx-rates file  >  RBI / local cache  >  --fx-rate (manual safety net)

The original USD figure, the rate, the date that rate came from and its source
are all retained on every journal line, and summarised on an FX Audit sheet.
--fx-source manual restores the old behaviour of one flat rate for everything.

Exit codes
----------
    0  clean run
    1  at least one voucher failed the debit==credit check, or --strict was
       passed and any ERROR-severity exception was recorded
    2  a fatal input problem (missing file, unreadable mapping, no FX rate)
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable

import pandas as pd

from rbi_fx import FxProvider, RbiFxProvider, load_fx_file_by_date

# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #

#: The accounting master database, used whenever --mapping is not given.
#: Table A (treatment by transaction type) and Table B (Account Name -> GL Name)
#: both live here; adding a freelancer means editing this file, nothing else.
DEFAULT_MASTER_MAPPING = Path(__file__).parent / "master" / "mapping_master.csv"

#: Running record of every document number issued, so a later statement -- a new
#: client, a new month -- continues the series instead of restarting it.
DEFAULT_DOC_REGISTRY = Path(__file__).parent / "master" / "doc_registry.csv"

#: Every Ref ID already posted to the books. Upwork's Ref ID is unique per
#: transaction, so it is the key that tells a genuinely new row from one that
#: arrived again in an overlapping statement.
DEFAULT_POSTED_LEDGER = Path(__file__).parent / "master" / "posted_refs.csv"

#: Opening positions, used when no registry exists yet. This one IS committed to
#: the repo, so a hosted instance with an empty disk still starts at the right
#: number rather than at 001.
DOC_REGISTRY_SEED_NAME = "doc_registry_seed.csv"

#: Cell token in Sheet2 meaning "substitute the wallet ledger from Table B".
GL_PLACEHOLDER = "gl name"

#: Cell token meaning "substitute the row's Client team" -- the party ledger that
#: revenue is billed to and later cleared against the Upwork wallet.
CLIENT_PLACEHOLDER = "client team"

#: Voucher types as the accounting system expects them, and the document-number
#: series each one draws from.  {fy} = Indian financial year (26-27),
#: {mon} = Jul, {seq} = a counter that restarts whenever the prefix changes --
#: so JE restarts every month and Sales every financial year.
VOUCHER_SALES = "Sales"
VOUCHER_JE = "JE"
DEFAULT_DOC_SERIES = {
    VOUCHER_JE: "{fy}/LLP/{mon}/{seq:03d}",
    VOUCHER_SALES: "{fy}/VASL/U/{seq:03d}",
}

#: Columns of the import file, in the order the accounting system reads them.
IMPORT_COLUMNS = [
    "Subsidiary", "Transaction date", "Period", "Document Number", "Type",
    "Ledger Name", "Amount", "Debit", "Credit", "Amount in base currency",
    "Amount in INR", "Currency", "Exchange Rate", "I/E", "Cost center Name",
    "Narration",
    # After every column the accounting system reads, so importing is unaffected.
    # It is the key that says which Upwork transaction a line came from, which is
    # what makes an imported file traceable back to the statement.
    "Ref ID",
]

#: Cell tokens meaning "this leg is not used".
NULL_TOKENS = {"", "na", "n/a", "nan", "none", "-", "nil"}

#: Cell token marking a transaction type that produces no voucher at all.
NO_ENTRY_TOKEN = "no entry"

#: Ledger names that identify the reverse-charge legs inside Sheet2.
RCM_INPUT = "rcm input igst"
RCM_OUTPUT = "rcm output igst"

#: Money is handled as Decimal throughout; this is the rounding quantum.
CENTS = Decimal("0.01")

#: Tolerance for the per-voucher balance assertion (guards against a stray
#: half-paisa from independent rounding of the two IGST legs).
BALANCE_TOLERANCE = Decimal("0.01")

#: Logical statement column -> candidate header spellings (normalised).
STATEMENT_COLUMNS: dict[str, tuple[str, ...]] = {
    "date": ("date",),
    "transaction_id": ("transactionid",),
    "txn_type": ("transactiontype",),
    "summary": ("transactionsummary",),
    "summary_details": ("transactionsummarydetails",),
    "desc1": ("description1",),
    "desc2": ("description2",),
    "desc3": ("description3",),
    "agency_team": ("agencyteam",),
    "freelancer": ("freelancer",),
    "client_team": ("clientteam",),
    "account_name": ("accountname",),
    "po": ("po",),
    "ref_id": ("refid",),
    "amount_usd": ("amount", "amountusd", "amount$"),
    "amount_local": ("amountinlocalcurrency",),
    "currency": ("currency",),
    "balance": ("currentbalance", "currentbalance$"),
    "payment_method": ("paymentmethod",),
}

#: Columns without which we cannot post anything.
REQUIRED_STATEMENT_COLUMNS = ("date", "txn_type", "amount_usd", "account_name")

#: Output column order for the journal.
JOURNAL_COLUMNS = [
    "Sr. No.", "Document Number", "Voucher Type", "Date", "Period",
    "Transaction ID", "Ref ID", "Transaction Type", "Subsidiary", "Account Name",
    "Cost Center", "Client Team", "Ledger", "Dr/Cr",
    "Debit", "Credit", "Currency", "Amount USD", "FX Rate", "Rate Date",
    "Rate Source", "Narration",
]


# --------------------------------------------------------------------------- #
# Text / number normalisation
# --------------------------------------------------------------------------- #

def clean_text(value: Any) -> str:
    """Trim a spreadsheet cell to comparable text.

    Upwork exports (and hand-edited mapping sheets) routinely carry non-breaking
    spaces, zero-width joiners and doubled spaces -- e.g. the sample Sheet2 has
    ``Upwork\\xa0Rahul\\xa0Menon`` and ``Vikram  Nair``.  Left alone these
    silently break dictionary lookups, so every key and value passes through
    here: Unicode-normalise, fold all whitespace (NBSP included) to plain
    spaces, collapse runs, and strip.
    """
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    # NFKC already folds NBSP to a plain space; these cover the zero-width
    # characters it leaves behind, which are invisible but break equality.
    text = text.replace("\u200b", "").replace("\ufeff", "")
    return re.sub(r"\s+", " ", text).strip()


def norm_key(value: Any) -> str:
    """Case-insensitive lookup key (whitespace already folded by clean_text)."""
    return clean_text(value).casefold()


def norm_header(value: Any) -> str:
    """Header key with punctuation and spaces removed: 'Amount $' -> 'amount'."""
    return re.sub(r"[^a-z0-9]", "", clean_text(value).casefold())


def is_null_cell(value: Any) -> bool:
    """True when a mapping cell means 'no ledger here'."""
    return norm_key(value) in NULL_TOKENS


def to_decimal(value: Any) -> Decimal | None:
    """Parse a spreadsheet amount to Decimal, or None if it isn't a number.

    Tolerates currency symbols, thousands separators and accounting-style
    parentheses for negatives, since these survive some Excel round-trips.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value))
        except InvalidOperation:
            return None

    text = clean_text(value)
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = re.sub(r"[^0-9.\-]", "", text.strip("()"))
    if text in ("", "-", "."):
        return None
    try:
        amount = Decimal(text)
    except InvalidOperation:
        return None
    return -amount if negative else amount


def money(value: Decimal) -> Decimal:
    """Round half-up to 2 decimals -- the convention Indian books expect."""
    return value.quantize(CENTS, rounding=ROUND_HALF_UP)


# --------------------------------------------------------------------------- #
# Generic tabular file loading
# --------------------------------------------------------------------------- #

def read_raw(path: Path, sheet: str | int | None = None) -> list[list[Any]]:
    """Read a .csv/.xlsx into a plain grid of cells, no header interpretation.

    Returning a raw grid (rather than a DataFrame) is what lets us find several
    tables laid out side by side on one sheet, as the sample Sheet2 does.
    """
    suffix = path.suffix.lower()
    if suffix == ".csv":
        # utf-8-sig strips the BOM Excel writes; latin-1 is a last-resort
        # fallback so a mis-encoded export still loads instead of crashing.
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                frame = pd.read_csv(path, header=None, dtype=object,
                                    keep_default_na=False, encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError(f"Could not decode {path.name} with any known encoding")
    elif suffix in (".xlsx", ".xlsm", ".xls"):
        frame = pd.read_excel(path, sheet_name=sheet if sheet is not None else 0,
                              header=None, dtype=object, engine=None)
    else:
        raise ValueError(f"Unsupported file type '{suffix}' for {path.name}")
    return frame.values.tolist()


def excel_sheet_names(path: Path) -> list[str]:
    """Sheet names for a workbook, or [] for CSV."""
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return pd.ExcelFile(path).sheet_names
    return []


def find_header(grid: list[list[Any]], wanted: str) -> tuple[int, int] | None:
    """Locate the (row, col) of a header cell matching `wanted` (normalised)."""
    target = norm_header(wanted)
    for r, row in enumerate(grid):
        for c, cell in enumerate(row):
            if norm_header(cell) == target:
                return r, c
    return None


def cell(grid: list[list[Any]], r: int, c: int) -> Any:
    """Bounds-safe grid access (ragged CSV rows are common)."""
    if 0 <= r < len(grid) and 0 <= c < len(grid[r]):
        return grid[r][c]
    return None


# --------------------------------------------------------------------------- #
# Mapping model
# --------------------------------------------------------------------------- #

@dataclass
class Treatment:
    """One row of Table A -- how a transaction type is to be posted."""
    nature: str                 # as written in Sheet2, for reporting
    kind: str                   # EXPENSE_RCM | SIMPLE | INCOME | SKIP
    debit_1: str
    debit_2: str
    credit_1: str
    credit_2: str


@dataclass
class Mapping:
    """Everything read out of the mapping file."""
    treatments: dict[str, Treatment] = field(default_factory=dict)
    wallets: dict[str, str] = field(default_factory=dict)
    #: account names that appeared with DIFFERENT GLs and could NOT be resolved
    #: by name match -- rows using them are refused rather than guessed at.
    ambiguous_wallets: dict[str, list[str]] = field(default_factory=dict)
    #: duplicates that WERE resolved by name match: key -> (chosen, rejected)
    resolved_wallets: dict[str, tuple[str, list[str]]] = field(default_factory=dict)
    fx_rates: dict[str, Decimal] = field(default_factory=dict)
    #: display-cased account names, for readable exception messages
    wallet_display: dict[str, str] = field(default_factory=dict)
    #: Account Name -> Cost Center (Table B, column I).  Optional: an account
    #: with no cost centre falls back to the statement's Freelancer.
    cost_centers: dict[str, str] = field(default_factory=dict)
    #: Account Name -> Subsidiary (Table B, column J).  Optional: an account
    #: with no subsidiary falls back to the --entity default.
    subsidiaries: dict[str, str] = field(default_factory=dict)


#: Tokens that carry no identifying information when matching an account name
#: against a candidate GL -- every wallet ledger starts with "Upwork".
GL_STOPWORDS = {"upwork", "gl", "ledger", "ac", "a/c", "account"}

#: A duplicate is only auto-resolved when the best candidate scores at least
#: this well, and beats the runner-up by at least MATCH_MARGIN.
MATCH_THRESHOLD = 0.72
MATCH_MARGIN = 0.15


def name_match_score(account: str, gl: str) -> float:
    """How strongly a wallet GL looks like it belongs to `account`.

    Mapping sheets are maintained by hand, so a duplicated account row is
    usually a paste slip -- the same name pasted against someone else's ledger.
    The genuine row is the one whose ledger actually carries the person's name,
    which this scores three ways:

      * exact token overlap    'Arjun Bhatt'    vs 'Upwork Arjun'   -> 1.00
      * initials              'Priya Sharma' vs 'Upwork PS'       -> 1.00
      * fuzzy token similarity 'Meera Iyer' vs 'Upwork Meera' -> ~0.88

    "Upwork" and friends are ignored, since every ledger has them and they would
    otherwise make every candidate look equally good.
    """
    account_tokens = [t for t in re.split(r"\W+", norm_key(account)) if t]
    gl_tokens = [t for t in re.split(r"\W+", norm_key(gl)) if t and t not in GL_STOPWORDS]
    if not account_tokens or not gl_tokens:
        return 0.0

    # Initials: 'Priya Sharma' -> 'mb', which is how Table B writes that GL.
    initials = "".join(t[0] for t in account_tokens)
    if len(initials) > 1 and any(t == initials for t in gl_tokens):
        return 1.0

    best = 0.0
    for a in account_tokens:
        for g in gl_tokens:
            if a == g:
                return 1.0
            best = max(best, SequenceMatcher(None, a, g).ratio())
    return best


def resolve_duplicate(account: str, candidates: list[str]) -> tuple[str | None, float]:
    """Pick the GL that belongs to `account`, or (None, score) if it's a toss-up.

    Requires both an absolute confidence floor and a clear margin over the
    runner-up, so a genuinely ambiguous pair still gets escalated rather than
    silently resolved the wrong way.
    """
    scored = sorted(((name_match_score(account, gl), gl) for gl in candidates),
                    key=lambda pair: pair[0], reverse=True)
    best_score, best_gl = scored[0]
    runner_up = scored[1][0] if len(scored) > 1 else 0.0
    if best_score >= MATCH_THRESHOLD and (best_score - runner_up) >= MATCH_MARGIN:
        return best_gl, best_score
    return None, best_score


def classify(debit_1: str, debit_2: str, credit_1: str, credit_2: str,
             explicit_kind: str = "") -> str:
    """Decide how a Table A row should be posted.

    An explicit `Kind` column in Sheet2 always wins, so an unusual treatment can
    be forced without code changes.  Otherwise the shape of the row tells us:

    * every leg says "NO Entry"                  -> SKIP
    * the second debit leg is RCM Input IGST     -> EXPENSE_RCM
    * debit-1 reappears as credit-2 (the AR      -> INCOME
      round-trip Sheet2 uses for earnings)
    * anything else                              -> SIMPLE (one Dr, one Cr)
    """
    if explicit_kind:
        return explicit_kind.upper().replace(" ", "_").replace("-", "_")

    legs = [debit_1, debit_2, credit_1, credit_2]
    if all(norm_key(leg) == NO_ENTRY_TOKEN for leg in legs if clean_text(leg)):
        return "SKIP"
    if norm_key(debit_2) == RCM_INPUT:
        return "EXPENSE_RCM"
    if (not is_null_cell(debit_1) and not is_null_cell(credit_2)
            and norm_key(debit_1) == norm_key(credit_2)):
        return "INCOME"
    return "SIMPLE"


def parse_treatments(grid: list[list[Any]], mapping: Mapping) -> None:
    """Read Table A (headed by 'Nature') out of the grid.

    Layout assumed: Nature | Debit | Debit | Credit | Credit [| Kind].
    Rows are consumed downward until the Nature cell goes blank, so Table A can
    be shorter than Table B on the same sheet (as it is in the sample).
    """
    found = find_header(grid, "Nature")
    if not found:
        raise ValueError("Mapping file has no 'Nature' header -- Table A not found")
    header_row, col = found

    # Optional Kind / Subsidiary columns may sit anywhere in the cells after
    # Nature. Scanned within this table's own block rather than by searching the
    # whole sheet, so Table B's Subsidiary column is never mistaken for this one.
    kind_col = None
    for offset in range(1, 7):
        if norm_header(cell(grid, header_row, col + offset)) == "kind":
            kind_col = col + offset
            break

    for r in range(header_row + 1, len(grid)):
        nature = clean_text(cell(grid, r, col))
        if not nature:
            # Skip, don't stop: deleting a rule leaves a gap mid-table, and
            # stopping here would silently drop every rule below it.
            continue
        d1 = clean_text(cell(grid, r, col + 1))
        d2 = clean_text(cell(grid, r, col + 2))
        c1 = clean_text(cell(grid, r, col + 3))
        c2 = clean_text(cell(grid, r, col + 4))
        explicit = clean_text(cell(grid, r, kind_col)) if kind_col is not None else ""
        mapping.treatments[norm_key(nature)] = Treatment(
            nature=nature, kind=classify(d1, d2, c1, c2, explicit),
            debit_1=d1, debit_2=d2, credit_1=c1, credit_2=c2,
        )


def parse_wallets(grid: list[list[Any]], mapping: Mapping, *,
                  on_duplicate: str = "resolve") -> None:
    """Read Table B (headed by 'Account Name') out of the grid.

    A repeated account name is only a problem when the rows disagree; an exact
    duplicate is harmless.  Conflicting duplicates are handed to
    `resolve_duplicate`, which picks the ledger whose name matches the account
    (with --on-duplicate fail, or when the match is too close to call, the
    conflict is escalated instead and rows for that account are refused).
    """
    found = find_header(grid, "Account Name")
    if not found:
        raise ValueError("Mapping file has no 'Account Name' header -- Table B not found")
    header_row, col = found

    # Gather every candidate first: resolution needs to see the full set, not
    # just whichever row happened to be read first.
    candidates: dict[str, list[str]] = {}
    for r in range(header_row + 1, len(grid)):
        account = clean_text(cell(grid, r, col))
        gl = clean_text(cell(grid, r, col + 1))
        if not account or not gl:
            continue  # Table B may have gaps; keep scanning to the sheet end
        key = norm_key(account)
        mapping.wallet_display.setdefault(key, account)
        # Column I, optional -- blank means "use the statement's Freelancer".
        cost_center = clean_text(cell(grid, r, col + 2))
        if cost_center:
            mapping.cost_centers.setdefault(key, cost_center)
        # Column J, optional -- blank means "use the --entity default".
        subsidiary = clean_text(cell(grid, r, col + 3))
        if subsidiary:
            mapping.subsidiaries.setdefault(key, subsidiary)
        seen = candidates.setdefault(key, [])
        if not any(norm_key(gl) == norm_key(existing) for existing in seen):
            seen.append(gl)

    for key, options in candidates.items():
        display = mapping.wallet_display.get(key, key)
        if len(options) == 1:
            mapping.wallets[key] = options[0]
            continue
        if on_duplicate == "fail":
            mapping.ambiguous_wallets[key] = options
            continue
        chosen, score = resolve_duplicate(display, options)
        if chosen is None:
            mapping.ambiguous_wallets[key] = options
        else:
            mapping.wallets[key] = chosen
            mapping.resolved_wallets[key] = (
                chosen, [gl for gl in options if gl != chosen])


def parse_fx_table(grid: list[list[Any]], mapping: Mapping) -> None:
    """Read an optional FX table (headers 'Period' and 'Rate'/'FX Rate').

    Periods are 'YYYY-MM' month keys; the literal '*' or 'default' supplies a
    catch-all used when a month has no explicit rate.
    """
    found = find_header(grid, "Period")
    if not found:
        return
    header_row, col = found

    rate_col = None
    for offset in (1, 2, 3):
        if norm_header(cell(grid, header_row, col + offset)) in ("rate", "fxrate", "usdinr"):
            rate_col = col + offset
            break
    if rate_col is None:
        rate_col = col + 1

    for r in range(header_row + 1, len(grid)):
        period = clean_text(cell(grid, r, col))
        if not period:
            break
        rate = to_decimal(cell(grid, r, rate_col))
        if rate is not None and rate > 0:
            mapping.fx_rates[norm_key(period)] = rate


def load_mapping(path: Path, *, on_duplicate: str = "resolve") -> Mapping:
    """Load Table A, Table B and any FX table from the mapping file.

    Handles both sample layouts: several tables side by side on one sheet
    (the provided Sheet2.csv), or one table per worksheet in an .xlsx.
    """
    mapping = Mapping()
    sheets = excel_sheet_names(path)
    grids = [read_raw(path, s) for s in sheets] if sheets else [read_raw(path)]

    for grid in grids:
        if not mapping.treatments and find_header(grid, "Nature"):
            parse_treatments(grid, mapping)
        if not mapping.wallets and find_header(grid, "Account Name"):
            parse_wallets(grid, mapping, on_duplicate=on_duplicate)
        parse_fx_table(grid, mapping)

    if not mapping.treatments:
        raise ValueError("No Table A ('Nature' header) found in the mapping file")
    if not mapping.wallets:
        raise ValueError("No Table B ('Account Name' header) found in the mapping file")
    return mapping


# --------------------------------------------------------------------------- #
# Editing the master database
# --------------------------------------------------------------------------- #

def _write_grid(path: Path, grid: list[list[Any]]) -> None:
    """Write a grid back as CSV, padded to a uniform width."""
    width = max((len(row) for row in grid), default=0)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        for row in grid:
            padded = list(row) + [""] * (width - len(row))
            writer.writerow(["" if v is None or (isinstance(v, float) and pd.isna(v))
                             else v for v in padded])


def _append_to_table(path: Path, header: str, values: list[str]) -> None:
    """Add a row to whichever table starts at `header`, in place.

    The two tables sit side by side on one sheet with different lengths, so a
    new row goes into the first blank slot underneath its own table -- reusing a
    row the other table already occupies rather than pushing the sheet longer
    than it needs to be.
    """
    if path.suffix.lower() != ".csv":
        raise ValueError("Only a .csv master database can be edited here; "
                         "edit an .xlsx master in Excel")

    grid = [list(row) for row in read_raw(path)]
    found = find_header(grid, header)
    if not found:
        raise ValueError(f"No '{header}' table found in {path.name}")
    header_row, col = found

    # First row under this table whose own cells are empty.
    target = None
    for r in range(header_row + 1, len(grid)):
        if not clean_text(cell(grid, r, col)):
            target = r
            break
    if target is None:
        target = len(grid)
        grid.append([])

    width = max(max((len(row) for row in grid), default=0), col + len(values))
    while len(grid[target]) < width:
        grid[target].append("")
    for offset, value in enumerate(values):
        grid[target][col + offset] = value

    _write_grid(path, grid)


def _row_of(grid: list[list[Any]], header: str, key: str) -> tuple[int, int]:
    """Locate the row under `header` whose first cell matches `key`."""
    found = find_header(grid, header)
    if not found:
        raise ValueError(f"No '{header}' table found in the master database")
    header_row, col = found
    for r in range(header_row + 1, len(grid)):
        if norm_key(cell(grid, r, col)) == norm_key(key):
            return r, col
    raise ValueError(f"'{key}' is not in the master database")


def _write_row(path: Path, header: str, key: str, values: list[str]) -> None:
    """Overwrite one row of a table in place, leaving the other table alone.

    The two tables share rows on one sheet, so only this table's own cells are
    touched -- a treatment sitting on the same line is never disturbed.
    """
    if path.suffix.lower() != ".csv":
        raise ValueError("Only a .csv master database can be edited here; "
                         "edit an .xlsx master in Excel")
    grid = [list(row) for row in read_raw(path)]
    r, col = _row_of(grid, header, key)
    width = max(max((len(row) for row in grid), default=0), col + len(values))
    while len(grid[r]) < width:
        grid[r].append("")
    for offset, value in enumerate(values):
        grid[r][col + offset] = value
    _write_grid(path, grid)


def update_wallet(path: Path, key: str, account: str, gl_name: str,
                  cost_center: str = "", subsidiary: str = "") -> None:
    """Rewrite the Table B row currently keyed on `key`."""
    account, gl_name = clean_text(account), clean_text(gl_name)
    cost_center, subsidiary = clean_text(cost_center), clean_text(subsidiary)
    if not account or not gl_name:
        raise ValueError("Both an account name and a GL name are required")
    existing = load_mapping(path).wallets
    # Renaming onto another account would give one name two ledgers.
    if norm_key(account) != norm_key(key) and norm_key(account) in existing:
        raise ValueError(f"'{account}' already exists in the master database")
    _write_row(path, "Account Name", key, [account, gl_name, cost_center, subsidiary])


def delete_wallet(path: Path, key: str) -> None:
    """Clear a Table B row. The line stays so Table A keeps its position."""
    _write_row(path, "Account Name", key, ["", "", "", ""])


def update_treatment(path: Path, key: str, nature: str, debit_1: str, debit_2: str,
                     credit_1: str, credit_2: str) -> None:
    """Rewrite the Table A row currently keyed on `key`."""
    nature = clean_text(nature)
    if not nature:
        raise ValueError("A transaction type is required")
    if not clean_text(debit_1) or not clean_text(credit_1):
        raise ValueError("At least one debit and one credit ledger are required")
    existing = load_mapping(path).treatments
    if norm_key(nature) != norm_key(key) and norm_key(nature) in existing:
        raise ValueError(f"A rule for '{nature}' already exists")
    _write_row(path, "Nature", key, [
        nature, clean_text(debit_1), clean_text(debit_2) or "NA",
        clean_text(credit_1), clean_text(credit_2) or "NA"])


def delete_treatment(path: Path, key: str) -> None:
    """Clear a Table A row. The line stays so Table B keeps its position."""
    _write_row(path, "Nature", key, ["", "", "", "", ""])


def add_wallet(path: Path, account: str, gl_name: str, cost_center: str = "",
               subsidiary: str = "") -> None:
    """Add an Account Name -> GL Name -> Cost Center -> Subsidiary row to Table B."""
    account, gl_name = clean_text(account), clean_text(gl_name)
    cost_center, subsidiary = clean_text(cost_center), clean_text(subsidiary)
    if not account or not gl_name:
        raise ValueError("Both an account name and a GL name are required")
    existing = load_mapping(path).wallets
    if norm_key(account) in existing:
        raise ValueError(f"'{account}' is already mapped to "
                         f"'{existing[norm_key(account)]}'")
    _append_to_table(path, "Account Name", [account, gl_name, cost_center, subsidiary])


def add_treatment(path: Path, nature: str, debit_1: str, debit_2: str,
                  credit_1: str, credit_2: str) -> None:
    """Add a transaction-type row to Table A."""
    nature = clean_text(nature)
    if not nature:
        raise ValueError("A transaction type is required")
    if not clean_text(debit_1) or not clean_text(credit_1):
        raise ValueError("At least one debit and one credit ledger are required")
    if norm_key(nature) in load_mapping(path).treatments:
        raise ValueError(f"A rule for '{nature}' already exists")
    _append_to_table(path, "Nature", [
        nature, clean_text(debit_1), clean_text(debit_2) or "NA",
        clean_text(credit_1), clean_text(credit_2) or "NA"])


def load_fx_file(path: Path) -> dict[str, Decimal]:
    """Load a standalone Period,Rate file supplied via --fx-rates."""
    mapping = Mapping()
    for grid in ([read_raw(path, s) for s in excel_sheet_names(path)] or [read_raw(path)]):
        parse_fx_table(grid, mapping)
    if not mapping.fx_rates:
        raise ValueError(f"No 'Period'/'Rate' table found in {path.name}")
    return mapping.fx_rates


# --------------------------------------------------------------------------- #
# Statement loading
# --------------------------------------------------------------------------- #

def load_statement(path: Path) -> pd.DataFrame:
    """Load Sheet1 and rename its columns to stable logical names.

    Header matching is punctuation-insensitive so 'Amount $', 'Amount$' and
    'amount' all resolve to `amount_usd`.
    """
    sheets = excel_sheet_names(path)
    grid = read_raw(path, sheets[0] if sheets else None)
    if not grid:
        raise ValueError(f"{path.name} is empty")

    # The header is the first row mentioning 'Transaction type'; anything above
    # it (title rows, export banners) is discarded.
    header_row = 0
    for r, row in enumerate(grid[:20]):
        if any(norm_header(c) == "transactiontype" for c in row):
            header_row = r
            break

    headers = [norm_header(c) for c in grid[header_row]]
    resolved: dict[int, str] = {}
    for logical, candidates in STATEMENT_COLUMNS.items():
        for idx, head in enumerate(headers):
            if head in candidates and idx not in resolved:
                resolved[idx] = logical
                break

    missing = [c for c in REQUIRED_STATEMENT_COLUMNS if c not in resolved.values()]
    if missing:
        raise ValueError(
            f"{path.name} is missing required column(s): {', '.join(missing)}")

    records = []
    for row in grid[header_row + 1:]:
        record = {logical: cell([row], 0, idx) for idx, logical in resolved.items()}
        if any(clean_text(v) for v in record.values()):  # drop blank filler rows
            records.append(record)
    return pd.DataFrame.from_records(records)


# --------------------------------------------------------------------------- #
# Journal construction
# --------------------------------------------------------------------------- #

@dataclass
class Leg:
    """One side of one voucher line."""
    ledger: str
    side: str        # "Dr" or "Cr"
    amount: Decimal


@dataclass
class Voucher:
    """A balanced set of legs, plus the voucher type driving its numbering."""
    legs: list[Leg]
    kind: str = VOUCHER_JE


@dataclass
class Context:
    """Per-row facts the leg builders need."""
    wallet_gl: str
    client: str
    base: Decimal            # positive base amount in the output currency
    igst: Decimal            # base x igst_rate, already rounded


def resolve_ledger(name: str, ctx: Context) -> str:
    """Expand a master-database cell into a real ledger name.

    'GL Name'     -> the row's wallet ledger (Table B, column H)
    'Client Team' -> the row's Client team, the party revenue is billed to
    'AR <x>'      -> legacy template form, re-pointed at the row's client

    Templating rather than hard-coding is what lets one statement carry several
    clients without every receivable landing on the same party.
    """
    key = norm_key(name)
    if key == GL_PLACEHOLDER:
        return ctx.wallet_gl
    if key == CLIENT_PLACEHOLDER:
        return ctx.client or name
    if ctx.client and re.match(r"^ar\b", key):
        return f"AR {ctx.client}"
    return name


def build_legs(treatment: Treatment, ctx: Context, income_mode: str) -> list[Voucher]:
    """Turn a treatment + row context into one or more balanced vouchers.

    Income in two-entry mode is the only case that yields more than one voucher:
    a Sales voucher recognising revenue against the client, then a JE clearing
    that client balance against the Upwork wallet the money actually landed in.
    """
    if treatment.kind == "EXPENSE_RCM":
        # Base expense plus a self-cancelling pair of reverse-charge IGST legs.
        legs = [
            Leg(resolve_ledger(treatment.debit_1, ctx), "Dr", ctx.base),
            Leg(resolve_ledger(treatment.debit_2, ctx), "Dr", ctx.igst),
            Leg(resolve_ledger(treatment.credit_1, ctx), "Cr", ctx.base),
            Leg(resolve_ledger(treatment.credit_2, ctx), "Cr", ctx.igst),
        ]
        return [Voucher([leg for leg in legs if leg.amount != 0], VOUCHER_JE)]

    if treatment.kind == "INCOME":
        revenue = resolve_ledger(treatment.debit_2, ctx)   # Revenue - TPT Export
        client = resolve_ledger(treatment.debit_1, ctx)    # the Client team party
        wallet = resolve_ledger(treatment.credit_1, ctx)   # Upwork wallet GL
        if income_mode == "two-entry":
            return [
                # 1. record the sale against the client
                Voucher([Leg(client, "Dr", ctx.base),
                         Leg(revenue, "Cr", ctx.base)], VOUCHER_SALES),
                # 2. move that balance onto the Upwork party that holds the cash
                Voucher([Leg(wallet, "Dr", ctx.base),
                         Leg(client, "Cr", ctx.base)], VOUCHER_JE),
            ]
        # combined: the client balance nets out in the same instant, so post the
        # wallet straight against revenue.
        return [Voucher([Leg(wallet, "Dr", ctx.base),
                         Leg(revenue, "Cr", ctx.base)], VOUCHER_SALES)]

    # SIMPLE -- a single Dr/Cr pair (WHT -> Dr TDS / Cr wallet).
    return [Voucher([
        Leg(resolve_ledger(treatment.debit_1, ctx), "Dr", ctx.base),
        Leg(resolve_ledger(treatment.credit_1, ctx), "Cr", ctx.base),
    ], VOUCHER_JE)]


def financial_year(day) -> str:
    """Indian financial year label for a date: Jul 2026 -> '26-27'.

    The year runs April to March, so anything before April belongs to the year
    that started the previous April.
    """
    start = day.year if day.month >= 4 else day.year - 1
    return f"{start % 100:02d}-{(start + 1) % 100:02d}"


#: Columns of the document-number registry.
REGISTRY_COLUMNS = ("document_number", "type", "date", "prefix", "seq",
                    "issued_at", "source")


#: Columns of the posted-transactions ledger.
POSTED_COLUMNS = ("ref_id", "transaction_id", "date", "transaction_type",
                  "amount_usd", "document_number", "source", "posted_at")


class PostedLedger:
    """Remembers which Upwork transactions have already been journalised.

    Statements are pulled every fortnight and overlap, and an old file gets
    re-uploaded by accident. Ref ID is unique per Upwork transaction, so keeping
    the ones already posted lets a re-run skip them instead of double-booking.

    Like document numbers, entries are held back until the export is actually
    downloaded -- previewing a statement must not mark it as posted.
    """

    def __init__(self, path: Path | None = None, *, source: str = "",
                 enabled: bool = True) -> None:
        self.path = path
        self.source = source
        self.enabled = enabled and path is not None
        self.seen: dict[str, dict[str, str]] = {}
        self._pending: list[dict[str, Any]] = []
        if self.enabled:
            self._load()

    def _load(self) -> None:
        if not self.path.exists():
            return
        try:
            with self.path.open(newline="", encoding="utf-8") as handle:
                for row in csv.DictReader(handle):
                    key = norm_key(row.get("ref_id", ""))
                    if key:
                        self.seen.setdefault(key, row)
        except OSError:
            # Refuse to run blind: silently treating the ledger as empty would
            # re-post everything it was meant to protect.
            raise ValueError(
                f"Posted-transaction ledger {self.path} exists but could not be "
                f"read; refusing to run rather than risk double-posting")

    def already_posted(self, ref: str) -> dict[str, str] | None:
        """The earlier posting for this Ref ID, or None if it is new."""
        if not self.enabled:
            return None
        return self.seen.get(norm_key(ref))

    def record(self, ref: str, row: dict[str, Any], doc_no: str,
               txn_type: str, amount_usd: Any) -> None:
        """Note a Ref ID as posted -- pending until `save` is called."""
        if not self.enabled or not clean_text(ref):
            return
        entry = {
            "ref_id": clean_text(ref),
            "transaction_id": clean_text(row.get("transaction_id")),
            "date": clean_text(row.get("date")),
            "transaction_type": txn_type,
            "amount_usd": f"{amount_usd}",
            "document_number": doc_no,
            "source": self.source,
            "posted_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        self._pending.append(entry)
        # Guards against the same Ref ID appearing twice inside one statement.
        self.seen.setdefault(norm_key(ref), entry)

    @property
    def pending(self) -> list[dict[str, Any]]:
        return list(self._pending)

    def save(self) -> int:
        """Append this run's Ref IDs to the ledger. Returns how many."""
        if not self.enabled or not self._pending:
            return 0
        self.path.parent.mkdir(parents=True, exist_ok=True)
        exists = self.path.exists()
        with self.path.open("a", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=POSTED_COLUMNS)
            if not exists:
                writer.writeheader()
            writer.writerows(self._pending)
        count = len(self._pending)
        self._pending.clear()
        return count


def read_posted_refs(path: Path) -> list[dict]:
    """Every Ref ID on record, newest first."""
    if not path.exists():
        return []
    try:
        with path.open(newline="", encoding="utf-8") as handle:
            rows = list(csv.DictReader(handle))
    except OSError:
        return []
    rows.reverse()
    return rows


def import_posted_refs(path: Path, refs: list[str], note: str = "") -> tuple[int, int]:
    """Add Ref IDs that were imported before this tool existed.

    Returns (added, already_known). Only the Ref ID matters -- everything else
    is recorded as an opening entry, because there is no journal behind these.
    """
    known = {norm_key(r.get("ref_id", "")) for r in read_posted_refs(path)}
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    fresh, seen = [], 0
    for ref in refs:
        ref = clean_text(ref)
        key = norm_key(ref)
        if not key:
            continue
        if key in known:
            seen += 1
            continue
        known.add(key)
        fresh.append({"ref_id": ref, "transaction_id": "", "date": "",
                      "transaction_type": "", "amount_usd": "",
                      "document_number": "", "source": note or "loaded manually",
                      "posted_at": stamp})
    if fresh:
        path.parent.mkdir(parents=True, exist_ok=True)
        exists = path.exists()
        with path.open("a", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=POSTED_COLUMNS)
            if not exists:
                writer.writeheader()
            writer.writerows(fresh)
    return len(fresh), seen


def update_posted_ref(path: Path, old_ref: str, new_ref: str) -> None:
    """Correct one Ref ID in place."""
    old_ref, new_ref = clean_text(old_ref), clean_text(new_ref)
    if not new_ref:
        raise ValueError("The new Ref ID cannot be blank")
    rows = list(reversed(read_posted_refs(path)))
    if not rows:
        raise ValueError("There are no Ref IDs on record")
    keys = {norm_key(r.get("ref_id", "")) for r in rows}
    if norm_key(new_ref) != norm_key(old_ref) and norm_key(new_ref) in keys:
        raise ValueError(f"Ref ID '{new_ref}' is already on record")
    hit = False
    for r in rows:
        if norm_key(r.get("ref_id", "")) == norm_key(old_ref):
            r["ref_id"] = new_ref
            hit = True
    if not hit:
        raise ValueError(f"Ref ID '{old_ref}' is not on record")
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=POSTED_COLUMNS)
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, "") for k in POSTED_COLUMNS})


class DocumentNumberer:
    """Issues document numbers like 26-27/LLP/Jul/001 and 26-27/V/RE-U/001.

    One counter per rendered prefix, so the JE series (which carries the month)
    restarts each month while the Sales series runs on through the year -- both
    falling out of the format string rather than being special-cased.

    Numbers issued are persisted to a registry file and the counters resume from
    it, so a second statement -- a different client, a later month, a re-run --
    never reuses a number already posted to the books.  Nothing is written until
    `save()` is called, so a run that fails partway consumes no numbers.
    """

    def __init__(self, series: dict[str, str] | None = None, *,
                 registry_path: Path | None = None, source: str = "",
                 reset: bool = False) -> None:
        self.series = dict(DEFAULT_DOC_SERIES)
        if series:
            self.series.update(series)
        self.registry_path = registry_path
        self.source = source
        self._counters: dict[str, int] = {}
        self._issued: list[dict[str, Any]] = []
        self.resumed_from: dict[str, int] = {}
        self.reset = reset

        if registry_path and not reset:
            self._load_registry(registry_path)

    def _load_registry(self, path: Path) -> None:
        """Resume each prefix's counter from the highest number already issued.

        When there is no registry yet -- a fresh machine, or a hosted instance
        whose disk was wiped -- fall back to the seed file beside it. The seed is
        committed to the repo and records the opening position, so a deployment
        starts from the agreed number instead of 001.
        """
        if not path.exists():
            seed = path.with_name(DOC_REGISTRY_SEED_NAME)
            if seed.exists():
                path = seed
            else:
                return
        try:
            with path.open(newline="", encoding="utf-8") as handle:
                for row in csv.DictReader(handle):
                    prefix = (row.get("prefix") or "").strip()
                    try:
                        seq = int((row.get("seq") or "0").strip())
                    except ValueError:
                        continue
                    if prefix and seq > self._counters.get(prefix, 0):
                        self._counters[prefix] = seq
        except OSError:
            # An unreadable registry must not stop a run, but it MUST NOT be
            # silently treated as empty either -- that would reissue numbers.
            raise ValueError(
                f"Document registry {path} exists but could not be read; refusing to "
                f"run rather than risk reusing document numbers")
        self.resumed_from = dict(self._counters)

    def _format(self, kind: str, day, seq: int) -> tuple[str, str]:
        """Render one number, and the prefix its counter is keyed on."""
        template = self.series.get(kind, DEFAULT_DOC_SERIES[VOUCHER_JE])
        fields = {"fy": financial_year(day), "mon": day.strftime("%b"),
                  "mm": f"{day.month:02d}", "yyyy": f"{day.year}"}
        prefix = template.replace("{seq:03d}", "").replace("{seq}", "").format(**fields)
        return template.format(seq=seq, **fields), prefix

    def peek(self, kind: str, day) -> tuple[str, str | None]:
        """(next number, last used) for a series -- without consuming anything.

        The sidebar needs to show where numbering stands; asking for the next
        number should never be what advances it.
        """
        _, prefix = self._format(kind, day, 1)
        used = self._counters.get(prefix, 0)
        nxt, _ = self._format(kind, day, used + 1)
        last = self._format(kind, day, used)[0] if used else None
        return nxt, last

    def next(self, kind: str, day) -> str:
        template = self.series.get(kind, DEFAULT_DOC_SERIES[VOUCHER_JE])
        fields = {"fy": financial_year(day), "mon": day.strftime("%b"),
                  "mm": f"{day.month:02d}", "yyyy": f"{day.year}"}
        prefix = template.replace("{seq:03d}", "").replace("{seq}", "").format(**fields)
        seq = self._counters.get(prefix, 0) + 1
        self._counters[prefix] = seq
        number = template.format(seq=seq, **fields)
        self._issued.append({
            "document_number": number, "type": kind, "date": day.isoformat(),
            "prefix": prefix, "seq": seq,
            "issued_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "source": self.source,
        })
        return number

    @property
    def issued(self) -> list[dict[str, Any]]:
        return list(self._issued)

    def save(self) -> int:
        """Append this run's numbers to the registry. Returns how many.

        A reset rewrites the file rather than appending, so the registry never
        holds the same number twice -- every row in it is a number that is spoken
        for, and that invariant is what makes it safe to resume from.
        """
        if not self.registry_path or not self._issued:
            return 0
        self.registry_path.parent.mkdir(parents=True, exist_ok=True)
        mode = "w" if self.reset else "a"
        write_header = self.reset or not self.registry_path.exists()
        with self.registry_path.open(mode, newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=REGISTRY_COLUMNS)
            if write_header:
                writer.writeheader()
            writer.writerows(self._issued)
        return len(self._issued)


def narration(row: dict[str, Any], treatment_nature: str) -> str:
    """Human-readable line explaining the voucher, for the accountant."""
    parts: list[str] = [f"Being {treatment_nature}"]

    detail = clean_text(row.get("desc1")) or clean_text(row.get("summary"))
    extra = clean_text(row.get("desc2"))
    if detail:
        parts.append(detail)
    if extra and extra != detail:
        parts.append(extra)

    refs = []
    if clean_text(row.get("transaction_id")):
        refs.append(f"Txn {clean_text(row['transaction_id'])}")
    if clean_text(row.get("ref_id")) and clean_text(row.get("ref_id")) != clean_text(row.get("transaction_id")):
        refs.append(f"Ref {clean_text(row['ref_id'])}")
    if refs:
        parts.append(" / ".join(refs))

    if clean_text(row.get("client_team")):
        parts.append(f"Client: {clean_text(row['client_team'])}")
    if clean_text(row.get("account_name")):
        parts.append(f"A/c: {clean_text(row['account_name'])}")

    return " | ".join(parts)


def statement_dates(statement: pd.DataFrame) -> list:
    """Every parseable transaction date, so FX can be fetched in one batch."""
    out = []
    for _, row in statement.iterrows():
        parsed = pd.to_datetime(clean_text(row.get("date")), errors="coerce")
        if not pd.isna(parsed):
            out.append(parsed.date())
    return out


class Reporter:
    """Collects exceptions and skips so nothing fails silently.

    Exceptions carry a severity: ERROR means the row was not posted and needs a
    human, INFO means the tool handled it but wants the decision on record (an
    auto-resolved duplicate, for instance).  Only ERRORs affect the exit code.
    """

    def __init__(self) -> None:
        self.exceptions: list[dict[str, Any]] = []
        self.skipped: list[dict[str, Any]] = []

    def exception(self, row_no: Any, category: str, detail: str,
                  severity: str = "ERROR", **extra: Any) -> None:
        self.exceptions.append({"Statement Row": row_no, "Severity": severity,
                                "Category": category, "Detail": detail, **extra})

    @property
    def errors(self) -> list[dict[str, Any]]:
        return [e for e in self.exceptions if e["Severity"] == "ERROR"]

    def skip(self, row_no: Any, reason: str, **extra: Any) -> None:
        self.skipped.append({"Statement Row": row_no, "Reason": reason, **extra})


def build_journal(statement: pd.DataFrame, mapping: Mapping, *,
                  igst_rate: Decimal, currency: str, fx_provider,
                  income_mode: str, reporter: Reporter,
                  doc_series: dict[str, str] | None = None,
                  numberer: "DocumentNumberer | None" = None,
                  entity: str = "",
                  posted: "PostedLedger | None" = None) -> pd.DataFrame:
    """Walk the statement and emit journal lines.

    `fx_provider` resolves one rate per transaction date (see rbi_fx.py); it is
    primed with every date up front so the whole run costs one request per month
    rather than one per row.

    Every voucher is checked for debit==credit before it is accepted; failures
    are still emitted (so the accountant can see the context) but are recorded
    as exceptions and make the process exit non-zero.
    """
    lines: list[dict[str, Any]] = []
    serial = 0
    if numberer is None:
        numberer = DocumentNumberer(doc_series)

    if currency != "USD":
        fx_provider.prepare(statement_dates(statement))

    # Post in date order. The Upwork export is newest-first, but a journal is
    # read and imported oldest-first, and document numbers must ascend with the
    # dates they carry. `kind="stable"` keeps same-day rows in statement order.
    statement = statement.assign(
        _sort_date=pd.to_datetime(statement["date"].map(clean_text), errors="coerce")
    ).sort_values("_sort_date", kind="stable").drop(columns="_sort_date")

    for idx, raw in statement.iterrows():
        row = raw.to_dict()
        row_no = idx + 2  # +2 = 1-based, past the header, matching Excel

        # --- parse and validate the row's own fields -------------------------
        txn_type = clean_text(row.get("txn_type"))
        if not txn_type:
            reporter.skip(row_no, "Blank transaction type")
            continue

        treatment = mapping.treatments.get(norm_key(txn_type))
        if treatment is None:
            reporter.exception(row_no, "UNMAPPED_TYPE",
                               f"Transaction type '{txn_type}' is not in Table A",
                               **{"Transaction Type": txn_type})
            continue

        if treatment.kind == "SKIP":
            reporter.skip(row_no, f"'{txn_type}' is marked NO ENTRY in Table A",
                          **{"Transaction Type": txn_type,
                             "Amount USD": clean_text(row.get("amount_usd"))})
            continue

        # Statements overlap by design, and an old one gets re-uploaded by
        # accident. Ref ID identifies the Upwork transaction, so anything
        # already journalised is skipped rather than posted twice.
        ref_id = clean_text(row.get("ref_id")) or clean_text(row.get("transaction_id"))
        if posted is not None:
            earlier = posted.already_posted(ref_id)
            if earlier:
                reporter.skip(
                    row_no,
                    f"Already imported on {earlier.get('posted_at', '?')[:10]} "
                    f"as {earlier.get('document_number', '?')}",
                    **{"Transaction Type": txn_type, "Ref ID": ref_id,
                       "Amount USD": clean_text(row.get("amount_usd"))})
                continue

        amount_usd = to_decimal(row.get("amount_usd"))
        if amount_usd is None:
            reporter.exception(row_no, "BAD_AMOUNT",
                               f"Amount '{clean_text(row.get('amount_usd'))}' is not a number",
                               **{"Transaction Type": txn_type})
            continue
        if amount_usd == 0:
            reporter.skip(row_no, "Zero amount", **{"Transaction Type": txn_type})
            continue

        date = pd.to_datetime(clean_text(row.get("date")), errors="coerce")
        if pd.isna(date):
            reporter.exception(row_no, "BAD_DATE",
                               f"Unparseable date '{clean_text(row.get('date'))}'",
                               **{"Transaction Type": txn_type})
            continue

        account = clean_text(row.get("account_name"))
        account_key = norm_key(account)
        if account_key in mapping.ambiguous_wallets:
            candidates = ", ".join(mapping.ambiguous_wallets[account_key])
            reporter.exception(
                row_no, "AMBIGUOUS_ACCOUNT",
                f"Account '{account}' maps to conflicting GLs in Table B ({candidates}) "
                f"and no one of them clearly matches the name -- resolve the duplicate "
                f"before posting",
                **{"Transaction Type": txn_type})
            continue
        wallet_gl = mapping.wallets.get(account_key, "")
        if not wallet_gl:
            reporter.exception(row_no, "UNMAPPED_ACCOUNT",
                               f"Account name '{account}' is not in Table B",
                               **{"Transaction Type": txn_type})
            continue

        # --- currency --------------------------------------------------------
        rate = Decimal("1")
        rate_date_text = ""
        rate_source = "-"
        if currency != "USD":
            decision = fx_provider.rate_for(date.date())
            if not decision.ok:
                reporter.exception(row_no, "MISSING_FX_RATE",
                                   f"No {currency} rate for {date.date().isoformat()}: "
                                   f"{decision.note}",
                                   **{"Transaction Type": txn_type})
                continue
            rate = decision.rate
            rate_date_text = decision.rate_date.isoformat() if decision.rate_date else ""
            rate_source = decision.source_label

        base = money(abs(amount_usd) * rate)
        igst = money(base * igst_rate) if treatment.kind == "EXPENSE_RCM" else Decimal("0")

        ctx = Context(wallet_gl=wallet_gl,
                      client=clean_text(row.get("client_team")),
                      base=base, igst=igst)

        # --- build, verify and emit -----------------------------------------
        text = narration(row, treatment.nature)
        # Cost centre: the master database decides, because the code the books
        # use ("TPT-Badshah") is rarely the name on the statement. Only when an
        # account has no cost centre there do we fall back to the Freelancer
        # column, and then to the wallet owner for rows with no freelancer at all
        # (Connects and Subscription are billed to the account, not a contract).
        cost_center = (mapping.cost_centers.get(account_key)
                       or clean_text(row.get("freelancer"))
                       or account)
        # Subsidiary: the entity this voucher belongs to. A rule-level override
        # wins if one is set, then the account's own entity, then the run default
        # -- so a shared statement can still post to more than one entity.
        subsidiary = mapping.subsidiaries.get(account_key) or entity

        first_doc = ""
        for voucher in build_legs(treatment, ctx, income_mode):
            doc_no = numberer.next(voucher.kind, date.date())
            if not first_doc:
                first_doc = doc_no

            debits = sum((leg.amount for leg in voucher.legs if leg.side == "Dr"),
                         Decimal("0"))
            credits = sum((leg.amount for leg in voucher.legs if leg.side == "Cr"),
                          Decimal("0"))
            if abs(debits - credits) > BALANCE_TOLERANCE:
                reporter.exception(
                    row_no, "UNBALANCED_VOUCHER",
                    f"{doc_no} debits {debits} != credits {credits}",
                    **{"Transaction Type": txn_type})

            for leg in voucher.legs:
                serial += 1
                lines.append({
                    "Sr. No.": serial,
                    "Document Number": doc_no,
                    "Voucher Type": voucher.kind,
                    "Date": date.date(),
                    "Period": date.strftime("%b-%Y"),
                    "Transaction ID": clean_text(row.get("transaction_id")),
                    "Ref ID": clean_text(row.get("ref_id")),
                    "Transaction Type": treatment.nature,
                    "Subsidiary": subsidiary,
                    "Account Name": account,
                    "Cost Center": cost_center,
                    "Client Team": ctx.client,
                    "Ledger": leg.ledger,
                    "Dr/Cr": leg.side,
                    "Debit": float(leg.amount) if leg.side == "Dr" else 0.0,
                    "Credit": float(leg.amount) if leg.side == "Cr" else 0.0,
                    "Currency": currency,
                    "Amount USD": float(abs(amount_usd)),
                    "FX Rate": float(rate),
                    "Rate Date": rate_date_text,
                    "Rate Source": rate_source,
                    "Narration": text,
                })

        if posted is not None:
            posted.record(ref_id, row, first_doc, treatment.nature, abs(amount_usd))

    return pd.DataFrame(lines, columns=JOURNAL_COLUMNS)


# --------------------------------------------------------------------------- #
# Reporting
# --------------------------------------------------------------------------- #

def build_import_frame(journal: pd.DataFrame, *, entity: str, currency: str,
                       ie_flag: str = "I") -> pd.DataFrame:
    """Reshape the journal into the accounting system's import layout.

    Amounts are already in the booking currency, so `Currency` is that currency
    and `Exchange Rate` is 1 -- the USD original, the RBI rate and the date it
    came from stay on the Journal and FX Audit sheets rather than being squeezed
    into a file whose columns the target system defines.

    `Amount in base currency` / `Amount in INR` are filled on Sales vouchers and
    left blank on JE vouchers, matching how the template is kept by hand.
    """
    if journal.empty:
        return pd.DataFrame(columns=IMPORT_COLUMNS)

    rows = []
    for _, line in journal.iterrows():
        amount = round(float(line["Debit"] if line["Dr/Cr"] == "Dr" else line["Credit"]), 2)
        is_sale = line["Voucher Type"] == VOUCHER_SALES
        rows.append({
            # Per-line, so one statement can span entities. Falls back to the
            # run-wide entity when the master says nothing.
            "Subsidiary": line.get("Subsidiary") or entity,
            "Transaction date": line["Date"].strftime("%d/%m/%Y"),
            "Period": line["Period"],
            "Document Number": line["Document Number"],
            "Type": line["Voucher Type"],
            "Ledger Name": line["Ledger"],
            "Amount": amount,
            "Debit": round(float(line["Debit"]), 2),
            "Credit": round(float(line["Credit"]), 2),
            "Amount in base currency": amount if is_sale else "",
            "Amount in INR": amount if is_sale else "",
            "Currency": currency,
            "Exchange Rate": 1,
            "I/E": ie_flag,
            "Cost center Name": line["Cost Center"],
            "Narration": line["Narration"],
            "Ref ID": line.get("Ref ID", ""),
        })
    return pd.DataFrame(rows, columns=IMPORT_COLUMNS)


def build_reconciliation(journal: pd.DataFrame, currency: str) -> pd.DataFrame:
    """Per-type debit/credit/voucher counts plus a grand total row."""
    if journal.empty:
        return pd.DataFrame(columns=["Transaction Type", "Vouchers", "Journal Lines",
                                     f"Total Debit ({currency})",
                                     f"Total Credit ({currency})", "Difference"])

    grouped = journal.groupby("Transaction Type", as_index=False).agg(
        Vouchers=("Document Number", "nunique"),
        **{"Journal Lines": ("Sr. No.", "count")},
        Debit=("Debit", "sum"),
        Credit=("Credit", "sum"),
    )
    grouped["Difference"] = (grouped["Debit"] - grouped["Credit"]).round(2)
    grouped = grouped.rename(columns={"Debit": f"Total Debit ({currency})",
                                      "Credit": f"Total Credit ({currency})"})

    total = {
        "Transaction Type": "TOTAL",
        "Vouchers": journal["Document Number"].nunique(),
        "Journal Lines": len(journal),
        f"Total Debit ({currency})": round(journal["Debit"].sum(), 2),
        f"Total Credit ({currency})": round(journal["Credit"].sum(), 2),
        "Difference": round(journal["Debit"].sum() - journal["Credit"].sum(), 2),
    }
    return pd.concat([grouped, pd.DataFrame([total])], ignore_index=True)


def autofit(writer: pd.ExcelWriter, sheet_name: str, frame: pd.DataFrame) -> None:
    """Widen columns so the workbook is readable without manual resizing."""
    worksheet = writer.sheets[sheet_name]
    for i, column in enumerate(frame.columns, start=1):
        longest = max([len(str(column))] +
                      [len(str(v)) for v in frame[column].head(200).tolist()] or [0])
        worksheet.column_dimensions[
            worksheet.cell(row=1, column=i).column_letter].width = min(longest + 2, 60)


def write_outputs(out_path: Path, journal: pd.DataFrame, reconciliation: pd.DataFrame,
                  exceptions: pd.DataFrame, skipped: pd.DataFrame,
                  fx_audit: pd.DataFrame | None = None,
                  import_frame: pd.DataFrame | None = None) -> Path:
    """Write the workbook plus the import CSV.

    The CSV is the *import* layout -- that file goes straight into the
    accounting system, so it carries exactly the template's columns and nothing
    else. Everything needed to audit a figure lives in the workbook alongside.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if fx_audit is None:
        fx_audit = pd.DataFrame()
    if import_frame is None:
        import_frame = pd.DataFrame(columns=IMPORT_COLUMNS)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, frame in (("Import", import_frame),
                            ("Journal", journal),
                            ("Reconciliation", reconciliation),
                            ("FX Audit", fx_audit),
                            ("Exceptions", exceptions),
                            ("Skipped", skipped)):
            # An empty sheet still gets written, so the workbook shape is
            # identical run to run and downstream tooling never guesses.
            frame.to_excel(writer, sheet_name=name, index=False)
            autofit(writer, name, frame)

    csv_path = out_path.with_suffix(".csv")
    import_frame.to_csv(csv_path, index=False, encoding="utf-8-sig")
    return csv_path


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def discover_in_folder(folder: Path) -> tuple[Path, Path | None]:
    """Identify the statement and mapping files inside a folder by content.

    Avoids depending on filenames: whichever file has a 'Transaction type'
    header is the statement, whichever has 'Nature' is the mapping.  The mapping
    is optional -- when the folder holds only a statement, the caller falls back
    to the master database.
    """
    statement = mapping = None
    candidates: Iterable[Path] = sorted(
        p for p in folder.iterdir()
        if p.suffix.lower() in (".csv", ".xlsx", ".xlsm", ".xls") and not p.name.startswith("~$")
    )
    for path in candidates:
        try:
            grid = read_raw(path, excel_sheet_names(path)[0] if excel_sheet_names(path) else None)
        except Exception:
            continue
        head = grid[:20]
        if statement is None and find_header(head, "Transaction type"):
            statement = path
        elif mapping is None and find_header(head, "Nature"):
            mapping = path

    if statement is None:
        raise ValueError(
            f"No Upwork statement found in {folder}. Pass --statement explicitly.")
    return statement, mapping


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert an Upwork account statement into journal entries.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("--statement", type=Path, help="Sheet1: raw Upwork statement (.csv/.xlsx)")
    parser.add_argument("--mapping", type=Path, default=None,
                        help=f"Accounting treatment map. Defaults to the master database "
                             f"at {DEFAULT_MASTER_MAPPING}")
    parser.add_argument("--folder", type=Path,
                        help="Folder holding both files; they are identified by their headers")
    parser.add_argument("--out", type=Path, default=Path("journal.xlsx"),
                        help="Output .xlsx path (a .csv of the journal is written alongside)")
    parser.add_argument("--currency", choices=("INR", "USD"), default="INR",
                        help="Book the journal in this currency")
    parser.add_argument("--fx-source", choices=("rbi", "manual"), default="rbi",
                        help="rbi: fetch the RBI reference rate per transaction date. "
                             "manual: use --fx-rate for everything")
    parser.add_argument("--fx-rate", type=str, default=None,
                        help="Flat USD->INR rate. With --fx-source rbi this is only the "
                             "safety net used when RBI and the cache have nothing")
    parser.add_argument("--fx-rates", type=Path, default=None,
                        help="File with Period/Date + Rate columns; wins over RBI for "
                             "the dates it covers (YYYY-MM-DD, DD/MM/YYYY or YYYY-MM)")
    parser.add_argument("--fx-cache", type=Path, default=Path("fx_cache.csv"),
                        help="Local cache of fetched RBI rates; only uncached dates are fetched")
    parser.add_argument("--fx-offline", action="store_true",
                        help="Never contact RBI; use the cache, then --fx-rate")
    parser.add_argument("--fx-max-fallback-days", type=int, default=7,
                        help="How far back to walk for the last published rate")
    parser.add_argument("--fx-stale-days", type=int, default=3,
                        help="Warn when the fallback reached back further than this")
    parser.add_argument("--fx-delay", type=float, default=1.0,
                        help="Minimum seconds between requests to RBI")
    parser.add_argument("--fx-timeout", type=float, default=30.0,
                        help="Per-request timeout in seconds")
    parser.add_argument("--fx-retries", type=int, default=3,
                        help="Attempts per request before giving up (exponential backoff)")
    parser.add_argument("--igst-rate", type=str, default="18",
                        help="Reverse-charge IGST rate in percent")
    parser.add_argument("--entity", default="Verve Advisory LLP",
                        help="Subsidiary name written to the import file")
    parser.add_argument("--ie-flag", default="I",
                        help="Value for the import file's I/E column")
    parser.add_argument("--doc-series-je", default=DEFAULT_DOC_SERIES[VOUCHER_JE],
                        help="Document-number format for JE vouchers")
    parser.add_argument("--doc-series-sales", default=DEFAULT_DOC_SERIES[VOUCHER_SALES],
                        help="Document-number format for Sales vouchers")
    parser.add_argument("--doc-registry", type=Path, default=DEFAULT_DOC_REGISTRY,
                        help="Record of document numbers already issued; counters "
                             "resume from it so numbers are never reused")
    parser.add_argument("--reset-doc-numbers", action="store_true",
                        help="Ignore the registry and start each series at 001 again")
    parser.add_argument("--posted-ledger", type=Path, default=DEFAULT_POSTED_LEDGER,
                        help="Record of Ref IDs already journalised; matching rows "
                             "in a later statement are skipped as duplicates")
    parser.add_argument("--ignore-posted", action="store_true",
                        help="Post every row even if its Ref ID was imported before")
    parser.add_argument("--income-mode", choices=("two-entry", "combined"), default="two-entry",
                        help="two-entry: Sales voucher (Dr Client/Cr Revenue) then a JE "
                             "(Dr Wallet/Cr Client). combined: one Dr Wallet/Cr Revenue")
    parser.add_argument("--on-duplicate", choices=("resolve", "fail"), default="resolve",
                        help="Table B account listed twice with different GLs: 'resolve' "
                             "picks the GL matching the name, 'fail' refuses those rows")
    parser.add_argument("--strict", action="store_true",
                        help="Exit non-zero if ANY error is recorded, not just imbalances")
    return parser.parse_args(argv)


def build_fx_provider(args, mapping: Mapping, manual_rate: Decimal | None, log=print):
    """Assemble the object that answers 'what rate applies on this date?'.

    USD output and --fx-source manual both collapse to a flat rate, so they use
    the trivial provider and never touch the network.  Returns None when the
    options can't produce any rate at all, which the caller treats as fatal.
    """
    if args.currency == "USD":
        return FxProvider(rate=Decimal("1"), source="-")

    if args.fx_source == "manual":
        if manual_rate is None:
            print("error: --fx-source manual needs --fx-rate", file=sys.stderr)
            return None
        return FxProvider(rate=manual_rate, source="MANUAL")

    # A user FX file wins for the dates it covers; a Period/Rate table inside the
    # mapping sheet is treated the same way, so old mapping files keep working.
    file_rates: dict = {}
    for period, rate in mapping.fx_rates.items():
        file_rates.update(_expand_period(period, rate))
    if args.fx_rates:
        file_rates.update(load_fx_file_by_date(args.fx_rates))

    return RbiFxProvider(
        cache_path=args.fx_cache, file_rates=file_rates, manual_rate=manual_rate,
        max_fallback_days=args.fx_max_fallback_days,
        stale_after_days=args.fx_stale_days, offline=args.fx_offline,
        timeout=args.fx_timeout, delay=args.fx_delay, retries=args.fx_retries,
        log=log)


def _expand_period(period: str, rate: Decimal) -> dict:
    """Turn a legacy 'YYYY-MM' (or '*') mapping-sheet entry into per-date rates.

    '*' can't be expanded to specific dates, so it is deliberately dropped here
    and stays available only as the manual safety net.
    """
    from datetime import date as _date, timedelta as _timedelta

    try:
        year, month = (int(part) for part in period.split("-")[:2])
        first = _date(year, month, 1)
    except (ValueError, TypeError):
        return {}
    last = (_date(year + 1, 1, 1) if month == 12 else _date(year, month + 1, 1)) \
        - _timedelta(days=1)
    out, cursor = {}, first
    while cursor <= last:
        out[cursor] = rate
        cursor += _timedelta(days=1)
    return out


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    # --- resolve inputs ------------------------------------------------------
    try:
        # The mapping falls back to the master database, so a routine run only
        # needs to name the statement.
        if args.folder:
            statement_path, found_mapping = discover_in_folder(args.folder)
            mapping_path = args.mapping or found_mapping or DEFAULT_MASTER_MAPPING
        elif args.statement:
            statement_path = args.statement
            mapping_path = args.mapping or DEFAULT_MASTER_MAPPING
        else:
            print("error: pass --statement (or --folder)", file=sys.stderr)
            return 2

        for path in (statement_path, mapping_path):
            if not path.exists():
                print(f"error: file not found: {path}", file=sys.stderr)
                return 2

        igst_rate = to_decimal(args.igst_rate)
        if igst_rate is None or igst_rate < 0:
            print(f"error: invalid --igst-rate '{args.igst_rate}'", file=sys.stderr)
            return 2
        igst_rate = igst_rate / Decimal("100")  # percent -> fraction

        mapping = load_mapping(mapping_path, on_duplicate=args.on_duplicate)
        statement = load_statement(statement_path)

        manual_rate = None
        if args.fx_rate:
            manual_rate = to_decimal(args.fx_rate)
            if manual_rate is None or manual_rate <= 0:
                print(f"error: invalid --fx-rate '{args.fx_rate}'", file=sys.stderr)
                return 2

        fx_provider = build_fx_provider(args, mapping, manual_rate)
        if fx_provider is None:
            return 2
    except Exception as exc:  # input problems are fatal and worth a clear message
        print(f"error: {exc}", file=sys.stderr)
        return 2

    # --- build ---------------------------------------------------------------
    reporter = Reporter()

    # Table B duplicates are reported once up front, independently of whether any
    # statement row happens to use them: auto-resolved ones as INFO (so the choice
    # is on the record), genuinely ambiguous ones as ERROR.
    for key, (chosen, rejected) in mapping.resolved_wallets.items():
        reporter.exception("-", "DUPLICATE_RESOLVED",
                           f"Table B lists '{mapping.wallet_display.get(key, key)}' more "
                           f"than once; used '{chosen}' as the name matches, ignored "
                           f"{', '.join(repr(gl) for gl in rejected)}",
                           severity="INFO")
    for key, candidates in mapping.ambiguous_wallets.items():
        reporter.exception("-", "AMBIGUOUS_ACCOUNT",
                           f"Table B maps '{mapping.wallet_display.get(key, key)}' to "
                           f"conflicting GLs: {', '.join(candidates)}")

    try:
        numberer = DocumentNumberer(
            {VOUCHER_JE: args.doc_series_je, VOUCHER_SALES: args.doc_series_sales},
            registry_path=args.doc_registry, source=statement_path.name,
            reset=args.reset_doc_numbers)
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    try:
        posted = PostedLedger(args.posted_ledger, source=statement_path.name,
                              enabled=not args.ignore_posted)
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    journal = build_journal(
        statement, mapping,
        igst_rate=igst_rate, currency=args.currency, fx_provider=fx_provider,
        income_mode=args.income_mode, reporter=reporter, numberer=numberer,
        entity=args.entity, posted=posted)

    # Every stale-rate decision goes on the record as an INFO note, so the
    # accountant sees which dates were priced off an older day's rate.
    for warning in fx_provider.warnings:
        reporter.exception("-", "FX_FALLBACK", warning, severity="INFO")

    import_frame = build_import_frame(journal, entity=args.entity,
                                      currency=args.currency, ie_flag=args.ie_flag)
    reconciliation = build_reconciliation(journal, args.currency)
    fx_audit = pd.DataFrame(fx_provider.audit_rows()
                            if hasattr(fx_provider, "audit_rows") else [])
    exceptions = pd.DataFrame(reporter.exceptions)
    skipped = pd.DataFrame(reporter.skipped)

    csv_path = write_outputs(args.out, journal, reconciliation, exceptions, skipped,
                             fx_audit=fx_audit, import_frame=import_frame)

    # Only commit the numbers -- and the Ref IDs -- once the output is written.
    recorded = numberer.save()
    posted_now = posted.save()

    # --- summarise to the console -------------------------------------------
    total_debit = round(journal["Debit"].sum(), 2) if not journal.empty else 0.0
    total_credit = round(journal["Credit"].sum(), 2) if not journal.empty else 0.0
    imbalances = [e for e in reporter.exceptions if e["Category"] == "UNBALANCED_VOUCHER"]
    errors = reporter.errors
    notes = len(reporter.exceptions) - len(errors)

    print(f"Statement      : {statement_path}")
    print(f"Mapping        : {mapping_path}")
    print(f"Rows read      : {len(statement)}")
    print(f"Vouchers       : {journal['Document Number'].nunique() if not journal.empty else 0}")
    print(f"Journal lines  : {len(journal)}")
    print(f"Total debits   : {total_debit:,.2f} {args.currency}")
    print(f"Total credits  : {total_credit:,.2f} {args.currency}")
    print(f"Difference     : {total_debit - total_credit:,.2f}")
    duplicates = sum(1 for s in reporter.skipped
                     if str(s.get("Reason", "")).startswith("Already imported"))
    print(f"Skipped rows   : {len(skipped)}"
          + (f"  ({duplicates} already imported)" if duplicates else ""))
    if posted_now:
        print(f"Ref IDs logged : {posted_now} in {args.posted_ledger.name}")
    if not fx_audit.empty:
        sources = fx_audit["Source"].value_counts().to_dict()
        print(f"FX dates       : {len(fx_audit)} "
              f"({', '.join(f'{v}x {k}' for k, v in sources.items())})")
    if recorded:
        resumed = ", ".join(f"{p}{s}" for p, s in sorted(numberer.resumed_from.items()))
        print(f"Doc numbers    : {recorded} recorded in {args.doc_registry.name}"
              + (f" (resumed after {resumed})" if resumed else " (first run)"))
    print(f"Errors         : {len(errors)}")
    print(f"Notes (INFO)   : {notes}")
    print(f"Written        : {args.out}")
    print(f"                 {csv_path}")

    for note in reporter.exceptions:
        if note["Severity"] == "INFO":
            print(f"  note: {note['Detail']}")

    if imbalances:
        print(f"\nFAILED: {len(imbalances)} voucher(s) do not balance -- see the "
              f"Exceptions sheet.", file=sys.stderr)
        return 1
    if args.strict and errors:
        print(f"\nFAILED (--strict): {len(errors)} error(s) recorded.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
