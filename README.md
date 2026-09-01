# Upwork Statement → Accounting Journal

Converts a raw Upwork account statement (Sheet1) into balanced double-entry journal
entries, using an externally-maintained treatment map (Sheet2). Outputs an Excel
workbook (Journal / Reconciliation / Exceptions / Skipped) plus a flat CSV of the journal.

**Nothing is hard-coded.** Ledger names, transaction types, wallet GLs and FX rates all
come from the mapping file — add a freelancer or a new transaction type by editing
Sheet2, not the script.

---

**Live interface preview:** https://amolbibwe-web.github.io/vervi-upwork/ — the real UI
with sample data, no install needed.

## Install

```bash
git clone https://github.com/amolbibwe-web/vervi-upwork.git
cd vervi-upwork
pip install -r requirements.txt
cp demo/mapping_master.csv master/mapping_master.csv   # then edit in your real accounts
python webapp.py                                       # http://127.0.0.1:5000
```

Requires Python 3.10+ (tested on 3.12.10 with pandas 3.0.5 / openpyxl 3.1.5).
Sign in with `admin` / `admin`, or set `VERVI_USER` / `VERVI_PASS` first.

### What is not in this repo

The repo is public so GitHub Pages can serve the interface for free, so **no live
business data is committed**. `.gitignore` keeps out `data/`, the real
`master/mapping_master.csv`, `master/doc_registry.csv`, `fx_cache.csv` and `out/`.
What ships instead is `demo/` — a fictional statement and master with the same shape.

That means each person clones the repo and points it at their own master database and
statements. Nothing about your books leaves your machine.

## Hosting

| What you want | Where | Runs Python? |
|---|---|---|
| Show the interface to anyone | **GitHub Pages** (`/docs`) | No — static preview only |
| A working converter on a URL | **Render** — `render.yaml` is ready | Yes |
| A quick private link to your local app | `.\share.ps1` (Cloudflare tunnel) | Yes, off your machine |

For Render: *New → Blueprint →* point at this repo. **Set `VERVI_PASS` in the dashboard
before sharing the URL.** Note the free tier has an ephemeral filesystem, so
`fx_cache.csv` and the document registry reset when the instance restarts — mount a disk
or commit a seeded registry if continuous numbering matters in the cloud.

## Run

```bash
python upwork_to_journal.py --statement Sheet1.xlsx --out journal.xlsx
```

The statement is the only file you supply. The mapping comes from the **master database**
at [master/mapping_master.csv](master/mapping_master.csv), and no FX rate is needed —
each transaction converts at the **RBI reference rate for its own date**, fetched
automatically and cached (see *Exchange rates* below).

With the bundled sample statement:

```bash
python upwork_to_journal.py --statement data/Sheet1.csv --out out/journal.xlsx
```

Or point at a folder and let it identify the two files by their headers:

```bash
python upwork_to_journal.py --folder data --fx-rate 87.5 --out out/journal.xlsx
```

Writes `out/journal.xlsx` **and** `out/journal.csv`.

## Vervi-Upwork — the browser UI

```bash
python webapp.py
```

Then open **http://127.0.0.1:5000** and drop in **only the Upwork statement** — drag and
drop works. Two steps: upload, generate. Everything else is already set to the standard
treatment and tucked behind *Change settings*.

The left panel carries the identity and every standing fact — company, master-database
counts, defaults in force, and where document numbering has reached — which keeps the top
of the page to a title and one line of context. On the results page it becomes the run
summary: totals, settings used, and the download buttons.

Results appear in tabs — **Import file**, Journal, Reconciliation, FX audit, Exceptions,
Skipped — with `Dr`/`Cr` and `Sales`/`JE` colour-coded and rules between vouchers.
Roboto throughout, with a dark mode that follows the system setting.

`--port 8000` to move it, `--master path.csv` to point at a different master.

The UI imports the same functions the CLI uses, so the two can't drift apart. It binds to
`127.0.0.1` only — it's a local convenience, not a service to expose.

## Options

| Flag | Default | Meaning |
|---|---|---|
| `--statement PATH` | — | The raw Upwork statement (`.csv` / `.xlsx`) |
| `--mapping PATH` | master database | Override the treatment map for a one-off run |
| `--folder PATH` | — | Folder holding the statement; identified by header content, not filename |
| `--out PATH` | `journal.xlsx` | Output workbook; the CSV is written alongside with the same stem |
| `--currency {INR,USD}` | `INR` | Book the journal in this currency |
| `--fx-source {rbi,manual}` | `rbi` | `rbi`: RBI reference rate per transaction date. `manual`: one flat rate |
| `--fx-rate N` | — | Flat USD→INR rate. Under `--fx-source rbi` this is only the **safety net** |
| `--fx-rates PATH` | — | Rate file that **wins over RBI** for the dates it covers |
| `--fx-cache PATH` | `fx_cache.csv` | Local rate cache; only uncached dates are fetched |
| `--fx-offline` | off | Never contact RBI — cache, then `--fx-rate` |
| `--fx-max-fallback-days N` | `7` | How far back to walk for the last published rate |
| `--fx-stale-days N` | `3` | Warn when the fallback reached back further than this |
| `--fx-delay N` | `1.0` | Minimum seconds between RBI requests |
| `--fx-timeout N` | `30` | Per-request timeout |
| `--fx-retries N` | `3` | Attempts per request, with exponential backoff |
| `--igst-rate N` | `18` | Reverse-charge IGST rate, in percent |
| `--income-mode {two-entry,combined}` | `two-entry` | See *Sales entries* below |
| `--entity NAME` | `Verve Advisory LLP` | Subsidiary column of the import file |
| `--ie-flag X` | `I` | I/E column of the import file |
| `--doc-series-je FMT` | `{fy}/LLP/{mon}/{seq:03d}` | Document numbers for JE vouchers |
| `--doc-series-sales FMT` | `{fy}/V/RE-U/{seq:03d}` | Document numbers for Sales vouchers |
| `--doc-registry PATH` | `master/doc_registry.csv` | Record of numbers already issued; series resume from it |
| `--reset-doc-numbers` | off | Restart every series at 001 and rewrite the registry |
| `--on-duplicate {resolve,fail}` | `resolve` | Account listed twice in Table B with different GLs — see *Duplicates* below |
| `--strict` | off | Exit non-zero if **any** error is recorded, not just an imbalance |

---

## Exchange rates

With `--currency INR` (the default) each transaction is converted at the **RBI USD/INR
reference rate for its own date**, scraped from the
[RBI Reference Rate Archive](https://rbi.org.in/scripts/ReferenceRateArchive.aspx).

The archive is an ASP.NET WebForms page, so [rbi_fx.py](rbi_fx.py) GETs it to capture
`__VIEWSTATE` / `__VIEWSTATEGENERATOR` / `__EVENTVALIDATION`, then POSTs those back with
`chkUSD=on`, `txtFromDate`/`txtToDate` (`dd/mm/yyyy`) and `btnSubmit=" GO "`, and parses
the table headed `Date | USD (INR / 1 USD)`. The form takes a **date range**, so the
whole run costs **one request per month**, not one per transaction.

### Resolution order, per date

```
--fx-rates file      wins for any date it covers
  └─> RBI / cache    exact date
      └─> RBI / cache  most recent previous published day  (--fx-max-fallback-days)
          └─> --fx-rate  manual safety net
              └─> reported as MISSING_FX_RATE — never silently mispriced
```

RBI publishes on business days only, so a weekend or bank-holiday transaction has no rate
of its own. The tool walks backwards day by day to the most recent published rate and
records **which date's rate it actually used** on every journal line and in the FX Audit
sheet. In the sample data, 12 Jul 2026 is a Sunday and correctly picks up Friday
10 Jul's rate of 95.3129.

### Caching

Fetched rates are written to `fx_cache.csv` (`date,rate,source,fetched_at`) and only
uncached dates are ever requested — a second run makes **no network calls at all**. CSV
rather than SQLite so you can open it, check a rate, or correct one by hand.

Non-publishing days are cached too, as `source=NO_RATE`, so repeat runs stop re-asking
about the same Sunday. Only *interior* gaps are recorded that way; dates after the last
published day stay unknown, so rates that aren't out yet get picked up on a later run.

### Politeness and failure

A normal browser User-Agent, a minimum 1s gap between requests, a 30s timeout, and 3
retries with exponential backoff. **A run never dies because RBI was down**: an
unreachable archive degrades to the cache, then to `--fx-rate`, and every decision is
logged. `--fx-offline` skips the network entirely.

### Rate audit

The **FX Audit** sheet gives one row per transaction date:

| Transaction Date | Rate Date Used | Rate (INR/USD) | Source | Days Back | Transactions | Stale |
|---|---|---|---|---|---|---|
| 2026-07-10 | 2026-07-10 | 95.3129 | RBI | 0 | 3 | |
| 2026-07-12 | 2026-07-10 | 95.3129 | RBI (prev-day fallback) | 2 | 1 | |

Every journal line also carries `FX Rate`, `Rate Date` and `Rate Source` alongside
`Amount USD`. Any date whose fallback reached back further than `--fx-stale-days` is
listed on the Exceptions sheet as an `FX_FALLBACK` note.

### Rate files

`--fx-rates` accepts exact dates or whole months, and wins over RBI for the dates it
covers — useful when your auditor mandates a specific rate:

```
Date,Rate            Period,Rate
2026-07-31,88.10     2026-07,87.50     <- applied to every day in July
```

A `Period`/`Rate` table inside the mapping sheet still works and is treated the same way.

### Exit codes

| Code | Meaning |
|---|---|
| `0` | Clean run (INFO notes may still be listed — check the Exceptions sheet) |
| `1` | A voucher failed the debit == credit check, or `--strict` and ERRORs exist |
| `2` | Fatal input problem: missing file, unreadable mapping, no FX rate |

Exceptions carry a **Severity**: `ERROR` means the row was not posted and needs a human;
`INFO` means the tool handled it but is putting the decision on record. Only `ERROR`s
affect the exit code, so `--strict` won't trip on a routine auto-resolved duplicate.

---

## Accounting treatment

`[Wallet]` = the GL from Table B for the row's `Account name`. `E` = the transaction
amount (absolute value, converted to the output currency). `R` = the IGST rate.

| Transaction type | Kind | Debit | Credit |
|---|---|---|---|
| Service Fee | `EXPENSE_RCM` | Upwork Charges `E`<br>RCM Input IGST `E×R` | `[Wallet]` `E`<br>RCM Output IGST `E×R` |
| Connects | `EXPENSE_RCM` | Connects charges `E`<br>RCM Input IGST `E×R` | `[Wallet]` `E`<br>RCM Output IGST `E×R` |
| Subscription | `EXPENSE_RCM` | Upwork Membership Fees `E`<br>RCM Input IGST `E×R` | `[Wallet]` `E`<br>RCM Output IGST `E×R` |
| Withdrawal Fee | `EXPENSE_RCM` | Withdrawal fees `E`<br>RCM Input IGST `E×R` | `[Wallet]` `E`<br>RCM Output IGST `E×R` |
| WHT | `SIMPLE` | TDS `E` | `[Wallet]` `E` |
| Hourly, Fixed-price | `INCOME` | see *Sales entries* below | |
| Withdrawal | `SKIP` | — no entry, logged to the Skipped sheet — | |

**RCM.** The two IGST legs are equal and opposite, so the reverse-charge liability nets
to zero while both halves stay visible for GSTR filing. The wallet GL only ever carries
the base expense, never the gross-of-IGST figure.

### Sales entries

`--income-mode two-entry` (the default) posts **two vouchers per earning** — the sale
against the client, then a transfer clearing that client balance onto the Upwork party
that actually holds the cash:

```
1.  Type: Sales      26-27/V/RE-U/001        2.  Type: JE       26-27/LLP/Jul/003
    Dr  <Client team>            E               Dr  [Wallet]              E
        Cr  Revenue - TPT Export     E               Cr  <Client team>         E
```

The `Client Team` cell in the master is a **template**, replaced with the row's actual
`Client team` value. A statement covering several clients therefore posts to `Northwind Ltd`
and `Aurora Systems` separately, rather than dumping every sale onto one party.

`--income-mode combined` collapses both into one `Dr [Wallet] / Cr Revenue - TPT Export`.

---

## Import file

The `.csv` output is the **accounting-system import file** and carries exactly these
columns, in this order:

| | | | |
|---|---|---|---|
| Subsidiary | Transaction date | Period | Document Number |
| Type | Ledger Name | Amount | Debit |
| Credit | Amount in base currency | Amount in INR | Currency |
| Exchange Rate | I/E | Cost center Name | Narration |

- **Subsidiary** — `--entity`, default `Verve Advisory LLP`
- **Transaction date** — `dd/mm/yyyy`; **Period** — `Jul-2026`
- **Document Number** — see below
- **Type** — `Sales` or `JE`
- **Cost center Name** — the statement's `Freelancer`, falling back to the account name
  for rows with no contract (Connects, Subscription)
- **Narration** — the full form, e.g.
  `Being WHT | Withholding tax | $200.00 x 5.0% = $10.00 | Txn 100015212 / Ref 100015214 | Client: Aurora Systems | A/c: Priya Sharma`
- **Amount in base currency / Amount in INR** — filled with the line amount on **Sales**
  vouchers, left blank on JE vouchers. Amounts are already in the booking currency, so
  `Currency` is that currency and `Exchange Rate` is `1`; the USD original, the RBI rate
  and its date live on the Journal and FX Audit sheets
- **I/E** — `--ie-flag`, default `I`

Rows are written **oldest first**, so document numbers ascend with their dates.

### Document numbers

| Type | Format | Example | Restarts |
|---|---|---|---|
| JE | `{fy}/LLP/{mon}/{seq:03d}` | `26-27/LLP/Jul/001` | each month |
| Sales | `{fy}/V/RE-U/{seq:03d}` | `26-27/V/RE-U/001` | each financial year |

`{fy}` is the Indian financial year (April–March), so July 2026 is `26-27`. Each series
has its own counter, keyed on the rendered prefix — which is why the JE series restarts
monthly and Sales doesn't, without either being special-cased. Override with
`--doc-series-je` / `--doc-series-sales`; `{mm}` and `{yyyy}` are also available.

### Already-imported detection

Statements are pulled every fortnight and overlap, and sooner or later an old
one gets re-uploaded. Upwork's **Ref ID** is unique per transaction, so every Ref
ID that has been exported is recorded in `master/posted_refs.csv`; a later
statement containing it is skipped rather than posted twice.

```
Rows read      : 11
Vouchers       : 4
Skipped rows   : 8  (8 already imported)
Ref IDs logged : 3
```

The Skipped sheet names each one and when it went in — *"Already imported on
2026-09-01 as 26-27/LLP/Jul/001"*. Like document numbers, Ref IDs are only
recorded when the export is **downloaded**, so previewing a statement never
marks it as posted. `--ignore-posted` (or the toggle in the UI) forces
everything through.

**Numbers are never reused.** Every number issued is recorded in
`master/doc_registry.csv` (`document_number, type, date, prefix, seq, issued_at, source`)
and each series resumes from the highest number already there. Convert a second statement
— a different client, a later month, a re-run of the same file — and it continues where
the last one stopped:

```
run 1 (Sheet1.csv)   26-27/LLP/Jul/001 … 028      26-27/V/RE-U/001 … 006
run 2 (Globex.csv)   26-27/LLP/Jul/029 …          26-27/V/RE-U/007 …
```

Nothing is written until the output file is safely produced, so a run that fails partway
consumes no numbers. The web UI shares the same registry, so numbers issued there are
never handed out again by the CLI or vice versa. If the registry exists but can't be
read, the run **aborts** rather than risk reissuing.

`--reset-doc-numbers` (or the toggle in the UI) restarts every series at 001 and rewrites
the registry, so it never ends up holding the same number twice.

---

## The master database

[master/mapping_master.csv](master/mapping_master.csv) is the single source of truth for
how transactions are booked. **To add a freelancer, add a row to Table B — that's it.**
No upload, no code change, no restart (the web UI re-reads it on every page load).

Two tables, which may sit side by side on one sheet (as they do here) or on separate
worksheets. They are located by their header cells, so their exact position doesn't matter.
The CLI's `--mapping` still accepts a different file if you need a one-off.

**Table A** — headed by `Nature`, then four ledger columns:

```
Nature        | Debit           | Debit          | Credit  | Credit
Service Fee   | Upwork Charges  | RCM Input IGST | GL Name | RCM Output IGST
WHT           | TDS             | NA             | GL Name | NA
Withdrawal    | NO Entry        | NO Entry       | NO Entry| NO Entry
```

- `GL Name` in a cell means *substitute the wallet ledger from Table B*.
- `NA` / blank means *this leg is unused*.
- `NO Entry` in every leg means *skip this transaction type entirely*.

**Table B** — headed by `Account Name`, mapping a statement `Account name` to its wallet GL:

```
Account Name      | GL Name
Priya Sharma  | Upwork PS
Rahul Menon    | Upwork Rahul Menon
```

### Duplicates in Table B

A hand-maintained mapping sheet picks up paste slips — the same account name pasted
against someone else's ledger. When an account appears more than once with **different**
GLs, `--on-duplicate resolve` (the default) keeps the ledger whose name actually matches
the account, scored three ways:

| | Account | Candidate GL | Score |
|---|---|---|---|
| exact token | `Arjun Bhatt` | `Upwork Arjun` | 1.00 |
| initials | `Priya Sharma` | `Upwork PS` | 1.00 |
| fuzzy token | `Meera Iyer` | `Upwork Meera` | 0.88 |
| *(rejected)* | `Arjun Bhatt` | `Upwork Vikram Nair` | 0.44 |

`Upwork` is ignored when scoring, since every ledger carries it. The winner must clear
0.72 **and** beat the runner-up by 0.15 — so two equally plausible candidates are still
escalated as `AMBIGUOUS_ACCOUNT` rather than silently resolved the wrong way. Every
resolution is logged to the console and to the Exceptions sheet as an `INFO` note naming
the ledger used and the one ignored.

Pass `--on-duplicate fail` to refuse every conflicting duplicate instead.

**Optional FX table** — headed by `Period`, with a `Rate` column alongside:

```
Period   | Rate
2026-07  | 87.50
*        | 87.20      <- catch-all for months with no explicit rate
```

### How a treatment row is classified

The script infers the posting shape from the row, so you can add types without touching code:

1. Every leg is `NO Entry` → **SKIP**
2. The second debit leg is `RCM Input IGST` → **EXPENSE_RCM**
3. Debit-1 reappears as Credit-2 (the AR round-trip) → **INCOME**
4. Anything else → **SIMPLE** (one debit, one credit)

To override the inference, add a `Kind` column to Table A with one of
`EXPENSE_RCM`, `SIMPLE`, `INCOME`, `SKIP`. An explicit `Kind` always wins.

---

## Output

**`journal.xlsx`** — six sheets:

- **Import** — the import file exactly as written to `.csv`. See *Import file* above.

- **Journal** — one row per posting leg, with a running `Sr. No.`, the `Document Number`
  and `Voucher Type`, the ledger, Dr/Cr, amount, cost centre, client, the original USD
  figure, the FX rate and the date it came from, and a full narration built from the
  transaction ID, ref ID and descriptions.
- **Reconciliation** — vouchers, line count, total debit and total credit per
  transaction type, with a `Difference` column and a `TOTAL` row.
- **FX Audit** — Transaction Date → Rate Date Used → Rate → Source, with the fallback
  distance and a `Stale` flag. See *Exchange rates* above.
- **Exceptions** — with a `Severity` column. `ERROR`: `UNMAPPED_TYPE`,
  `UNMAPPED_ACCOUNT`, `AMBIGUOUS_ACCOUNT`, `MISSING_FX_RATE`, `BAD_DATE`, `BAD_AMOUNT`,
  `UNBALANCED_VOUCHER`. `INFO`: `DUPLICATE_RESOLVED`, `FX_FALLBACK`.
  A row that raises an `ERROR` is **not posted** (except an unbalanced voucher, which is
  posted *and* flagged so you can see it in context).
- **Skipped** — zero-amount rows, blank transaction types, and `NO Entry` types, with
  the reason.

**`journal.csv`** — the **import file** (UTF-8 with BOM, so Excel opens it cleanly). This
is the file you feed to the accounting system.

Every voucher is checked for debit == credit before it is written; a failure is recorded
and makes the process exit `1`.

---

## Data-quality notes on the supplied sample

The tool surfaces these rather than guessing:

- **`Arjun Bhatt` appeared twice**, mapped to both `Upwork Arjun` and `Upwork Vikram
  Nair`. Confirmed as a paste slip, so the duplicate row is **removed from the master**.
  The name-match resolver described above remains, to catch the next one.
- **Non-breaking spaces and doubled spaces** (`Upwork·Rahul·Menon`, `Vikram  Nair`)
  would otherwise break lookups silently. Cleaned in the master, and all keys and values
  are still Unicode-normalised and whitespace-folded before matching, so a name pasted
  from Upwork with a stray NBSP resolves anyway.
- **`Withdrawal` of −$5,000** is correctly skipped as `NO Entry` — it is a movement of
  the freelancer's own funds, not a P&L event.

## Sample run

```
FX: fetching 2 month(s) from RBI: 2026-06, 2026-07
FX: 2026-06: 21 published rate(s), 9 non-publishing day(s)
FX: 2026-07: 23 published rate(s), 8 non-publishing day(s)
FX: cache saved -> fx_cache.csv
Rows read      : 29
Vouchers       : 34
Journal lines  : 100
Total debits   : 196,660.79 INR
Total credits  : 196,660.79 INR
Difference     : 0.00
Skipped rows   : 1
FX dates       : 15 (14x RBI, 1x RBI (prev-day fallback))
Errors         : 0
Notes (INFO)   : 1
  note: Table B lists 'Arjun Bhatt' more than once; used 'Upwork Arjun' as the
        name matches, ignored 'Upwork Vikram Nair'
```

A second run prints `FX: all 15 transaction date(s) already cached` and makes no network
requests.
