#!/usr/bin/env python3
"""
rbi_fx.py -- USD/INR reference rates from the RBI Reference Rate Archive.

The archive (https://rbi.org.in/scripts/ReferenceRateArchive.aspx) is an ASP.NET
WebForms page, so a plain GET returns only the empty form.  The sequence is:

    1. GET  the page, scrape __VIEWSTATE / __VIEWSTATEGENERATOR / __EVENTVALIDATION
    2. POST those back along with:
           chkUSD      = "on"          (which currency column to return)
           txtFromDate = "dd/mm/yyyy"  (the form accepts a RANGE, so one request
           txtToDate   = "dd/mm/yyyy"   per month is enough -- not one per day)
           btnSubmit   = " GO "
    3. parse the table headed  ['Date', 'USD (INR / 1 USD)']

RBI publishes on business days only, so a statement dated on a weekend or a bank
holiday has no rate of its own.  `RbiFxProvider.rate_for` then walks backwards a
day at a time to the most recent published rate, and records how far it had to
go, so the choice is visible in the audit output rather than buried.

Resolution order for any one date:

    FX FILE  (user-supplied, wins for dates it covers)
      -> CACHE / RBI  (exact date)
      -> CACHE / RBI  (previous published day, within --fx-max-fallback-days)
      -> MANUAL       (--fx-rate, the safety net when RBI is unreachable)
      -> UNAVAILABLE  (row is reported, never silently mispriced)

Nothing here raises on a network problem: an unreachable RBI degrades to the
cache, then to the manual rate.  A run never dies because a website was down.
"""

from __future__ import annotations

import csv
import time
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

import requests
from bs4 import BeautifulSoup

ARCHIVE_URL = "https://rbi.org.in/scripts/ReferenceRateArchive.aspx"

#: A real browser UA -- the archive is inconsistent with unusual clients.
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

#: ASP.NET hidden fields that must be echoed back on every POST.
ASPNET_TOKENS = ("__VIEWSTATE", "__VIEWSTATEGENERATOR", "__EVENTVALIDATION")

#: Written into the cache for a date RBI confirmed has no rate (holiday/weekend),
#: so repeat runs don't keep asking about the same Sunday.
NO_RATE = "NO_RATE"

CACHE_COLUMNS = ("date", "rate", "source", "fetched_at")


# --------------------------------------------------------------------------- #
# Rate decisions
# --------------------------------------------------------------------------- #

@dataclass
class RateDecision:
    """Which rate was used for one transaction date, and why."""
    txn_date: date
    rate: Decimal | None
    rate_date: date | None
    source: str              # RBI | CACHE | FX FILE | MANUAL | UNAVAILABLE
    days_back: int = 0
    note: str = ""

    @property
    def ok(self) -> bool:
        return self.rate is not None

    @property
    def source_label(self) -> str:
        """Source as shown in the audit, marking a previous-day fallback."""
        if self.days_back > 0:
            return f"{self.source} (prev-day fallback)"
        return self.source


# --------------------------------------------------------------------------- #
# Cache
# --------------------------------------------------------------------------- #

class RateCache:
    """A CSV of `date,rate,source,fetched_at`, one row per calendar day.

    CSV rather than SQLite deliberately: an accountant can open it, eyeball a
    rate and correct one by hand without any tooling.

    A row whose source is NO_RATE means "RBI confirmed nothing is published for
    this date" -- caching the *absence* is what stops every run re-asking about
    the same weekend.
    """

    def __init__(self, path: Path) -> None:
        self.path = path
        self.rates: dict[date, Decimal] = {}
        self.known_empty: set[date] = set()
        self._load()

    def _load(self) -> None:
        if not self.path.exists():
            return
        try:
            with self.path.open(newline="", encoding="utf-8") as handle:
                for row in csv.DictReader(handle):
                    parsed = _parse_iso(row.get("date", ""))
                    if parsed is None:
                        continue
                    if (row.get("source") or "").strip() == NO_RATE:
                        self.known_empty.add(parsed)
                        continue
                    try:
                        self.rates[parsed] = Decimal((row.get("rate") or "").strip())
                    except (InvalidOperation, ValueError):
                        continue
        except OSError:
            # A corrupt or unreadable cache must not stop a run; we just refetch.
            self.rates.clear()
            self.known_empty.clear()

    def covers(self, day: date) -> bool:
        """True when we already know this date's answer, rate or no rate."""
        return day in self.rates or day in self.known_empty

    def update(self, rates: dict[date, Decimal], empties: set[date]) -> None:
        self.rates.update(rates)
        self.known_empty.update(empties)
        self.known_empty -= set(self.rates)  # a real rate supersedes a NO_RATE

    def save(self) -> None:
        """Rewrite the cache, newest first."""
        self.path.parent.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows = [{"date": d.isoformat(), "rate": f"{r}", "source": "RBI", "fetched_at": stamp}
                for d, r in self.rates.items()]
        rows += [{"date": d.isoformat(), "rate": "", "source": NO_RATE, "fetched_at": stamp}
                 for d in self.known_empty]
        rows.sort(key=lambda r: r["date"], reverse=True)
        with self.path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=CACHE_COLUMNS)
            writer.writeheader()
            writer.writerows(rows)


def _parse_iso(text: str) -> date | None:
    try:
        return datetime.strptime(text.strip(), "%Y-%m-%d").date()
    except (ValueError, AttributeError):
        return None


# --------------------------------------------------------------------------- #
# Scraper
# --------------------------------------------------------------------------- #

class RbiArchiveClient:
    """Fetches USD/INR rows from the archive for a date range."""

    def __init__(self, *, timeout: float = 30.0, delay: float = 1.0,
                 retries: int = 3, log=None) -> None:
        self.timeout = timeout
        self.delay = delay          # polite pause between requests
        self.retries = retries
        self.log = log or (lambda message: None)
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        })
        self._tokens: dict[str, str] = {}
        self._last_request: float = 0.0

    # -- plumbing ----------------------------------------------------------- #

    def _pause(self) -> None:
        """Keep a minimum gap between hits so we stay a good citizen."""
        elapsed = time.monotonic() - self._last_request
        if self._last_request and elapsed < self.delay:
            time.sleep(self.delay - elapsed)
        self._last_request = time.monotonic()

    def _remember_tokens(self, html: str) -> None:
        """Capture the ASP.NET hidden fields from whatever page we just got.

        Every response carries a fresh set; reusing a stale __VIEWSTATE makes the
        next POST fail, so we refresh from each response rather than only the GET.
        """
        soup = BeautifulSoup(html, "lxml")
        for name in ASPNET_TOKENS:
            element = soup.find("input", {"name": name})
            if element is not None:
                self._tokens[name] = element.get("value", "")

    def _request(self, method: str, **kwargs) -> requests.Response | None:
        """One request with retry + exponential backoff. Returns None if it fails."""
        for attempt in range(1, self.retries + 1):
            try:
                self._pause()
                response = self.session.request(method, ARCHIVE_URL,
                                                timeout=self.timeout, **kwargs)
                response.raise_for_status()
                self._remember_tokens(response.text)
                return response
            except requests.RequestException as exc:
                wait = 2 ** (attempt - 1)
                if attempt == self.retries:
                    self.log(f"RBI {method} failed after {attempt} attempts: "
                             f"{type(exc).__name__}: {exc}")
                    return None
                self.log(f"RBI {method} attempt {attempt} failed ({type(exc).__name__}); "
                         f"retrying in {wait}s")
                time.sleep(wait)
        return None

    # -- public ------------------------------------------------------------- #

    def fetch_range(self, start: date, end: date) -> tuple[dict[date, Decimal], bool]:
        """Rates published between `start` and `end` inclusive.

        Returns (rates, reached_server). `reached_server` distinguishes "RBI said
        there is nothing here" from "we never got an answer" -- only the former
        justifies caching the dates as NO_RATE.
        """
        if not self._tokens:
            landing = self._request("GET")
            if landing is None:
                return {}, False

        payload = {
            "__EVENTTARGET": "",
            "__EVENTARGUMENT": "",
            "UsrFontCntr$txtSearch": "",
            "chkUSD": "on",
            "txtFromDate": start.strftime("%d/%m/%Y"),
            "txtToDate": end.strftime("%d/%m/%Y"),
            "btnSubmit": " GO ",
            **self._tokens,
        }
        response = self._request("POST", data=payload,
                                 headers={"Referer": ARCHIVE_URL})
        if response is None:
            return {}, False
        return parse_rate_table(response.text), True


def parse_rate_table(html: str) -> dict[date, Decimal]:
    """Pull {date: rate} out of the archive's results table.

    The results table is identified by its header (`Date` + a `USD` column)
    rather than by position, so extra layout tables on the page can come and go
    without breaking the parse.
    """
    soup = BeautifulSoup(html, "lxml")
    rates: dict[date, Decimal] = {}

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue
        header = [c.get_text(" ", strip=True).upper() for c in rows[0].find_all(["td", "th"])]
        if not header or "DATE" not in header[0]:
            continue

        # Which column holds USD? ('USD (INR / 1 USD)' when only USD was ticked,
        # but the page can return several currency columns.)
        usd_index = next((i for i, h in enumerate(header) if i > 0 and "USD" in h), None)
        if usd_index is None:
            continue

        for row in rows[1:]:
            cells = [c.get_text(" ", strip=True) for c in row.find_all(["td", "th"])]
            if len(cells) <= usd_index:
                continue
            try:
                day = datetime.strptime(cells[0].strip(), "%d/%m/%Y").date()
            except ValueError:
                continue
            raw = cells[usd_index].replace(",", "").strip()
            try:
                rate = Decimal(raw)
            except InvalidOperation:
                continue
            if rate > 0:
                rates[day] = rate

        if rates:
            return rates
    return rates


# --------------------------------------------------------------------------- #
# Provider
# --------------------------------------------------------------------------- #

@dataclass
class FxProvider:
    """A flat rate for every date -- USD mode, or an explicit --fx-source manual."""
    rate: Decimal | None = None
    source: str = "MANUAL"
    decisions: list[RateDecision] = field(default_factory=list)

    def prepare(self, dates) -> None:  # noqa: D102 - nothing to pre-fetch
        return

    def rate_for(self, day: date) -> RateDecision:
        decision = RateDecision(
            txn_date=day, rate=self.rate, rate_date=day if self.rate else None,
            source=self.source if self.rate else "UNAVAILABLE")
        self.decisions.append(decision)
        return decision

    @property
    def warnings(self) -> list[str]:
        return []


class RbiFxProvider:
    """Per-date USD/INR rates, with cache, fallback and graceful degradation."""

    def __init__(self, *, cache_path: Path, file_rates: dict[date, Decimal] | None = None,
                 manual_rate: Decimal | None = None, max_fallback_days: int = 7,
                 stale_after_days: int = 3, offline: bool = False,
                 timeout: float = 30.0, delay: float = 1.0, retries: int = 3,
                 log=None) -> None:
        self.log = log or (lambda message: None)
        self.cache = RateCache(cache_path)
        self.file_rates = file_rates or {}
        self.manual_rate = manual_rate
        self.max_fallback_days = max_fallback_days
        self.stale_after_days = stale_after_days
        self.offline = offline
        self.client = None if offline else RbiArchiveClient(
            timeout=timeout, delay=delay, retries=retries, log=self.log)
        self.decisions: list[RateDecision] = []
        self.server_reachable: bool | None = None
        self._months_tried: set[tuple[int, int]] = set()

    # -- fetching ----------------------------------------------------------- #

    def prepare(self, dates) -> None:
        """Fetch, in one request per month, everything the run will need.

        The window is widened by `max_fallback_days` before the earliest date so
        a fallback never has to trigger a second round-trip mid-run.
        """
        wanted = sorted({d for d in dates if d is not None})
        if not wanted:
            return

        earliest = min(wanted) - timedelta(days=self.max_fallback_days)
        months = sorted({(d.year, d.month) for d in wanted} |
                        {(earliest.year, earliest.month)})

        missing = [m for m in months if self._month_needs_fetch(m, wanted, earliest)]
        if not missing:
            self.log(f"FX: all {len(wanted)} transaction date(s) already cached "
                     f"in {self.cache.path.name}")
            return
        if self.offline:
            self.log("FX: --fx-offline set, using cache only")
            return

        self.log(f"FX: fetching {len(missing)} month(s) from RBI: "
                 + ", ".join(f"{y}-{m:02d}" for y, m in missing))
        for year, month in missing:
            self._fetch_month(year, month)

        try:
            self.cache.save()
            self.log(f"FX: cache saved -> {self.cache.path}")
        except OSError as exc:
            self.log(f"FX: could not write cache ({exc}); continuing in memory")

    def _month_needs_fetch(self, month_key, wanted, earliest) -> bool:
        """True if any date we might consult in this month is still unknown."""
        year, month = month_key
        day = date(year, month, 1)
        last = _month_end(day)
        cursor, horizon = max(day, earliest), min(last, max(wanted))
        while cursor <= horizon:
            if not self.cache.covers(cursor):
                return True
            cursor += timedelta(days=1)
        return False

    def _fetch_month(self, year: int, month: int) -> None:
        start = date(year, month, 1)
        end = _month_end(start)
        rates, reached = self.client.fetch_range(start, end)

        if not reached:
            self.server_reachable = False
            self.log(f"FX: RBI unreachable for {year}-{month:02d}; "
                     f"falling back to cache/manual")
            return
        self.server_reachable = True
        self._months_tried.add((year, month))

        if not rates:
            self.log(f"FX: RBI returned no rows for {year}-{month:02d}")
            return

        # Interior gaps -- dates between the first and last published day with no
        # rate -- are genuine non-publishing days, so cache the absence. Dates
        # after the last published day are simply not out yet; leave them unknown
        # so a later run picks them up.
        published_last = max(rates)
        empties = set()
        cursor = min(rates)
        while cursor <= published_last:
            if cursor not in rates:
                empties.add(cursor)
            cursor += timedelta(days=1)

        self.cache.update(rates, empties)
        self.log(f"FX: {year}-{month:02d}: {len(rates)} published rate(s), "
                 f"{len(empties)} non-publishing day(s)")

    # -- resolution --------------------------------------------------------- #

    def rate_for(self, day: date) -> RateDecision:
        """Resolve one transaction date, recording the decision for the audit."""
        decision = self._resolve(day)
        self.decisions.append(decision)
        return decision

    def _resolve(self, day: date) -> RateDecision:
        # 1. A user-supplied FX file wins outright for dates it covers.
        if day in self.file_rates:
            return RateDecision(day, self.file_rates[day], day, "FX FILE")

        # 2. Exact date from cache (which is where anything just fetched landed).
        if day in self.cache.rates:
            return RateDecision(day, self.cache.rates[day], day,
                                "CACHE" if self.offline else "RBI")

        # 3. Walk backwards to the most recent published day.
        for back in range(1, self.max_fallback_days + 1):
            earlier = day - timedelta(days=back)
            if earlier in self.file_rates:
                return RateDecision(day, self.file_rates[earlier], earlier,
                                    "FX FILE", days_back=back)
            if earlier in self.cache.rates:
                return RateDecision(
                    day, self.cache.rates[earlier], earlier,
                    "CACHE" if self.offline else "RBI", days_back=back,
                    note=f"no rate published for {day.isoformat()}")

        # 4. Nothing from RBI or cache -- the manual rate is the safety net.
        if self.manual_rate is not None:
            return RateDecision(
                day, self.manual_rate, day, "MANUAL",
                note=f"no RBI rate within {self.max_fallback_days} days")

        return RateDecision(
            day, None, None, "UNAVAILABLE",
            note=f"no RBI rate within {self.max_fallback_days} days and no --fx-rate given")

    # -- reporting ---------------------------------------------------------- #

    @property
    def warnings(self) -> list[str]:
        """One line per date whose fallback reached back further than allowed."""
        seen: set[date] = set()
        out: list[str] = []
        for d in self.decisions:
            if d.txn_date in seen:
                continue
            seen.add(d.txn_date)
            if d.days_back > self.stale_after_days:
                out.append(
                    f"{d.txn_date.isoformat()}: used the rate from "
                    f"{d.rate_date.isoformat()} ({d.days_back} days earlier, "
                    f"threshold {self.stale_after_days})")
            elif d.source == "MANUAL":
                out.append(f"{d.txn_date.isoformat()}: no RBI rate available, "
                           f"used the manual rate {d.rate}")
            elif not d.ok:
                out.append(f"{d.txn_date.isoformat()}: no rate available from any source")
        return out

    def audit_rows(self) -> list[dict]:
        """Transaction Date -> Rate Date Used -> Rate -> Source, one row per date."""
        by_date: dict[date, RateDecision] = {}
        counts: dict[date, int] = {}
        for d in self.decisions:
            by_date.setdefault(d.txn_date, d)
            counts[d.txn_date] = counts.get(d.txn_date, 0) + 1

        rows = []
        for day in sorted(by_date):
            d = by_date[day]
            rows.append({
                "Transaction Date": day,
                "Rate Date Used": d.rate_date,
                "Rate (INR/USD)": float(d.rate) if d.rate is not None else None,
                "Source": d.source_label,
                "Days Back": d.days_back,
                "Transactions": counts[day],
                "Stale": "YES" if d.days_back > self.stale_after_days else "",
                "Note": d.note,
            })
        return rows


def _month_end(day: date) -> date:
    """Last calendar day of `day`'s month."""
    if day.month == 12:
        return date(day.year, 12, 31)
    return date(day.year, day.month + 1, 1) - timedelta(days=1)


def load_fx_file_by_date(path: Path) -> dict[date, Decimal]:
    """Read a user FX file into {date: rate}.

    Accepts exact dates (`2026-07-15`, `15/07/2026`) and whole months
    (`2026-07`), the latter expanded across every day of that month so a monthly
    rate keeps working exactly as it did before RBI fetching existed.
    """
    import pandas as pd  # local import: only needed when a file is supplied

    suffix = path.suffix.lower()
    frame = (pd.read_csv(path, dtype=object, keep_default_na=False)
             if suffix == ".csv" else
             pd.read_excel(path, dtype=object))

    columns = {str(c).strip().lower(): c for c in frame.columns}
    period_col = next((columns[k] for k in ("period", "date", "month") if k in columns), None)
    rate_col = next((columns[k] for k in ("rate", "fx rate", "fxrate", "usdinr") if k in columns),
                    None)
    if period_col is None or rate_col is None:
        raise ValueError(f"{path.name} needs 'Period' (or 'Date') and 'Rate' columns")

    out: dict[date, Decimal] = {}
    for _, row in frame.iterrows():
        key = str(row[period_col]).strip()
        try:
            rate = Decimal(str(row[rate_col]).strip())
        except (InvalidOperation, ValueError):
            continue
        if rate <= 0:
            continue

        exact = None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
            try:
                exact = datetime.strptime(key, fmt).date()
                break
            except ValueError:
                continue
        if exact is not None:
            out[exact] = rate
            continue

        try:  # a whole month: apply to every day in it
            first = datetime.strptime(key, "%Y-%m").date()
        except ValueError:
            continue
        cursor, last = first, _month_end(first)
        while cursor <= last:
            out.setdefault(cursor, rate)
            cursor += timedelta(days=1)
    return out
