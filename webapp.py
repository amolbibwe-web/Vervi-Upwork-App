#!/usr/bin/env python3
"""
webapp.py -- Vervi-Upwork, the local browser UI for upwork_to_journal.

Runs a small Flask server on localhost so the conversion can be driven without
the command line: drop in the Upwork statement, pick the options, and get the
import file, journal, FX audit and exceptions rendered in the page with download
links for the generated .xlsx / .csv.

    python webapp.py                  ->  http://127.0.0.1:5000
    python webapp.py --port 8000      ->  http://127.0.0.1:8000

All the accounting logic lives in upwork_to_journal.py; this module only handles
uploads, option parsing and presentation.  Nothing is duplicated, so the web UI
and the CLI can never drift apart -- including the document-number registry, so
numbers issued here are never reissued by a later CLI run or vice versa.

The mapping is never uploaded -- it always comes from the master database, shown
in the sidebar so the Account Name -> GL Name lookup can be checked before a run.

Bound to 127.0.0.1 only -- this is a local convenience UI, not a hosted service.
"""

from __future__ import annotations

import argparse
import csv
import io
import itertools
import os
import zipfile
import secrets
import shutil
import time
import uuid
from datetime import datetime
from decimal import Decimal
from html import escape
from pathlib import Path
from tempfile import mkdtemp

import pandas as pd
from flask import (Flask, abort, redirect, render_template_string, request,
                   send_file, session, url_for)
from werkzeug.utils import secure_filename

from rbi_fx import (ARCHIVE_URL, FxProvider, RateCache, RbiArchiveClient,
                    RbiFxProvider)
from datetime import datetime

from upwork_to_journal import (DEFAULT_DOC_REGISTRY, DEFAULT_DOC_SERIES,
                               VOUCHER_JE, VOUCHER_SALES, DocumentNumberer,
                               Reporter, add_treatment, add_wallet,
                               build_import_frame, build_journal,
                               PostedLedger, DEFAULT_POSTED_LEDGER,
                               build_reconciliation, clean_text, delete_treatment,
                               delete_wallet, load_mapping, load_statement,
                               to_decimal, update_treatment, update_wallet,
                               write_outputs)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB upload ceiling

# A fresh key each start, so sessions do not survive a restart.  Set
# VERVI_SECRET to keep people signed in across restarts.
app.secret_key = os.environ.get("VERVI_SECRET") or secrets.token_hex(32)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    # Only set this when the app is reached over https, or the cookie is
    # dropped and sign-in silently fails on plain http.
    SESSION_COOKIE_SECURE=os.environ.get("VERVI_HTTPS_ONLY", "") == "1",
)

APP_NAME = "Vervi-Upwork"
COMPANY = "Verve Advisory LLP"

#: Credentials. Override with VERVI_USER / VERVI_PASS -- and do override them
#: before putting this anywhere the public can reach: admin/admin is guessed by
#: the first automated scan that finds the URL.
APP_USER = os.environ.get("VERVI_USER", "admin")
APP_PASS = os.environ.get("VERVI_PASS", "admin")

#: Routes reachable without signing in.
PUBLIC_ENDPOINTS = {"login", "healthz"}

#: Simple per-IP brute-force throttle: after this many failures, that address
#: waits before it may try again.  Not a substitute for a real password.
MAX_ATTEMPTS = 5
LOCKOUT_SECONDS = 60
_attempts: dict[str, list] = {}

HERE = Path(__file__).parent.resolve()
SAMPLE_STATEMENT = HERE / "data" / "Sheet1.csv"

#: The accounting master database.  Table A (treatment by transaction type) and
#: Table B (Account Name -> GL Name) both live here, and are edited in this one
#: file when a freelancer joins.
MASTER_MAPPING = HERE / "master" / "mapping_master.csv"

#: Completed conversions: the outputs are kept here so a statement converted
#: last fortnight can be found and downloaded again without re-running it.
EXPORTS_DIR = HERE / "exports"
HISTORY_FILE = EXPORTS_DIR / "history.csv"
HISTORY_COLUMNS = ("converted_at", "statement", "rows_read", "vouchers",
                   "journal_lines", "duplicates", "doc_numbers", "folder")

#: run id -> {"xlsx": Path, "csv": Path}, so downloads survive the redirect.
RUNS: dict[str, dict[str, Path]] = {}

ALLOWED_SUFFIXES = {".csv", ".xlsx", ".xlsm", ".xls"}


# --------------------------------------------------------------------------- #
# Styling
# --------------------------------------------------------------------------- #

#: Inter is the face used across Vervi-Books, so the two products read as one
#: family.  Roboto stays in the stack as the first fallback.
FONT_LINK = ('<link rel="preconnect" href="https://fonts.googleapis.com">'
             '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
             '<link href="https://fonts.googleapis.com/css2?'
             'family=Inter:wght@400;500;600;700&family=Roboto+Mono:wght@400;500'
             '&display=swap" rel="stylesheet">')

#: Dark palette, written once and applied twice below -- once for people whose
#: system is dark and who have not chosen, once for an explicit choice.
DARK_TOKENS = """
    --bg:#0d1117; --panel:#161b24; --panel-2:#12171f; --ink:#e6ebf4; --muted:#93a0b5;
    --line:#252c38; --line-2:#1d232d;
    --brand:#8579ff; --brand-2:#a99dff; --brand-soft:#1e1a3d;
    --teal:#2dd4bf; --sky:#57b6f5; --amber:#e0a83c; --pink:#f06595;
    --ok:#3ecf8e; --ok-bg:#0e2b20; --warn:#e0a83c; --warn-bg:#2e2510;
    --err:#f77a95; --err-bg:#33161f;
    --shadow:0 1px 2px rgba(0,0,0,.35), 0 8px 24px -14px rgba(0,0,0,.75);
"""

BASE_CSS = """
:root{
  --bg:#f5f6fa; --panel:#fff; --panel-2:#fafbfd; --ink:#171b26; --muted:#68738a;
  --line:#e5e8f0; --line-2:#eff1f7;
  --brand:#5b4bf5; --brand-2:#7c6cff; --brand-soft:#eeecff;
  --side-ink:#f2f0ff; --side-muted:#b6b0e8;
  --teal:#0f8b7e; --sky:#0b74c4; --amber:#a86a00; --pink:#d6336c;
  --ok:#12794f; --ok-bg:#e8f7f0; --warn:#8a5a00; --warn-bg:#fff5e3;
  --err:#c02646; --err-bg:#ffeef2;
  --shadow:0 1px 2px rgba(16,24,40,.04), 0 6px 18px -10px rgba(16,24,40,.16);
  --font:"Inter","Roboto",system-ui,-apple-system,"Segoe UI",sans-serif;
  --mono:"Roboto Mono",ui-monospace,Consolas,monospace;
}
/* System preference, unless the reader has explicitly picked light. */
@media (prefers-color-scheme:dark){
  :root:not([data-theme="light"]){""" + DARK_TOKENS + """}
}
/* An explicit pick from the toggle wins over the system in both directions. */
:root[data-theme="dark"]{""" + DARK_TOKENS + """}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:var(--font);
     font-size:15.5px;line-height:1.55;-webkit-font-smoothing:antialiased}

/* ---- shell: fixed sidebar + scrolling main ---- */
.shell{display:flex;min-height:100vh}
.side{width:280px;flex:0 0 280px;color:var(--side-ink);padding:24px 20px;
      position:sticky;top:0;height:100vh;overflow-y:auto;
      background:linear-gradient(168deg,#5b4bf5 0%,#6d43e8 38%,#3f2d9e 78%,#241a63 100%);
      position:sticky}
/* soft colour bloom, so the panel reads as branded rather than merely dark */
.side::before{content:"";position:absolute;inset:0;pointer-events:none;
  background:radial-gradient(420px 200px at 15% 4%,rgba(255,255,255,.22),transparent 70%),
             radial-gradient(320px 260px at 95% 32%,rgba(236,72,153,.30),transparent 70%),
             radial-gradient(300px 240px at 0% 82%,rgba(45,212,191,.22),transparent 70%)}
.side>*{position:relative;z-index:1}
.main{flex:1;min-width:0;padding:0 32px 60px}

.logo{display:flex;align-items:center;gap:12px;margin-bottom:22px}
.logo .mark{width:42px;height:42px;border-radius:12px;flex:0 0 42px;
     background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.28);
     backdrop-filter:blur(6px);
     display:grid;place-items:center;color:#fff}
.logo .mark svg{width:23px;height:23px}
.logo .nm{font-size:17.5px;font-weight:700;letter-spacing:-.015em;line-height:1.15}
.logo .co{font-size:12.5px;color:var(--side-muted);margin-top:1px}

.sblock{background:rgba(255,255,255,.09);border:1px solid rgba(255,255,255,.14);
        border-radius:12px;padding:14px 15px;margin-top:13px;backdrop-filter:blur(6px)}
.slabel{font-size:10.5px;text-transform:uppercase;letter-spacing:.1em;
        color:#fff;opacity:.72;font-weight:700;margin-bottom:9px;
        display:flex;align-items:center;gap:7px}
.slabel::before{content:"";width:7px;height:7px;border-radius:50%;
        background:var(--accent,#fff);box-shadow:0 0 9px var(--accent,#fff)}
.sblock.c1{--accent:#67e8f9} .sblock.c2{--accent:#fca5f1}
.sblock.c3{--accent:#86efac} .sblock.c4{--accent:#fcd34d}
.srow{display:flex;justify-content:space-between;gap:10px;padding:4px 0;font-size:13px}
.srow .k{color:var(--side-muted)}
.srow .v{font-weight:500;text-align:right;font-variant-numeric:tabular-nums;color:#fff}
.srow .v.big{font-size:15px;font-weight:700}
.srow .v.ok{color:#86efac} .srow .v.bad{color:#fda4af}
.snote{font-size:12px;color:var(--side-muted);line-height:1.5;margin-top:9px}
.snote code{background:rgba(255,255,255,.14);border:0;color:#fff;
            padding:1px 5px;border-radius:4px;font-size:11.5px}
.snote a{color:#fff;text-decoration:underline;text-underline-offset:2px}
/* The sidebar is a purple gradient, so the page's ghost button -- brand-purple
   text on transparent -- disappears into it. Sidebar buttons get their own
   treatment: solid white for the primary action, outlined white for the rest. */
.side button{background:#fff;color:#3f2d9e;border:0;font-weight:600;
  box-shadow:0 2px 8px -2px rgba(20,10,60,.45)}
.side button:hover{background:#f4f2ff;color:#2c1d70;filter:none}
.side button.sidealt{background:rgba(255,255,255,.14);color:#fff;
  border:1px solid rgba(255,255,255,.45);box-shadow:none;font-weight:500}
.side button.sidealt:hover{background:rgba(255,255,255,.26);color:#fff}
.side .row a{text-decoration:none}
/* Numbered how-to in the sidebar: one short line per step. */
.steps-list{margin:0;padding:0;list-style:none;counter-reset:s}
.steps-list li{counter-increment:s;position:relative;padding:3px 0 3px 24px;
  font-size:12.5px;color:var(--side-muted);line-height:1.45}
.steps-list li::before{content:counter(s);position:absolute;left:0;top:4px;
  width:16px;height:16px;border-radius:50%;background:rgba(255,255,255,.16);
  color:#fff;font-size:9.5px;font-weight:700;display:grid;place-items:center}
.steps-list strong{color:#fff;font-weight:500}

/* ---- thin top bar ---- */
.top{display:flex;align-items:baseline;gap:12px;padding:26px 0 18px;
     border-bottom:1px solid var(--line);margin-bottom:22px;flex-wrap:wrap}
.top h1{font-size:22px;margin:0;font-weight:700;letter-spacing:-.02em}
.top .s{color:var(--muted);font-size:14px}
.topright{margin-left:auto;display:flex;align-items:center;gap:9px;align-self:center}
.iconbtn{width:38px;height:38px;padding:0;border-radius:10px;background:var(--panel);
  border:1px solid var(--line);color:var(--muted);display:grid;place-items:center;
  cursor:pointer;box-shadow:none;font-size:16px;line-height:1;transition:.15s}
.iconbtn:hover{background:var(--brand-soft);border-color:var(--brand);color:var(--brand);
  transform:none;filter:none}
.iconbtn .sun{display:none}
:root[data-theme="dark"] .iconbtn .sun{display:block}
:root[data-theme="dark"] .iconbtn .moon{display:none}
@media (prefers-color-scheme:dark){
  :root:not([data-theme="light"]) .iconbtn .sun{display:block}
  :root:not([data-theme="light"]) .iconbtn .moon{display:none}
}
.backbtn{display:inline-flex;align-items:center;gap:8px;height:38px;padding:0 16px;
  border-radius:10px;background:var(--brand-soft);border:1px solid transparent;
  color:var(--brand);font-size:14px;font-weight:600;text-decoration:none;transition:.15s}
.backbtn:hover{background:var(--brand);color:#fff;text-decoration:none}
.signout{display:inline-flex;align-items:center;gap:7px;height:38px;padding:0 15px;
  border-radius:10px;background:var(--panel);border:1px solid var(--line);
  color:var(--ink);font-size:14px;font-weight:500;text-decoration:none;transition:.15s}
.signout:hover{background:var(--err-bg);border-color:var(--err);color:var(--err);
  text-decoration:none}
.signout .who{color:var(--muted);font-weight:400}
.signout:hover .who{color:var(--err)}

/* ---- panels ---- */
.panel{background:var(--panel);border:1px solid var(--line);border-radius:14px;
       box-shadow:var(--shadow);padding:22px;margin-bottom:18px}
/* Upload and Generate sit side by side, each taking half the row. */
.steps{display:grid;gap:18px;grid-template-columns:1fr 1fr;align-items:start;
       margin-bottom:18px}
.steps .panel{margin-bottom:0;height:100%}
.step{display:flex;align-items:center;gap:10px;margin-bottom:3px}
.step .n{width:24px;height:24px;border-radius:50%;
   background:linear-gradient(135deg,var(--brand),var(--brand-2));color:#fff;
   display:grid;place-items:center;font-size:12.5px;font-weight:700;flex:0 0 24px}
.step h2{font-size:16.5px;margin:0;font-weight:700;letter-spacing:-.01em}
.psub{color:var(--muted);font-size:14px;margin:0 0 14px 34px}

/* ---- form ---- */
.grid{display:grid;gap:14px;grid-template-columns:repeat(auto-fit,minmax(195px,1fr))}
label{display:block;font-weight:500;font-size:13.5px;margin-bottom:5px;color:var(--muted)}
input[type=text],select{width:100%;padding:9px 11px;border:1px solid var(--line);
  border-radius:9px;background:var(--panel-2);color:var(--ink);font-size:14.5px;
  font-family:var(--font);transition:border-color .15s,box-shadow .15s}
input[type=text]:focus,select:focus{outline:0;border-color:var(--brand);
  box-shadow:0 0 0 3px rgba(91,75,245,.18)}
/* Half the height of the full-width version it replaces. */
.drop{border:2px dashed var(--line);border-radius:11px;background:var(--panel-2);
      padding:13px 16px;text-align:left;cursor:pointer;transition:.18s;position:relative;
      display:flex;align-items:center;gap:13px}
.drop:hover,.drop.over{border-color:var(--brand);background:var(--brand-soft)}
.drop input{position:absolute;inset:0;opacity:0;cursor:pointer}
.drop .big{font-size:21px;line-height:1;flex:0 0 auto}
.drop .t{font-weight:500;font-size:14.5px}
.drop .s{color:var(--muted);font-size:13px}
.drop.has{border-style:solid;border-color:var(--teal);background:rgba(15,139,126,.07)}
details{margin-top:16px;border-top:1px solid var(--line);padding-top:14px}
/* A panel-level disclosure: no rule above it, and a heading-sized summary. */
details.fold{margin:0;border-top:0;padding-top:0}
details.fold>summary{font-size:16px;font-weight:700;color:var(--ink);letter-spacing:-.01em}
details.fold>summary:hover{color:var(--brand)}
details.fold[open]>summary{margin-bottom:14px}
summary{cursor:pointer;font-size:14px;color:var(--muted);font-weight:500;
        list-style:none;display:flex;align-items:center;gap:6px}
summary::-webkit-details-marker{display:none}
summary::before{content:"\\25B8";font-size:10px;transition:.15s}
details[open] summary::before{transform:rotate(90deg)}
summary:hover{color:var(--brand)}
.switch{display:flex;align-items:center;gap:9px;font-size:14px;margin:14px 0 0;
        cursor:pointer;color:var(--ink);font-weight:400}
.switch input{width:16px;height:16px;accent-color:var(--brand)}

button{background:linear-gradient(135deg,var(--brand),var(--brand-2));color:#fff;border:0;
  border-radius:9px;padding:11px 21px;font-size:14.5px;font-weight:600;cursor:pointer;
  font-family:var(--font);transition:filter .15s,transform .12s;
  box-shadow:0 4px 12px -5px rgba(91,75,245,.7)}
button:hover{filter:brightness(1.08);transform:translateY(-1px)}
button.ghost{background:transparent;color:var(--brand);border:1px solid var(--line);
             box-shadow:none;font-weight:500}
button.ghost:hover{background:var(--brand-soft);border-color:var(--brand);transform:none}
.row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:18px}
a{color:var(--brand);text-decoration:none}
a:hover{text-decoration:underline}
.hint{color:var(--muted);font-size:13px}

/* ---- banners ---- */
.banner{border-radius:10px;padding:12px 15px;margin:0 0 16px;font-size:14.5px;
        border:1px solid;display:flex;gap:10px;align-items:flex-start}
.banner.ok{background:var(--ok-bg);border-color:rgba(18,121,79,.3);color:var(--ok)}
.banner.warn{background:var(--warn-bg);border-color:rgba(138,90,0,.3);color:var(--warn)}
.banner.err{background:var(--err-bg);border-color:rgba(192,38,70,.3);color:var(--err)}
.banner ul{margin:5px 0 0 16px;padding:0}

/* ---- tabs ---- */
.tabs{display:flex;gap:2px;flex-wrap:wrap;border-bottom:1px solid var(--line)}
.tab{padding:10px 15px;border:0;background:none;color:var(--muted);font-weight:500;
     font-size:14px;cursor:pointer;border-bottom:2px solid transparent;
     font-family:var(--font);transition:.15s;box-shadow:none;border-radius:0}
.tab:hover{color:var(--ink);background:none;filter:none;transform:none}
.tab.on{color:var(--brand);border-bottom-color:var(--brand)}
.tab .n{background:var(--line-2);color:var(--muted);border-radius:99px;padding:1px 7px;
        font-size:11px;margin-left:5px}
.tab.on .n{background:var(--brand-soft);color:var(--brand)}
.pane{display:none} .pane.on{display:block}

/* ---- tables ---- */
.scroll{overflow:auto;max-height:620px;border:1px solid var(--line);border-radius:9px;
        background:var(--panel);margin-top:14px}
table{border-collapse:separate;border-spacing:0;width:100%;font-size:13.5px}
/* Every cell is one line tall and ruled on all sides, so a figure always sits
   in its own box -- a wrapping cell used to stretch its whole row and leave the
   amounts floating in the middle of the gap. */
th,td{padding:7px 12px;text-align:left;white-space:nowrap;vertical-align:middle;
      height:34px;line-height:20px;
      border-bottom:1px solid var(--line-2);border-right:1px solid var(--line-2)}
th:last-child,td:last-child{border-right:0}
th{position:sticky;top:0;background:var(--panel-2);font-size:11px;text-transform:uppercase;
   letter-spacing:.06em;color:var(--muted);z-index:2;font-weight:700;
   border-bottom:1px solid var(--line);border-right:1px solid var(--line)}
/* Heading text and its filter control share the cell. */
.thin{display:flex;align-items:center;gap:8px;justify-content:space-between}
.fbtn{background:none;border:0;padding:2px;border-radius:5px;cursor:pointer;
   color:var(--muted);opacity:.55;display:grid;place-items:center;box-shadow:none;
   flex:0 0 auto;transition:.15s}
.fbtn svg{width:14px;height:14px;display:block}
.fbtn:hover{opacity:1;color:var(--brand);background:var(--brand-soft);
   transform:none;filter:none}
.fbtn.on{opacity:1;color:var(--brand);background:var(--brand-soft)}
th:has(.fbtn.on){background:var(--brand-soft)}
.fbtn.sorted{opacity:1;color:var(--brand)}
/* One shared popover, positioned over the page so the scroll box can't clip it. */
.fpop{position:fixed;z-index:60;background:var(--panel);border:1px solid var(--line);
   border-radius:11px;box-shadow:0 16px 40px -12px rgba(16,24,40,.34);padding:8px;
   width:272px;font-size:13.5px}
.fpop[hidden]{display:none}
.fsort{display:flex;flex-direction:column;gap:1px;padding-bottom:7px;
   border-bottom:1px solid var(--line-2);margin-bottom:8px}
.fsort button{display:flex;align-items:center;gap:9px;width:100%;height:31px;padding:0 9px;
   background:none;border:0;border-radius:7px;color:var(--ink);font-size:13.5px;
   font-weight:400;text-align:left;cursor:pointer;box-shadow:none;font-family:var(--font)}
.fsort button span{color:var(--muted);font-size:13px;width:12px;text-align:center}
.fsort button:hover{background:var(--brand-soft);color:var(--brand);transform:none;filter:none}
.fsort button:hover span{color:var(--brand)}
#fpop-search{width:100%;height:33px;padding:0 10px;border:1px solid var(--line);
   border-radius:8px;background:var(--panel-2);color:var(--ink);font-size:13.5px;
   font-family:var(--font);margin-bottom:8px}
#fpop-search:focus{outline:0;border-color:var(--brand);
   box-shadow:0 0 0 2px rgba(91,75,245,.16)}
.flist{max-height:216px;overflow-y:auto;border:1px solid var(--line);border-radius:8px;
   padding:4px;background:var(--panel-2);
   /* keep the wheel in this list instead of chaining to the page behind it */
   overscroll-behavior:contain;scrollbar-width:thin;
   scrollbar-color:var(--muted) transparent}
.flist::-webkit-scrollbar{width:10px}
.flist::-webkit-scrollbar-track{background:transparent}
.flist::-webkit-scrollbar-thumb{background:var(--line);border-radius:6px;
   border:2px solid var(--panel-2)}
.flist::-webkit-scrollbar-thumb:hover{background:var(--muted)}
.fitem{display:flex;align-items:center;gap:9px;padding:5px 7px;border-radius:6px;
   cursor:pointer;line-height:1.3}
.fitem:hover{background:var(--brand-soft)}
.fitem input{width:15px;height:15px;flex:0 0 15px;accent-color:var(--brand);margin:0}
.fitem span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.fitem.fall{font-weight:600;border-bottom:1px solid var(--line-2);border-radius:6px 6px 0 0;
   margin-bottom:3px;padding-bottom:7px}
.fnone{padding:16px;text-align:center;color:var(--muted);font-size:13px}
.fpop-foot{display:flex;gap:7px;margin-top:8px}
.fpop-foot button{flex:1;height:33px;font-size:13px;font-weight:500;background:none;
   border:1px solid var(--line);color:var(--muted);border-radius:8px;box-shadow:none}
.fpop-foot button:hover{border-color:var(--brand);color:var(--brand);
   background:var(--brand-soft);transform:none;filter:none}
.fpop-foot button.primary{background:var(--brand);border-color:var(--brand);color:#fff;
   font-weight:600}
.fpop-foot button.primary:hover{filter:brightness(1.08);color:#fff}
.tablewrap{margin-top:14px}
.tablewrap .scroll{margin-top:0}  /* the wrapper owns the spacing */
.tbar{display:flex;align-items:center;gap:12px;justify-content:flex-end;
   margin-bottom:7px;font-size:12.5px;color:var(--muted)}
.fclear{background:none;border:1px solid var(--line);color:var(--muted);border-radius:7px;
   padding:4px 11px;font-size:12.5px;font-weight:500;cursor:pointer;box-shadow:none}
.fclear:hover{border-color:var(--brand);color:var(--brand);background:var(--brand-soft);
   transform:none;filter:none}
/* A source you can go and check reads as a link, not as a label. */
.srclink{color:var(--brand);font-weight:500;text-decoration:underline;
   text-decoration-style:dotted;text-underline-offset:3px}
.srclink:hover{text-decoration-style:solid}
tbody tr:nth-child(even){background:color-mix(in srgb,var(--line-2) 45%,transparent)}
tbody tr:hover{background:var(--brand-soft)}
td.num{text-align:right;font-variant-numeric:tabular-nums;font-family:var(--mono);font-size:13px}
tr.total td{font-weight:700;background:var(--panel-2)}
tr.vstart td{border-top:2px solid var(--line)}
/* Long narrations are clipped to the column and shown in full on hover, rather
   than being allowed to set the row height for every other column. */
td.narr{color:var(--muted);max-width:340px;overflow:hidden;text-overflow:ellipsis}
.pill{display:inline-block;padding:2px 9px;border-radius:99px;font-size:11px;font-weight:700}
.pill.dr{background:var(--ok-bg);color:var(--ok)}
.pill.cr{background:var(--warn-bg);color:var(--warn)}
.pill.Sales{background:rgba(11,116,196,.14);color:var(--sky)}
.pill.JE{background:var(--brand-soft);color:var(--brand)}
.pill.ERROR{background:var(--err-bg);color:var(--err)}
.pill.INFO{background:var(--warn-bg);color:var(--warn)}
/* ---- add-a-row forms under the master tables ---- */
.addrow{margin-top:14px;padding:15px 16px;border:1px dashed var(--line);border-radius:11px;
        background:var(--panel-2)}
.addhead{font-size:13px;color:var(--muted);margin-bottom:11px}
.addgrid{display:grid;gap:11px;align-items:end;grid-template-columns:1fr 1fr auto}
.addgrid.three{grid-template-columns:repeat(3,1fr) auto}
.addgrid.four{grid-template-columns:repeat(4,1fr) auto}
.addgrid.five{grid-template-columns:repeat(5,1fr) auto}
.addgrid button{white-space:nowrap}
@media (max-width:900px){.addgrid,.addgrid.five{grid-template-columns:1fr}}

/* ---- per-row edit ---- */
th.actcol,td.actcol{width:1%;white-space:nowrap;text-align:center}
.editbtn{background:none;border:1px solid var(--line);color:var(--muted);border-radius:7px;
   padding:3px 11px;font-size:12.5px;font-weight:500;cursor:pointer;box-shadow:none;
   font-family:var(--font)}
.editbtn:hover{border-color:var(--brand);color:var(--brand);background:var(--brand-soft);
   transform:none;filter:none}
.modal{position:fixed;inset:0;z-index:80;background:rgba(15,18,32,.45);
   display:grid;place-items:center;padding:20px}
.modal[hidden]{display:none}
.modal-box{background:var(--panel);border:1px solid var(--line);border-radius:14px;
   padding:24px;width:100%;max-width:560px;box-shadow:0 24px 60px -20px rgba(16,24,40,.5)}
.modal-box h3{margin:0 0 4px;font-size:17px;font-weight:700;letter-spacing:-.01em}
.modal-box .sub{color:var(--muted);font-size:13.5px;margin:0 0 18px}
.modal-grid{display:grid;gap:13px;grid-template-columns:1fr 1fr}
.modal-grid .wide{grid-column:1 / -1}
.modal-foot{display:flex;gap:9px;margin-top:22px;align-items:center}
.modal-foot .spacer{margin-left:auto}
.btn-danger{background:none;border:1px solid var(--err);color:var(--err);box-shadow:none;
   font-weight:500;padding:10px 16px;border-radius:9px;font-size:14px;cursor:pointer;
   font-family:var(--font)}
.btn-danger:hover{background:var(--err-bg);filter:none;transform:none}

code{font-family:var(--mono);font-size:12.5px;background:var(--panel-2);padding:2px 5px;
     border-radius:4px;border:1px solid var(--line)}
.empty{padding:28px;color:var(--muted);font-size:14px;text-align:center}

@media (max-width:1080px){ .steps{grid-template-columns:1fr} }
@media (max-width:860px){
  .shell{flex-direction:column}
  .side{width:100%;flex:none;height:auto;position:static}
  .main{padding:0 18px 40px}
}
"""

#: NB: both CSS and JS are injected with |safe. <style> and <script> are raw-text
#: elements, so Jinja's default escaping would turn a quote into &#34; and leave
#: it undecoded -- silently breaking every quoted font-family and string literal.
#: Runs in <head>, before anything paints, so a reader who chose dark never sees
#: a white flash while the rest of the page loads.
THEME_BOOT = ("<script>try{var t=localStorage.getItem('vervi-theme');"
              "if(t)document.documentElement.setAttribute('data-theme',t);}catch(e){}</script>")

#: Leading the results top bar: the way back to a new run. Always visible, so
#: nobody has to hunt for it or reload the page by hand.
BACK_BTN = """<a class="backbtn" href="{{ url_for('index') }}">
    <span aria-hidden="true">&larr;</span> New journal</a>"""

#: Rendered into the top-right of every signed-in page.
TOPRIGHT = """<div class="topright">
  <button type="button" class="iconbtn" onclick="toggleTheme()"
          title="Switch between light and dark" aria-label="Switch theme">
    <span class="moon">&#9789;</span><span class="sun">&#9788;</span>
  </button>
  <a class="signout" href="{{ url_for('logout') }}">
    <span>Sign out</span><span class="who">{{ session.get('user','admin') }}</span>
  </a>
</div>"""

TAB_JS = """
// Theme choice is remembered per browser; with nothing chosen the page follows
// the operating system.
function applyTheme(t){
  if (t === 'dark' || t === 'light') { document.documentElement.setAttribute('data-theme', t); }
  else { document.documentElement.removeAttribute('data-theme'); }
}
function currentTheme(){
  var set = document.documentElement.getAttribute('data-theme');
  if (set) return set;
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}
function toggleTheme(){
  var next = currentTheme() === 'dark' ? 'light' : 'dark';
  applyTheme(next);
  try { localStorage.setItem('vervi-theme', next); } catch (e) {}
}
try { applyTheme(localStorage.getItem('vervi-theme')); } catch (e) {}

// Per-column filtering. Terms are held per table so they survive the popover
// being closed and reopened; a row survives only if every term matches its own
// column. Case-insensitive substring, so "IGST" or "07/07" both work.
var FILTERS = {};
var POP = {table: null, col: -1, values: []};

// ---- editing a master row -------------------------------------------------
// Values come from data- attributes on the button, so the dialog always shows
// exactly what is on the row rather than a second copy that could drift.
function openEdit(btn){
  var d = btn.dataset;
  var modal = document.getElementById('editmodal');
  var isAccount = d.kind === 'account';
  document.getElementById('em-title').textContent =
    isAccount ? 'Edit account' : 'Edit treatment rule';
  document.getElementById('em-sub').textContent =
    'Saving rewrites this row in the master database.';
  document.getElementById('em-key').value = d.key || '';
  document.getElementById('em-form').action =
    isAccount ? '/master/account/update' : '/master/treatment/update';
  document.getElementById('em-account').style.display = isAccount ? '' : 'none';
  document.getElementById('em-treatment').style.display = isAccount ? 'none' : '';

  if (isAccount){
    document.getElementById('em-a-name').value = d.account || '';
    document.getElementById('em-a-gl').value = d.gl || '';
    document.getElementById('em-a-cc').value = d.cc || '';
    document.getElementById('em-a-sub').value = d.sub || '';
  } else {
    document.getElementById('em-t-nat').value = d.nature || '';
    document.getElementById('em-t-d1').value = d.d1 || '';
    document.getElementById('em-t-d2').value = d.d2 || '';
    document.getElementById('em-t-c1').value = d.c1 || '';
    document.getElementById('em-t-c2').value = d.c2 || '';
  }

  var del = document.getElementById('em-delete');
  del.onclick = function(){
    var what = isAccount ? (d.account || 'this account') : (d.nature || 'this rule');
    if (!confirm('Remove ' + what + ' from the master database?')) return;
    var f = document.getElementById('em-form');
    f.action = isAccount ? '/master/account/delete' : '/master/treatment/delete';
    f.submit();
  };

  modal.hidden = false;
  (isAccount ? document.getElementById('em-a-name')
             : document.getElementById('em-t-nat')).focus();
}
function closeEdit(){
  var m = document.getElementById('editmodal');
  if (m) m.hidden = true;
}
document.addEventListener('keydown', function(e){
  if (e.key === 'Escape') closeEdit();
});
document.addEventListener('click', function(e){
  // Clicking the dimmed backdrop, but not the dialog itself, dismisses it.
  if (e.target && e.target.id === 'editmodal') closeEdit();
});

function cellText(row, col){
  var c = row.cells[col];
  return c ? c.textContent.trim() : '';
}

// Body rows that count as data -- a pinned totals row is a summary, not a row.
function dataRows(table){
  return Array.prototype.filter.call(table.tBodies[0].rows, function(r){
    return !r.classList.contains('keep');
  });
}

function popEl(){
  var el = document.getElementById('fpop');
  if (el) return el;
  el = document.createElement('div');
  el.id = 'fpop';
  el.className = 'fpop';
  el.hidden = true;
  el.innerHTML =
      '<div class="fsort">'
    + '  <button type="button" data-dir="asc"><span>\\u2191</span> Sort A to Z</button>'
    + '  <button type="button" data-dir="desc"><span>\\u2193</span> Sort Z to A</button>'
    + '</div>'
    + '<input type="text" id="fpop-search" placeholder="Search\\u2026" aria-label="Search values">'
    + '<div class="flist" id="fpop-list"></div>'
    + '<div class="fpop-foot">'
    + '  <button type="button" id="fpop-clr">Clear</button>'
    + '  <button type="button" id="fpop-ok" class="primary">Apply</button>'
    + '</div>';
  document.body.appendChild(el);

  el.querySelector('#fpop-search').addEventListener('input', function(){
    renderValues(this.value);
  });
  el.querySelectorAll('.fsort button').forEach(function(b){
    b.addEventListener('click', function(){ sortByColumn(POP.table, POP.col, b.dataset.dir); });
  });
  el.querySelector('#fpop-clr').addEventListener('click', function(){
    if (!POP.table) return;
    delete (FILTERS[POP.table] || {})[POP.col];
    applyFilters(POP.table);
    closeFilter();
  });
  el.querySelector('#fpop-ok').addEventListener('click', commitSelection);
  el.addEventListener('keydown', function(e){
    if (e.key === 'Escape') closeFilter();
    if (e.key === 'Enter') commitSelection();
  });
  return el;
}

// Distinct values in a column, numerically where the column holds numbers.
function distinctValues(table, col){
  var seen = {}, out = [];
  dataRows(table).forEach(function(r){
    var t = cellText(r, col);
    if (!(t in seen)){ seen[t] = 1; out.push(t); }
  });
  var numeric = out.every(function(v){
    return v === '' || !isNaN(parseFloat(v.replace(/,/g, '')));
  });
  out.sort(function(a, b){
    if (a === '') return -1;
    if (b === '') return 1;
    return numeric
      ? parseFloat(a.replace(/,/g, '')) - parseFloat(b.replace(/,/g, ''))
      : a.localeCompare(b, undefined, {numeric: true, sensitivity: 'base'});
  });
  return out;
}

function renderValues(search){
  var el = popEl();
  var list = el.querySelector('#fpop-list');
  var table = document.getElementById(POP.table);
  if (!table) return;
  var chosen = (FILTERS[POP.table] || {})[POP.col] || null;  // null = everything
  var needle = (search || '').trim().toLowerCase();
  var values = POP.values.filter(function(v){
    return !needle || v.toLowerCase().indexOf(needle) !== -1;
  });

  var allOn = values.every(function(v){ return !chosen || chosen.indexOf(v) !== -1; });
  var html = '<label class="fitem fall"><input type="checkbox" data-all="1"'
           + (allOn ? ' checked' : '') + '><span>(Select all)</span></label>';
  values.forEach(function(v){
    var on = !chosen || chosen.indexOf(v) !== -1;
    html += '<label class="fitem"><input type="checkbox" value="' + v.replace(/"/g, '&quot;')
          + '"' + (on ? ' checked' : '') + '><span>'
          + (v === '' ? '<em>(blank)</em>' : v.replace(/[<>&]/g, function(ch){
              return {'<': '&lt;', '>': '&gt;', '&': '&amp;'}[ch]; }))
          + '</span></label>';
  });
  if (!values.length) html = '<div class="fnone">No matching values</div>';
  list.innerHTML = html;

  var all = list.querySelector('input[data-all]');
  if (all){
    all.addEventListener('change', function(){
      list.querySelectorAll('input:not([data-all])').forEach(function(b){ b.checked = all.checked; });
    });
  }
  list.querySelectorAll('input:not([data-all])').forEach(function(b){
    b.addEventListener('change', function(){
      if (!all) return;
      var boxes = list.querySelectorAll('input:not([data-all])');
      all.checked = Array.prototype.every.call(boxes, function(x){ return x.checked; });
    });
  });
}

function commitSelection(){
  var el = popEl();
  var list = el.querySelector('#fpop-list');
  var boxes = list.querySelectorAll('input:not([data-all])');
  if (!boxes.length){ closeFilter(); return; }

  // Start from what is already allowed, so narrowing by search and ticking a few
  // does not silently discard choices scrolled out of view.
  var chosen = (FILTERS[POP.table] || {})[POP.col];
  var allowed = chosen ? chosen.slice() : POP.values.slice();
  boxes.forEach(function(b){
    var i = allowed.indexOf(b.value);
    if (b.checked && i === -1) allowed.push(b.value);
    if (!b.checked && i !== -1) allowed.splice(i, 1);
  });

  FILTERS[POP.table] = FILTERS[POP.table] || {};
  if (allowed.length >= POP.values.length) delete FILTERS[POP.table][POP.col];
  else FILTERS[POP.table][POP.col] = allowed;
  applyFilters(POP.table);
  closeFilter();
}

function sortByColumn(tableId, col, dir){
  var table = document.getElementById(tableId);
  if (!table) return;
  var body = table.tBodies[0];
  var rows = dataRows(table);
  var pinned = Array.prototype.filter.call(body.rows, function(r){
    return r.classList.contains('keep');
  });
  var numeric = rows.every(function(r){
    var t = cellText(r, col);
    return t === '' || !isNaN(parseFloat(t.replace(/,/g, '')));
  });
  rows.sort(function(a, b){
    var x = cellText(a, col), y = cellText(b, col);
    var n = numeric
      ? (parseFloat(x.replace(/,/g, '')) || 0) - (parseFloat(y.replace(/,/g, '')) || 0)
      : x.localeCompare(y, undefined, {numeric: true, sensitivity: 'base'});
    return dir === 'desc' ? -n : n;
  });
  // Voucher rules describe the original order, so they stop being true once sorted.
  rows.forEach(function(r){ r.classList.remove('vstart'); body.appendChild(r); });
  pinned.forEach(function(r){ body.appendChild(r); });
  table.querySelectorAll('thead .fbtn').forEach(function(b, i){
    b.classList.toggle('sorted', i === col);
  });
  closeFilter();
}

function openFilter(btn){
  var th = btn.closest('th');
  var table = btn.closest('table');
  var col = Array.prototype.indexOf.call(th.parentNode.children, th);
  var el = popEl();
  if (!el.hidden && POP.table === table.id && POP.col === col){ closeFilter(); return; }

  POP.table = table.id;
  POP.col = col;
  POP.values = distinctValues(table, col);
  el.querySelector('#fpop-search').value = '';
  renderValues('');
  el.hidden = false;

  var r = btn.getBoundingClientRect();
  var w = el.offsetWidth || 270;
  var h = el.offsetHeight || 340;
  el.style.left = Math.round(Math.min(Math.max(8, r.right - w), window.innerWidth - w - 8)) + 'px';
  // Flip above the heading when there is no room below it.
  el.style.top = (r.bottom + h + 10 > window.innerHeight && r.top - h - 6 > 0)
    ? Math.round(r.top - h - 6) + 'px'
    : Math.round(r.bottom + 6) + 'px';
  el.querySelector('#fpop-search').focus();
}

function closeFilter(){
  var el = document.getElementById('fpop');
  if (el) el.hidden = true;
  POP.table = null; POP.col = -1; POP.values = [];
}

function applyFilters(tableId){
  var table = document.getElementById(tableId);
  if (!table) return;
  var sets = FILTERS[tableId] || {};
  var rows = table.tBodies[0].rows, shown = 0, counted = 0;
  for (var r = 0; r < rows.length; r++){
    var row = rows[r];
    if (row.classList.contains('keep')){ row.style.display = ''; continue; }
    counted++;
    var ok = true;
    for (var c in sets){
      if (sets[c].indexOf(cellText(row, c)) === -1){ ok = false; break; }
    }
    row.style.display = ok ? '' : 'none';
    if (ok) shown++;
  }
  table.querySelectorAll('thead .fbtn').forEach(function(b, i){
    b.classList.toggle('on', Object.prototype.hasOwnProperty.call(sets, String(i)));
  });
  var label = document.getElementById(tableId + '-count');
  if (label){
    label.textContent = (shown === counted) ? counted + ' rows'
                                            : shown + ' of ' + counted + ' rows';
  }
}

function clearFilters(id){
  FILTERS[id] = {};
  applyFilters(id);
  var table = document.getElementById(id);
  if (table) table.querySelectorAll('thead .fbtn').forEach(function(b){
    b.classList.remove('sorted');
  });
  closeFilter();
}

function insidePopover(node){
  return node && node.nodeType === 1 && node.closest && !!node.closest('#fpop');
}

document.addEventListener('click', function(e){
  var el = document.getElementById('fpop');
  if (!el || el.hidden) return;
  if (!insidePopover(e.target) && !e.target.closest('.fbtn')) closeFilter();
});
window.addEventListener('resize', closeFilter);

// The popover is anchored to a heading, so a scroll of the page or the table
// moves it out of place and it should close. Scrolling the value list inside it
// must NOT close it -- that is the list doing its job.
document.addEventListener('scroll', function(e){
  var el = document.getElementById('fpop');
  if (!el || el.hidden) return;
  if (insidePopover(e.target)) return;
  closeFilter();
}, true);

// Keep the wheel inside the popover: scroll the value list, and never let the
// page scroll underneath while the pointer is over the panel.
document.addEventListener('wheel', function(e){
  var el = document.getElementById('fpop');
  if (!el || el.hidden || !insidePopover(e.target)) return;
  var list = document.getElementById('fpop-list');
  if (list && list.contains(e.target)){
    var atTop = list.scrollTop <= 0;
    var atEnd = list.scrollTop + list.clientHeight >= list.scrollHeight - 1;
    // Only block the page once the list itself has nowhere left to go.
    if ((e.deltaY < 0 && atTop) || (e.deltaY > 0 && atEnd)) e.preventDefault();
    return;
  }
  e.preventDefault();
}, {passive: false});

function showTab(btn, id){
  var root = btn.closest('.tabwrap');
  root.querySelectorAll('.tab').forEach(function(t){ t.classList.remove('on'); });
  root.querySelectorAll('.pane').forEach(function(p){ p.classList.remove('on'); });
  btn.classList.add('on');
  root.querySelector('#'+id).classList.add('on');
}
document.addEventListener('DOMContentLoaded', function(){
  document.querySelectorAll('.drop').forEach(function(zone){
    var input = zone.querySelector('input[type=file]');
    var label = zone.querySelector('.t');
    var original = label.textContent;
    input.addEventListener('change', function(){
      if (input.files.length){ label.textContent = input.files[0].name; zone.classList.add('has'); }
      else { label.textContent = original; zone.classList.remove('has'); }
    });
    ['dragenter','dragover'].forEach(function(e){
      zone.addEventListener(e, function(ev){ ev.preventDefault(); zone.classList.add('over'); }); });
    ['dragleave','drop'].forEach(function(e){
      zone.addEventListener(e, function(ev){ ev.preventDefault(); zone.classList.remove('over'); }); });
  });
});
"""

#: The sidebar carries the identity and every standing fact, which is what keeps
#: the top of the page down to a title and one line of context.
#: The Vervi mark -- an open book, matching Vervi-Books so the two products are
#: recognisably one family. Inlined rather than linked so it needs no network.
BOOK_ICON = (
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" '
    'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
    '<path d="M2 4.5h6a3 3 0 0 1 3 3V20a2.5 2.5 0 0 0-2.5-2.5H2z"/>'
    '<path d="M22 4.5h-6a3 3 0 0 0-3 3V20a2.5 2.5 0 0 1 2.5-2.5H22z"/></svg>')

MAIL_ICON = (
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" '
    'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
    '<rect x="2.5" y="4.5" width="19" height="15" rx="2.5"/>'
    '<path d="m3 7 9 6 9-6"/></svg>')

#: The little "narrow this column" glyph that sits beside each heading.
FILTER_ICON = (
    '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" '
    'stroke-linecap="round" aria-hidden="true">'
    '<path d="M2.5 4.5h11M4.5 8h7M6.5 11.5h3"/></svg>')

LOCK_ICON = (
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" '
    'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
    '<rect x="4" y="10.5" width="16" height="10" rx="2.5"/>'
    '<path d="M8 10.5V7a4 4 0 0 1 8 0v3.5"/></svg>')


SIDEBAR = """<aside class="side">
  <div class="logo">
    <div class="mark">""" + BOOK_ICON + """</div>
    <div><div class="nm">""" + APP_NAME + """</div>
         <div class="co">""" + COMPANY + """</div></div>
  </div>
  {% block sidebody %}{% endblock %}
  <div class="sblock c2">
    <div class="slabel">History</div>
    <div class="srow"><span class="k">Statements converted</span>
      <span class="v big">{{ n_history }}</span></div>
    <div class="row" style="margin-top:8px">
      <a href="{{ url_for('history') }}" style="width:100%">
        <button type="button" class="sidealt" style="width:100%">See converted files</button></a>
    </div>
  </div>
  <div class="sblock c4">
    <div class="slabel">How to import</div>
    <ol class="steps-list">
      <li>On Upwork: <strong>Reports &rarr; Transaction History</strong></li>
      <li>Pick the period, then <strong>Download CSV</strong></li>
      <li>Drop that file into <strong>Upload</strong></li>
      <li>Check every account appears below</li>
      <li>Hit <strong>Generate journal</strong></li>
      <li>Review <strong>Exceptions</strong> &mdash; should be 0</li>
      <li>Download the <strong>CSV</strong> &mdash; that locks the numbers</li>
    </ol>
  </div>
</aside>"""

#: One dialog serves both tabs; JS shows whichever fieldset applies.
EDIT_MODAL = """
<div class="modal" id="editmodal" hidden>
  <div class="modal-box" role="dialog" aria-modal="true" aria-labelledby="em-title">
    <h3 id="em-title">Edit</h3>
    <p class="sub" id="em-sub"></p>

    <form method="post" id="em-form">
      <input type="hidden" name="key" id="em-key">

      <div class="modal-grid" id="em-account">
        <div><label for="em-a-name">Account name (col G)</label>
          <input type="text" id="em-a-name" name="account"></div>
        <div><label for="em-a-gl">GL name (col H)</label>
          <input type="text" id="em-a-gl" name="gl_name"></div>
        <div><label for="em-a-cc">Cost centre (col I)</label>
          <input type="text" id="em-a-cc" name="cost_center"></div>
        <div><label for="em-a-sub">Subsidiary (col J)</label>
          <input type="text" id="em-a-sub" name="subsidiary"></div>
        <div class="wide hint">A blank cost centre falls back to the statement's
          Freelancer; a blank subsidiary falls back to the run's entity.</div>
      </div>

      <div class="modal-grid" id="em-treatment">
        <div class="wide"><label for="em-t-nat">Transaction type</label>
          <input type="text" id="em-t-nat" name="nature"></div>
        <div><label for="em-t-d1">Debit 1</label>
          <input type="text" id="em-t-d1" name="debit_1"></div>
        <div><label for="em-t-d2">Debit 2</label>
          <input type="text" id="em-t-d2" name="debit_2"></div>
        <div><label for="em-t-c1">Credit 1</label>
          <input type="text" id="em-t-c1" name="credit_1"></div>
        <div><label for="em-t-c2">Credit 2</label>
          <input type="text" id="em-t-c2" name="credit_2"></div>
        <div class="wide hint">Use <code>GL Name</code> for the wallet,
          <code>Client Team</code> for the client, <code>NA</code> for an unused leg.</div>
      </div>

      <div class="modal-foot">
        <button type="button" class="btn-danger" id="em-delete">Delete</button>
        <span class="spacer"></span>
        <button type="button" class="ghost" onclick="closeEdit()">Cancel</button>
        <button type="submit">Save changes</button>
      </div>
    </form>
  </div>
</div>"""

FORM_SIDE = """
  <div class="sblock c1">
    <div class="slabel">Master database</div>
    <div class="srow"><span class="k">Accounts</span><span class="v">{{ n_accounts }}</span></div>
    <div class="srow"><span class="k">Treatments</span><span class="v">{{ n_treatments }}</span></div>
    <div class="snote">Ledgers are read from <code>{{ master_path }}</code>.
      Add a row there to add a freelancer.</div>
  </div>
  <div class="sblock c2">
    <div class="slabel">Defaults</div>
    <div class="srow"><span class="k">Currency</span><span class="v">INR</span></div>
    <div class="srow"><span class="k">IGST</span><span class="v">18%</span></div>
    <div class="srow"><span class="k">Rates</span><span class="v">RBI, per date</span></div>
    <div class="srow"><span class="k">Sales</span><span class="v">2 entries</span></div>
  </div>
  <div class="sblock c3">
    <div class="slabel">Document numbers</div>
    {% if reg.ok %}
    <div class="srow"><span class="k">Last Sales used</span>
      <span class="v">{{ reg.sales_last or '&mdash;'|safe }}</span></div>
    <div class="srow"><span class="k">Next Sales</span>
      <span class="v big">{{ reg.sales_next }}</span></div>
    <div class="srow"><span class="k">Last JE used ({{ reg.je_month }})</span>
      <span class="v">{{ reg.je_last or '&mdash;'|safe }}</span></div>
    <div class="srow"><span class="k">Next JE</span>
      <span class="v big">{{ reg.je_next }}</span></div>
    {% endif %}
    <div class="snote">Numbering continues from the last one you actually
      exported &mdash; a new statement never reuses a number.</div>
  </div>
"""

FORM_HTML = """<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>""" + APP_NAME + """</title>""" + FONT_LINK + """<style>{{ css|safe }}</style>
""" + THEME_BOOT + """</head><body>
<div class="shell">""" + SIDEBAR.replace(
    "{% block sidebody %}{% endblock %}", FORM_SIDE) + """
<main class="main">
  <div class="top"><h1>New journal</h1>
    <span class="s">Upwork statement &rarr; import-ready entries</span>
    """ + TOPRIGHT + """</div>

  {% if error %}<div class="banner err"><span>&#9888;</span>
  <div><strong>Couldn't do that.</strong> {{ error }}</div></div>{% endif %}
  {% if notice %}<div class="banner ok"><span>&#10004;</span>
  <div><strong>Master database updated.</strong> {{ notice }}</div></div>{% endif %}
  {% if master_error %}<div class="banner err"><span>&#9888;</span>
  <div><strong>Master database problem.</strong> {{ master_error }}</div></div>{% endif %}

  <form method="post" action="{{ url_for('run') }}" enctype="multipart/form-data">
    <div class="steps">
      <div class="panel">
        <div class="step"><span class="n">1</span><h2>Upload the statement</h2></div>
        <p class="psub">Your Upwork account statement is the only file needed.</p>
        <div class="drop">
          <input type="file" name="statement" accept=".csv,.xlsx,.xlsm,.xls">
          <div class="big">&#128196;</div>
          <div>
            <div class="t">Choose a file or drag it here</div>
            <div class="s">.csv or .xlsx</div>
          </div>
        </div>
      </div>

      <div class="panel">
        <div class="step"><span class="n">2</span><h2>Generate</h2></div>
        <p class="psub">Everything is already set to your standard treatment.</p>
        <div class="row" style="margin-top:0">
          <button type="submit">Generate journal</button>
          {% if samples_exist %}
          <button type="submit" name="use_sample" value="1" class="ghost">Try a sample</button>
          {% endif %}
        </div>
      </div>
    </div>

    <div class="panel" style="padding-top:6px">
      <details {{ 'open' if error }}>
        <summary>Change settings</summary>
        <div class="grid" style="margin-top:14px">
          <div><label>Currency</label>
            <select name="currency">
              <option value="INR" {{ 'selected' if currency=='INR' }}>INR (convert)</option>
              <option value="USD" {{ 'selected' if currency=='USD' }}>USD (as-is)</option>
            </select></div>
          <div><label>Exchange rate</label>
            <select name="fx_source">
              <option value="rbi" {{ 'selected' if fx_source=='rbi' }}>RBI, per date</option>
              <option value="manual" {{ 'selected' if fx_source=='manual' }}>Manual flat rate</option>
            </select></div>
          <div><label>Manual rate (fallback)</label>
            <input type="text" name="fx_rate" value="{{ fx_rate }}" placeholder="optional"></div>
          <div><label>IGST %</label>
            <input type="text" name="igst_rate" value="{{ igst_rate }}"></div>
          <div><label>Sales entries</label>
            <select name="income_mode">
              <option value="two-entry" {{ 'selected' if income_mode=='two-entry' }}>Two &mdash; Sales + JE</option>
              <option value="combined" {{ 'selected' if income_mode=='combined' }}>One combined</option>
            </select></div>
          <div><label>Subsidiary</label>
            <input type="text" name="entity" value="{{ entity }}"></div>
          <div><label>JE document series</label>
            <input type="text" name="doc_je" value="{{ doc_je }}"></div>
          <div><label>Sales document series</label>
            <input type="text" name="doc_sales" value="{{ doc_sales }}"></div>
          <div><label>I/E flag</label>
            <input type="text" name="ie_flag" value="{{ ie_flag }}"></div>
        </div>
        <label class="switch"><input type="checkbox" name="fx_offline" {{ 'checked' if fx_offline }}>
          Offline &mdash; use cached rates only</label>
        <label class="switch"><input type="checkbox" name="reset_docs" {{ 'checked' if reset_docs }}>
          Restart document numbering at 001</label>
        <label class="switch"><input type="checkbox" name="ignore_posted" {{ 'checked' if ignore_posted }}>
          Import everything, even rows posted before (ignores Ref ID history)</label>
      </details>
    </div>
  </form>

  {% if not master_error %}
  <div class="panel tabwrap">
    <details class="fold" {{ 'open' if notice }}>
      <summary>Master database &mdash; {{ n_accounts }} accounts, {{ n_treatments }} treatments</summary>
    <div class="tabs">
      <button type="button" class="tab on" onclick="showTab(this,'m-acc')">Accounts<span class="n">{{ n_accounts }}</span></button>
      <button type="button" class="tab" onclick="showTab(this,'m-trt')">Treatments<span class="n">{{ n_treatments }}</span></button>
    </div>
    <div id="m-acc" class="pane on">
      {{ wallet_table|safe }}
      <form class="addrow" method="post" action="{{ url_for('add_account') }}">
        <div class="addhead">Add an account &mdash; appended to the end of the master.
          Leave the cost centre blank to fall back to the statement's Freelancer.</div>
        <div class="addgrid four">
          <div><label for="a-name">Account name (col G)</label>
            <input type="text" id="a-name" name="account" placeholder="as it appears on the statement" required></div>
          <div><label for="a-gl">GL name (col H)</label>
            <input type="text" id="a-gl" name="gl_name" placeholder="Upwork ..." required></div>
          <div><label for="a-cc">Cost centre (col I)</label>
            <input type="text" id="a-cc" name="cost_center" placeholder="e.g. TPT-Badshah"></div>
          <div><label for="a-sub">Subsidiary (col J)</label>
            <input type="text" id="a-sub" name="subsidiary" value="{{ entity }}"></div>
          <button type="submit" class="ghost">Add account</button>
        </div>
      </form>
    </div>
    <div id="m-trt" class="pane">
      {{ treatment_table|safe }}
      <form class="addrow" method="post" action="{{ url_for('add_rule') }}">
        <div class="addhead">Add a rule &mdash; appended to the end of the master.
          Use <code>GL Name</code> for the wallet, <code>Client Team</code> for the
          client, <code>NA</code> for an unused leg.</div>
        <div class="addgrid five">
          <div><label for="t-nat">Transaction type</label>
            <input type="text" id="t-nat" name="nature" placeholder="e.g. Bonus" required></div>
          <div><label for="t-d1">Debit 1</label>
            <input type="text" id="t-d1" name="debit_1" required></div>
          <div><label for="t-d2">Debit 2</label>
            <input type="text" id="t-d2" name="debit_2" placeholder="NA"></div>
          <div><label for="t-c1">Credit 1</label>
            <input type="text" id="t-c1" name="credit_1" value="GL Name" required></div>
          <div><label for="t-c2">Credit 2</label>
            <input type="text" id="t-c2" name="credit_2" placeholder="NA"></div>
          <button type="submit" class="ghost">Add rule</button>
        </div>
      </form>
    </div>
    </details>
  </div>
  {% endif %}
</main></div>
""" + EDIT_MODAL + """
<script>""" + TAB_JS + """</script></body></html>"""

VERIFY_HTML = """<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Rate check &mdash; {{ rate_date }}</title>""" + FONT_LINK + """
<style>{{ css|safe }}
.vwrap{max-width:760px;margin:0 auto;padding:34px 22px 60px}
.vhead{display:flex;align-items:center;gap:13px;margin-bottom:6px}
.vhead .mark{width:40px;height:40px;border-radius:11px;flex:0 0 40px;color:#fff;
  background:linear-gradient(145deg,#6a5cf0,#7b4ef5);display:grid;place-items:center}
.vhead .mark svg{width:21px;height:21px}
.vhead h1{font-size:21px;margin:0;font-weight:700;letter-spacing:-.02em}
.vsub{color:var(--muted);font-size:14px;margin:0 0 22px 53px}
.verdict{border-radius:12px;padding:17px 19px;margin-bottom:20px;border:1px solid;
  display:flex;gap:13px;align-items:flex-start;font-size:15px}
.verdict .big{font-size:20px;line-height:1.2}
.verdict.match{background:var(--ok-bg);border-color:var(--ok);color:var(--ok)}
.verdict.differ{background:var(--err-bg);border-color:var(--err);color:var(--err)}
.verdict.unknown{background:var(--warn-bg);border-color:var(--warn);color:var(--warn)}
.vrow{display:flex;justify-content:space-between;gap:16px;padding:11px 0;
  border-bottom:1px solid var(--line-2);font-size:14.5px}
.vrow:last-child{border-bottom:0}
.vrow .k{color:var(--muted)}
.vrow .v{font-weight:600;font-variant-numeric:tabular-nums;text-align:right}
.vrow .v.mono{font-family:var(--mono)}
.repro{background:var(--panel-2);border:1px solid var(--line);border-radius:11px;
  padding:15px 17px;margin-top:18px;font-size:13.5px;color:var(--muted)}
.repro dl{display:grid;grid-template-columns:auto 1fr;gap:5px 14px;margin:11px 0 0}
.repro dt{font-family:var(--mono);font-size:12.5px;color:var(--ink)}
.repro dd{margin:0;font-family:var(--mono);font-size:12.5px}
.vback{display:inline-block;margin-top:22px}
</style>""" + THEME_BOOT + """</head><body>
<div class="vwrap">
  <div class="vhead"><div class="mark">""" + BOOK_ICON + """</div>
    <h1>Exchange rate check</h1></div>
  <p class="vsub">RBI USD/INR reference rate for <strong>{{ rate_date_long }}</strong>,
     fetched from the archive just now.</p>

  {% if verdict == 'match' %}
  <div class="verdict match"><span class="big">&#10004;</span>
    <div><strong>Verified.</strong> The rate used in the journal is exactly what RBI
    publishes for {{ rate_date_long }}.</div></div>
  {% elif verdict == 'differ' %}
  <div class="verdict differ"><span class="big">&#9888;</span>
    <div><strong>Does not match.</strong> The journal used {{ used }} but RBI publishes
    {{ published }} for {{ rate_date_long }}. Re-run with a cleared cache
    (<code>fx_cache.csv</code>) before filing.</div></div>
  {% elif verdict == 'nopublish' %}
  <div class="verdict unknown"><span class="big">&#8213;</span>
    <div><strong>Nothing published.</strong> RBI has no reference rate for
    {{ rate_date_long }} &mdash; a weekend or bank holiday. A rate carried forward from
    an earlier day is the expected treatment.</div></div>
  {% else %}
  <div class="verdict unknown"><span class="big">&#9888;</span>
    <div><strong>Couldn't reach RBI.</strong> {{ note }} The journal figure is shown
    below from the local cache; try again in a moment to confirm it against the source.</div></div>
  {% endif %}

  <div class="panel">
    <div class="vrow"><span class="k">Rate date checked</span>
      <span class="v mono">{{ rate_date }}</span></div>
    {% if txn_date and txn_date != rate_date %}
    <div class="vrow"><span class="k">Applied to transactions dated</span>
      <span class="v mono">{{ txn_date }}</span></div>
    <div class="vrow"><span class="k">Carried forward by</span>
      <span class="v">{{ days_back }} day{{ '' if days_back == 1 else 's' }}
      &mdash; no rate published on {{ txn_date }}</span></div>
    {% endif %}
    <div class="vrow"><span class="k">Rate used in the journal</span>
      <span class="v mono">{{ used or '—' }}</span></div>
    <div class="vrow"><span class="k">Published by RBI (live)</span>
      <span class="v mono">{{ published or '—' }}</span></div>
    <div class="vrow"><span class="k">In the local cache</span>
      <span class="v mono">{{ cached or 'not cached' }}</span></div>
  </div>

  <div class="repro">
    RBI's archive is a form, not a page you can link to a date &mdash; it only answers a
    POST. These are the exact values submitted, so you can reproduce this by hand at
    <a href="{{ archive_url }}" target="_blank" rel="noopener">the Reference Rate Archive</a>:
    <dl>
      <dt>From Date</dt><dd>{{ ddmmyyyy }}</dd>
      <dt>To Date</dt><dd>{{ ddmmyyyy }}</dd>
      <dt>Currency</dt><dd>USD</dd>
    </dl>
  </div>

  <a class="vback" href="javascript:window.close()">Close this tab</a>
</div></body></html>"""

LOGIN_HTML = """<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Sign in &mdash; """ + APP_NAME + """</title>""" + FONT_LINK + """
<style>{{ css|safe }}
/* Sign-in screen, matching the Vervi-Books login: centred column, product mark
   above the card, and the card itself carrying only the form. */
.auth{min-height:100vh;display:flex;flex-direction:column;align-items:center;
  justify-content:center;padding:40px 20px;gap:0;
  background:linear-gradient(180deg,#eef1fd 0%,#f6f7ff 55%,#eff2fe 100%)}
.authmark{width:78px;height:78px;border-radius:21px;color:#fff;display:grid;
  place-items:center;background:linear-gradient(145deg,#6a5cf0 0%,#7b4ef5 100%);
  box-shadow:0 14px 30px -12px rgba(106,92,240,.75)}
.authmark svg{width:38px;height:38px}
.authname{font-size:32px;font-weight:700;letter-spacing:-.028em;margin:20px 0 0;
  color:#141a2e;text-align:center}
.authtag{font-size:15px;color:#93a0b8;margin:7px 0 0;text-align:center}
.authbox{width:100%;max-width:428px;background:#fff;border-radius:17px;padding:34px 32px;
  margin-top:30px;box-shadow:0 18px 45px -22px rgba(38,44,90,.30),
  0 2px 6px rgba(38,44,90,.05)}
.authbox h2{font-size:23px;font-weight:700;margin:0;letter-spacing:-.02em;color:#141a2e}
.authbox .sub{font-size:14.5px;color:#93a0b8;margin:6px 0 24px}
.authbox label{display:block;font-size:13.5px;font-weight:500;color:#48546c;
  margin:0 0 7px}
.field{position:relative;margin-bottom:19px}
.field svg{position:absolute;left:15px;top:50%;transform:translateY(-50%);
  width:19px;height:19px;color:#98a5bb;pointer-events:none}
.field input{width:100%;height:50px;padding:0 15px 0 44px;border:1px solid transparent;
  border-radius:11px;background:#edf1f8;color:#141a2e;font-size:15px;
  font-family:var(--font);transition:border-color .15s,box-shadow .15s,background .15s}
.field input::placeholder{color:#a9b4c6}
.field input:focus{outline:0;background:#fff;border-color:#6a5cf0;
  box-shadow:0 0 0 3px rgba(106,92,240,.16)}
.authbox button{width:100%;height:52px;margin-top:4px;border-radius:11px;font-size:15.5px;
  font-weight:700;background:linear-gradient(90deg,#4c5fe0 0%,#8b3df0 100%);
  box-shadow:0 10px 22px -12px rgba(90,70,230,.9)}
.authfoot{font-size:13.5px;color:#9aa6bb;margin-top:26px;text-align:center}
.authbox .banner{margin:0 0 18px;font-size:13.5px}
@media (prefers-color-scheme:dark){
  .auth{background:linear-gradient(180deg,#0c1017 0%,#11161f 55%,#0c1017 100%)}
  .authname{color:#eaeff8}
  .authbox{background:#161b25;box-shadow:0 18px 45px -20px rgba(0,0,0,.8)}
  .authbox h2{color:#eaeff8}
  .authbox label{color:#aeb9cd}
  .field input{background:#1e2530;color:#eaeff8}
  .field input:focus{background:#232b38}
}
</style></head><body>
<div class="auth">
  <div class="authmark">""" + BOOK_ICON + """</div>
  <h1 class="authname">""" + APP_NAME + """</h1>
  <p class="authtag">Upwork Statement to Journal Automation</p>

  <div class="authbox">
    <h2>Welcome back</h2>
    <p class="sub">Sign in to your account to continue</p>

    {% if error %}<div class="banner err"><span>&#9888;</span><div>{{ error }}</div></div>{% endif %}

    <form method="post">
      <label for="u">Username</label>
      <div class="field">""" + MAIL_ICON + """
        <input type="text" id="u" name="username" autocomplete="username" autofocus
               placeholder="admin" value="{{ username }}" required>
      </div>
      <label for="p">Password</label>
      <div class="field">""" + LOCK_ICON + """
        <input type="password" id="p" name="password" autocomplete="current-password"
               placeholder="&bull;&bull;&bull;&bull;&bull;&bull;&bull;&bull;" required>
      </div>
      <button type="submit">Sign In</button>
    </form>
  </div>

  <div class="authfoot">""" + COMPANY + """ &copy; 2026 &middot; Secure &amp; Private</div>
</div></body></html>"""

RESULT_SIDE = """
  <div class="sblock c1">
    <div class="slabel">This run</div>
    <div class="srow"><span class="k">Statement</span><span class="v">{{ statement_name }}</span></div>
    <div class="srow"><span class="k">Rows read</span><span class="v">{{ rows_read }}</span></div>
    <div class="srow"><span class="k">Vouchers</span><span class="v">{{ vouchers }}</span></div>
    <div class="srow"><span class="k">Journal lines</span><span class="v">{{ total_lines }}</span></div>
    {% if duplicates %}<div class="srow"><span class="k">Already imported</span>
      <span class="v">{{ duplicates }} skipped</span></div>{% endif %}
  </div>
  <div class="sblock c3">
    <div class="slabel">Totals ({{ currency }})</div>
    <div class="srow"><span class="k">Debit</span><span class="v big">{{ total_debit }}</span></div>
    <div class="srow"><span class="k">Credit</span><span class="v big">{{ total_credit }}</span></div>
    <div class="srow"><span class="k">Difference</span>
      <span class="v big {{ 'ok' if balanced else 'bad' }}">{{ difference }}</span></div>
  </div>
  <div class="sblock c4">
    <div class="slabel">Download</div>
    <div class="row" style="margin-top:2px;gap:8px">
      <a href="{{ url_for('download', run_id=run_id, kind='csv') }}" style="width:100%">
        <button style="width:100%">Download import CSV</button></a>
    </div>
    <div class="row" style="margin-top:8px;gap:8px">
      <a href="{{ url_for('download', run_id=run_id, kind='split') }}" style="width:100%">
        <button class="sidealt" style="width:100%">JE + Sales as two files</button></a>
    </div>
    <div class="snote">Numbers are reserved, not locked &mdash; they are only
      taken once you download. Leave without downloading and they come round again.</div>
    <div class="row" style="margin-top:10px;gap:8px">
      <a href="{{ url_for('index') }}" style="width:100%">
        <button type="button" class="sidealt" style="width:100%">&larr; Convert another statement</button></a>
    </div>
  </div>
  <div class="sblock c2">
    <div class="slabel">Settings used</div>
    <div class="srow"><span class="k">IGST</span><span class="v">{{ igst_pct }}%</span></div>
    <div class="srow"><span class="k">Rates</span><span class="v">{{ fx_note }}</span></div>
    <div class="srow"><span class="k">Sales</span><span class="v">{{ income_mode }}</span></div>
    <div class="srow"><span class="k">Doc numbers</span><span class="v">{{ n_docs }} reserved</span></div>
  </div>
"""

RESULT_HTML = """<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>""" + APP_NAME + """ &mdash; {{ total_lines }} lines</title>""" + FONT_LINK + """
<style>{{ css|safe }}</style>
""" + THEME_BOOT + """</head><body>
<div class="shell">""" + SIDEBAR.replace(
    "{% block sidebody %}{% endblock %}", RESULT_SIDE) + """
<main class="main">
  <div class="top"><h1>Journal ready</h1>
    <span class="s">{{ vouchers }} vouchers &middot; {{ total_lines }} lines</span>
    """ + TOPRIGHT.replace('<div class="topright">',
                           '<div class="topright">' + BACK_BTN) + """</div>

  {% if balanced %}
  <div class="banner ok"><span>&#10004;</span><div><strong>Balanced.</strong>
  All {{ vouchers }} vouchers pass the debit = credit check.{% if rcm_net is not none %}
  RCM input and output IGST both total {{ rcm_total }} and net to {{ rcm_net }}.{% endif %}</div></div>
  {% else %}
  <div class="banner err"><span>&#9888;</span><div><strong>Not balanced.</strong>
  Debits and credits differ by {{ difference }} &mdash; see Exceptions.</div></div>
  {% endif %}

  {% if errors %}<div class="banner err"><span>&#9940;</span>
  <div><strong>{{ errors }} error{{ '' if errors == 1 else 's' }}.</strong>
  Those rows were not posted &mdash; see Exceptions.</div></div>{% endif %}

  {% if duplicates %}<div class="banner warn"><span>&#128260;</span>
  <div><strong>{{ duplicates }} row{{ '' if duplicates == 1 else 's' }} already imported.</strong>
  Matched on Ref ID against earlier runs and skipped, so nothing is posted twice.
  See the <strong>Already imported</strong> tab for the Ref IDs and when they
  went in.</div></div>{% endif %}

  {% if fx_warnings %}<div class="banner warn"><span>&#128197;</span>
  <div><strong>Exchange-rate fallback.</strong>
  <ul>{% for w in fx_warnings %}<li>{{ w }}</li>{% endfor %}</ul></div></div>{% endif %}

  <div class="panel tabwrap">
    <div class="tabs">
      <button type="button" class="tab on" onclick="showTab(this,'p-imp')">Import file<span class="n">{{ total_lines }}</span></button>
      <button type="button" class="tab" onclick="showTab(this,'p-jrn')">Journal</button>
      <button type="button" class="tab" onclick="showTab(this,'p-rec')">Reconciliation</button>
      <button type="button" class="tab" onclick="showTab(this,'p-fx')">FX audit<span class="n">{{ n_fx }}</span></button>
      <button type="button" class="tab" onclick="showTab(this,'p-exc')">Exceptions<span class="n">{{ n_exc }}</span></button>
      <button type="button" class="tab" onclick="showTab(this,'p-dup')">Already imported<span class="n">{{ n_dup }}</span></button>
      <button type="button" class="tab" onclick="showTab(this,'p-skp')">Skipped<span class="n">{{ n_skip }}</span></button>
    </div>
    <div id="p-imp" class="pane on">{{ import_table|safe }}</div>
    <div id="p-jrn" class="pane">{{ journal_table|safe }}</div>
    <div id="p-rec" class="pane">{{ recon_table|safe }}</div>
    <div id="p-fx"  class="pane">{{ fx_audit_table|safe }}</div>
    <div id="p-exc" class="pane">{{ exceptions_table|safe }}</div>
    <div id="p-dup" class="pane">{{ duplicates_table|safe }}</div>
    <div id="p-skp" class="pane">{{ skipped_table|safe }}</div>
  </div>
</main></div><script>""" + TAB_JS + """</script></body></html>"""


# --------------------------------------------------------------------------- #
# Rendering helpers
# --------------------------------------------------------------------------- #

def fmt_money(value: float) -> str:
    # Collapse floating-point negative zero, so a balanced run shows 0.00 not -0.00.
    if abs(value) < 0.005:
        value = 0.0
    return f"{value:,.2f}"


_table_seq = itertools.count(1)


def html_table(frame: pd.DataFrame, empty_message: str,
               numeric: tuple[str, ...] = (), voucher_breaks: bool = False,
               precise: tuple[str, ...] = (), filterable: bool = True,
               links: dict | None = None, actions=None) -> str:
    """Render a DataFrame as a styled table.

    Hand-rolled rather than `DataFrame.to_html` so amounts can be right-aligned,
    Dr/Cr, Type and Severity rendered as pills, a rule drawn between vouchers,
    and a filter box put under every column heading -- which together are what
    make a hundred-line journal usable rather than merely visible.
    """
    if frame is None or frame.empty:
        return f'<div class="scroll"><div class="empty">{empty_message}</div></div>'

    table_id = f"t{next(_table_seq)}"
    if filterable:
        # Filter control lives inside the heading, next to the name -- the column
        # and the way to narrow it are the same thing, so they belong together.
        head = "".join(
            f'<th><span class="thin"><span>{escape(str(c))}</span>'
            f'<button type="button" class="fbtn" onclick="openFilter(this)" '
            f'aria-label="Filter {escape(str(c), quote=True)}">{FILTER_ICON}</button>'
            f"</span></th>" for c in frame.columns)
    else:
        head = "".join(f"<th>{escape(str(c))}</th>" for c in frame.columns)
    if actions:
        head += '<th class="actcol">Edit</th>'
    body: list[str] = []
    previous_doc = None
    doc_col = "Document Number" if "Document Number" in frame.columns else None

    for _, row in frame.iterrows():
        classes = []
        if voucher_breaks and doc_col:
            if previous_doc is not None and row[doc_col] != previous_doc:
                classes.append("vstart")
            previous_doc = row[doc_col]
        # A totals row is a summary of the table, not a row of it -- filtering
        # must never hide it, or the figures stop adding up on screen.
        if str(row.get(frame.columns[0], "")) == "TOTAL":
            classes.append("total")
            classes.append("keep")

        cells = []
        for col in frame.columns:
            value = row[col]
            if col == "Dr/Cr" and str(value) in ("Dr", "Cr"):
                cells.append(f'<td><span class="pill {str(value).lower()}">{value}</span></td>')
            elif col in ("Type", "Voucher Type") and str(value) in (VOUCHER_JE, VOUCHER_SALES):
                cells.append(f'<td><span class="pill {value}">{value}</span></td>')
            elif col == "Severity" and str(value) in ("ERROR", "INFO"):
                cells.append(f'<td><span class="pill {value}">{value}</span></td>')
            elif col in precise:
                # FX rates are published to 4 decimals -- rounding them to money
                # precision would hide the very number being audited.
                cells.append(f'<td class="num">'
                             f'{"" if pd.isna(value) else f"{float(value):,.4f}"}</td>')
            elif col in numeric:
                shown = "" if value == "" or pd.isna(value) else fmt_money(float(value))
                # A zero on the unused side of a leg is noise; blank it out.
                if shown == "0.00":
                    shown = ""
                cells.append(f'<td class="num">{escape(shown)}</td>')
            elif links and col in links:
                # The cell text stays as-is; the link is what turns a claim
                # ("RBI") into something the reader can go and check.
                href = links[col](row) or ""
                label = "" if pd.isna(value) else escape(str(value))
                cells.append(
                    f'<td><a class="srclink" href="{escape(href, quote=True)}" '
                    f'target="_blank" rel="noopener">{label}</a></td>'
                    if href else f"<td>{label}</td>")
            elif col == "Narration":
                # Clipped to the column width; the title carries the full text so
                # nothing is lost -- hover any narration to read all of it.
                text = "" if pd.isna(value) else str(value)
                cells.append(f'<td class="narr" title="{escape(text, quote=True)}">'
                             f'{escape(text)}</td>')
            else:
                cells.append(f'<td>{"" if pd.isna(value) else escape(str(value))}</td>')
        if actions:
            cells.append(f'<td class="actcol">{actions(row)}</td>')
        body.append(f'<tr class="{" ".join(classes)}">{"".join(cells)}</tr>')

    rows_total = len(body)
    toolbar = ""
    if filterable:
        toolbar = (f'<div class="tbar"><span class="fcount" id="{table_id}-count">'
                   f'{rows_total} rows</span>'
                   f'<button type="button" class="fclear" '
                   f'onclick="clearFilters(\'{table_id}\')">Clear filters</button></div>')

    return (f'<div class="tablewrap" data-total="{rows_total}">{toolbar}'
            f'<div class="scroll"><table id="{table_id}">'
            f'<thead><tr>{head}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody></table></div></div>')


def fx_source_link(row) -> str:
    """Link an FX Audit row's Source cell to a live check of that rate.

    Only rates that claim to come from RBI are checkable; a manual rate or one
    supplied in a file has no published source to compare against, so those
    cells stay plain text.
    """
    source = str(row.get("Source", ""))
    rate_date = row.get("Rate Date Used")
    if "RBI" not in source and "CACHE" not in source:
        return ""
    if rate_date is None or pd.isna(rate_date):
        return ""
    rate = row.get("Rate (INR/USD)")
    return url_for("fx_verify", rate_date=rate_date.isoformat(),
                   used=("" if rate is None or pd.isna(rate) else f"{rate:.4f}"),
                   txn=row.get("Transaction Date").isoformat()
                   if row.get("Transaction Date") is not None else "")


def save_upload(file_storage, folder: Path) -> Path | None:
    """Persist an uploaded file, rejecting anything that isn't a spreadsheet."""
    if not file_storage or not file_storage.filename:
        return None
    name = secure_filename(file_storage.filename)
    if Path(name).suffix.lower() not in ALLOWED_SUFFIXES:
        raise ValueError(f"'{file_storage.filename}' is not a .csv or .xlsx file")
    target = folder / name
    file_storage.save(target)
    return target


def doc_range(journal: pd.DataFrame) -> str:
    """A one-line summary of the numbers a run used, for the history."""
    if journal.empty:
        return ""
    parts = []
    for kind in (VOUCHER_SALES, VOUCHER_JE):
        docs = sorted(journal[journal["Voucher Type"] == kind]["Document Number"].unique())
        if docs:
            parts.append(docs[0] if len(docs) == 1 else f"{docs[0]} - {docs[-1]}")
    return " | ".join(parts)


def archive_run(run: dict) -> None:
    """Keep a downloaded run's outputs, and note it in the history.

    Only downloaded runs are kept -- the same rule the document numbers follow,
    so the history is a record of what was actually taken, not what was looked at.
    """
    stamp = datetime.now()
    stem = secure_filename(run.get("stem") or "statement").rstrip("_") or "statement"
    folder = EXPORTS_DIR / f"{stamp:%Y%m%d-%H%M%S} {stem}"
    folder.mkdir(parents=True, exist_ok=True)

    for kind in ("csv", "xlsx"):
        src = run.get(kind)
        if src and src.exists():
            shutil.copy2(src, folder / f"{stem} - import.{kind}")
    frame = run.get("import_frame")
    if frame is not None and not frame.empty:
        for kind in (VOUCHER_JE, VOUCHER_SALES):
            part = frame[frame["Type"] == kind]
            part.to_csv(folder / f"{stem} - {kind}.csv", index=False,
                        encoding="utf-8-sig")

    row = {
        "converted_at": stamp.strftime("%Y-%m-%d %H:%M:%S"),
        "statement": run.get("stem", ""),
        "rows_read": run.get("rows_read", ""),
        "vouchers": run.get("vouchers", ""),
        "journal_lines": run.get("total_lines", ""),
        "duplicates": run.get("duplicates", 0),
        "doc_numbers": run.get("doc_range", ""),
        "folder": folder.name,
    }
    exists = HISTORY_FILE.exists()
    with HISTORY_FILE.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=HISTORY_COLUMNS)
        if not exists:
            writer.writeheader()
        writer.writerow(row)


def read_history() -> list[dict]:
    """Past conversions, newest first."""
    if not HISTORY_FILE.exists():
        return []
    try:
        with HISTORY_FILE.open(newline="", encoding="utf-8") as handle:
            rows = list(csv.DictReader(handle))
    except OSError:
        return []
    rows.reverse()
    for r in rows:
        folder = EXPORTS_DIR / (r.get("folder") or "")
        r["files"] = sorted(f.name for f in folder.glob("*")) if folder.is_dir() else []
    return rows


def registry_state(doc_series: dict | None = None) -> dict:
    """Where numbering stands: the last number used, and the next one out.

    A count of rows was the wrong thing to show -- what matters before a run is
    "the last invoice I actually issued", so numbering visibly continues from
    there. Read through DocumentNumberer so the seed file counts too, and via
    peek() so looking never consumes a number.
    """
    try:
        n = DocumentNumberer(doc_series, registry_path=DEFAULT_DOC_REGISTRY,
                             source="preview")
    except Exception:
        return {"ok": False}

    today = datetime.now().date()
    sales_next, sales_last = n.peek(VOUCHER_SALES, today)
    je_next, je_last = n.peek(VOUCHER_JE, today)
    return {"ok": True, "sales_next": sales_next, "sales_last": sales_last,
            "je_next": je_next, "je_last": je_last,
            "je_month": today.strftime("%b")}


# --------------------------------------------------------------------------- #
# Routes
# --------------------------------------------------------------------------- #

@app.after_request
def no_store(response):
    """Stop the browser serving a cached result page for a fresh upload."""
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    return response


def client_ip() -> str:
    """Caller's address, honouring the tunnel's forwarding header."""
    forwarded = request.headers.get("CF-Connecting-IP") or \
        request.headers.get("X-Forwarded-For", "")
    return (forwarded.split(",")[0].strip() or request.remote_addr or "?")


def locked_out(ip: str) -> int:
    """Seconds this address must still wait, or 0 if it may try now."""
    now = time.monotonic()
    recent = [t for t in _attempts.get(ip, []) if now - t < LOCKOUT_SECONDS]
    _attempts[ip] = recent
    if len(recent) >= MAX_ATTEMPTS:
        return int(LOCKOUT_SECONDS - (now - recent[0])) + 1
    return 0


@app.before_request
def require_login():
    """Every page needs a session, except the login screen itself."""
    if request.endpoint in PUBLIC_ENDPOINTS or session.get("auth"):
        return None
    return redirect(url_for("login", next=request.path))


@app.route("/healthz")
def healthz():
    """Unauthenticated liveness check, so a tunnel can be verified."""
    return "ok", 200


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("auth"):
        return redirect(url_for("index"))

    error = None
    username = ""
    if request.method == "POST":
        ip = client_ip()
        wait = locked_out(ip)
        if wait:
            error = f"Too many attempts. Try again in {wait} seconds."
        else:
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            # compare_digest on both halves, so neither is leaked by timing.
            ok = (secrets.compare_digest(username, APP_USER)
                  and secrets.compare_digest(password, APP_PASS))
            if ok:
                session.clear()
                session["auth"] = True
                session["user"] = username
                session.permanent = False
                _attempts.pop(ip, None)
                nxt = request.args.get("next", "")
                # Only ever redirect within this app.
                return redirect(nxt if nxt.startswith("/") and not nxt.startswith("//")
                                else url_for("index"))
            _attempts.setdefault(ip, []).append(time.monotonic())
            error = "That username and password don't match. Try again."

    return render_template_string(LOGIN_HTML, css=BASE_CSS, error=error,
                                  username=username)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


def master_preview() -> dict:
    """Render the master database for the sidebar and tabs.

    Read fresh on every request, so editing mapping_master.csv shows up on a
    refresh without restarting the server.
    """
    try:
        mapping = load_mapping(MASTER_MAPPING, on_duplicate="resolve")
    except Exception as exc:
        return {"master_error": f"{type(exc).__name__}: {exc}", "wallet_table": "",
                "treatment_table": "", "n_accounts": 0, "n_treatments": 0}

    wallets = pd.DataFrame(
        [{"Sr. No.": i,
          "Account Name (col G)": mapping.wallet_display.get(key, key),
          "GL Name (col H)": gl,
          "Cost Center (col I)": mapping.cost_centers.get(key, ""),
          "Subsidiary (col J)": mapping.subsidiaries.get(key, "")}
         for i, (key, gl) in enumerate(sorted(mapping.wallets.items()), start=1)])
    treatments = pd.DataFrame(
        [{"Sr. No.": i, "Nature": t.nature, "Treatment": t.kind,
          "Debit": " + ".join(x for x in (t.debit_1, t.debit_2) if x and x.upper() != "NA"),
          "Credit": " + ".join(x for x in (t.credit_1, t.credit_2) if x and x.upper() != "NA")}
         for i, t in enumerate(mapping.treatments.values(), start=1)])
    def account_action(row):
        return (f'<button type="button" class="editbtn" onclick="openEdit(this)" '
                f'data-kind="account" '
                f'data-key="{escape(str(row["Account Name (col G)"]), quote=True)}" '
                f'data-account="{escape(str(row["Account Name (col G)"]), quote=True)}" '
                f'data-gl="{escape(str(row["GL Name (col H)"]), quote=True)}" '
                f'data-cc="{escape(str(row["Cost Center (col I)"]), quote=True)}" '
                f'data-sub="{escape(str(row["Subsidiary (col J)"]), quote=True)}">Edit</button>')

    def treatment_action(row):
        t = by_nature[row["Nature"]]
        return (f'<button type="button" class="editbtn" onclick="openEdit(this)" '
                f'data-kind="treatment" '
                f'data-key="{escape(t.nature, quote=True)}" '
                f'data-nature="{escape(t.nature, quote=True)}" '
                f'data-d1="{escape(t.debit_1, quote=True)}" '
                f'data-d2="{escape(t.debit_2, quote=True)}" '
                f'data-c1="{escape(t.credit_1, quote=True)}" '
                f'data-c2="{escape(t.credit_2, quote=True)}">Edit</button>')

    by_nature = {t.nature: t for t in mapping.treatments.values()}
    return {
        "master_error": None,
        "wallet_table": html_table(wallets, "No accounts in the master.",
                                   actions=account_action),
        "treatment_table": html_table(treatments, "No treatments in the master.",
                                      actions=treatment_action),
        "n_accounts": len(wallets),
        "n_treatments": len(treatments),
    }


def render_form(message: str | None = None, form=None, notice: str | None = None) -> str:
    """The upload form, optionally with an error, preserving what was typed."""
    form = form or {}
    doc_series = {
        VOUCHER_JE: form.get("doc_je", "").strip() or DEFAULT_DOC_SERIES[VOUCHER_JE],
        VOUCHER_SALES: form.get("doc_sales", "").strip() or DEFAULT_DOC_SERIES[VOUCHER_SALES],
    }
    reg = registry_state(doc_series)
    return render_template_string(
        FORM_HTML, css=BASE_CSS, error=message, notice=notice,
        currency=form.get("currency", "INR"), fx_rate=form.get("fx_rate", ""),
        igst_rate=form.get("igst_rate", "18"),
        fx_source=form.get("fx_source", "rbi"),
        fx_offline=bool(form.get("fx_offline")),
        reset_docs=bool(form.get("reset_docs")),
        ignore_posted=bool(form.get("ignore_posted")),
        income_mode=form.get("income_mode", "two-entry"),
        entity=form.get("entity", COMPANY),
        ie_flag=form.get("ie_flag", "I"),
        doc_je=form.get("doc_je", DEFAULT_DOC_SERIES[VOUCHER_JE]),
        doc_sales=form.get("doc_sales", DEFAULT_DOC_SERIES[VOUCHER_SALES]),
        samples_exist=SAMPLE_STATEMENT.exists(),
        master_path=MASTER_MAPPING.name,
        reg=reg, n_history=len(read_history()),
        **master_preview())


@app.route("/")
def index():
    return render_form(request.args.get("error"), notice=request.args.get("added"))


def form_with_error(message: str, form) -> str:
    return render_form(message, form)


@app.route("/master/account", methods=["POST"])
def add_account():
    """Append an Account Name -> GL Name -> Cost Center row to the master."""
    try:
        account = request.form.get("account", "")
        gl_name = request.form.get("gl_name", "")
        cost_center = request.form.get("cost_center", "")
        add_wallet(MASTER_MAPPING, account, gl_name, cost_center,
                   request.form.get("subsidiary", ""))
    except Exception as exc:
        return redirect(url_for("index", error=str(exc)))
    note = f"Added {clean_text(account)} -> {clean_text(gl_name)}"
    if clean_text(cost_center):
        note += f", cost centre {clean_text(cost_center)}"
    return redirect(url_for("index", added=note))


@app.route("/master/account/update", methods=["POST"])
def edit_account():
    """Rewrite one Table B row."""
    key = request.form.get("key", "")
    try:
        update_wallet(MASTER_MAPPING, key, request.form.get("account", ""),
                      request.form.get("gl_name", ""),
                      request.form.get("cost_center", ""),
                      request.form.get("subsidiary", ""))
    except Exception as exc:
        return redirect(url_for("index", error=str(exc)))
    return redirect(url_for("index", added=f"Updated {clean_text(key)}"))


@app.route("/master/account/delete", methods=["POST"])
def remove_account():
    """Remove one Table B row."""
    key = request.form.get("key", "")
    try:
        delete_wallet(MASTER_MAPPING, key)
    except Exception as exc:
        return redirect(url_for("index", error=str(exc)))
    return redirect(url_for("index", added=f"Removed {clean_text(key)}"))


@app.route("/master/treatment/update", methods=["POST"])
def edit_rule():
    """Rewrite one Table A row."""
    key = request.form.get("key", "")
    try:
        update_treatment(MASTER_MAPPING, key, request.form.get("nature", ""),
                         request.form.get("debit_1", ""), request.form.get("debit_2", ""),
                         request.form.get("credit_1", ""), request.form.get("credit_2", ""))
    except Exception as exc:
        return redirect(url_for("index", error=str(exc)))
    return redirect(url_for("index", added=f"Updated the rule for {clean_text(key)}"))


@app.route("/master/treatment/delete", methods=["POST"])
def remove_rule():
    """Remove one Table A row."""
    key = request.form.get("key", "")
    try:
        delete_treatment(MASTER_MAPPING, key)
    except Exception as exc:
        return redirect(url_for("index", error=str(exc)))
    return redirect(url_for("index", added=f"Removed the rule for {clean_text(key)}"))


@app.route("/master/treatment", methods=["POST"])
def add_rule():
    """Append a transaction-type rule to the master database."""
    try:
        nature = request.form.get("nature", "")
        add_treatment(MASTER_MAPPING, nature,
                      request.form.get("debit_1", ""), request.form.get("debit_2", ""),
                      request.form.get("credit_1", ""), request.form.get("credit_2", ""))
    except Exception as exc:
        return redirect(url_for("index", error=str(exc)))
    return redirect(url_for("index", added=f"Added the rule for {clean_text(nature)}"))


@app.route("/run", methods=["POST"])
def run():
    form = request.form
    workdir = Path(mkdtemp(prefix="upwork_journal_"))

    # --- resolve inputs ------------------------------------------------------
    # The mapping is never uploaded -- it always comes from the master database.
    mapping_path = MASTER_MAPPING
    if not mapping_path.exists():
        return form_with_error(f"Master database not found at {mapping_path}.", form)

    try:
        statement_path = save_upload(request.files.get("statement"), workdir)

        # An uploaded file ALWAYS wins over the sample button, so choosing a file
        # can never be silently ignored in favour of the sample.
        if statement_path is None:
            if form.get("use_sample") and SAMPLE_STATEMENT.exists():
                statement_path = SAMPLE_STATEMENT
            else:
                return form_with_error("Choose an Upwork statement to convert.", form)
    except ValueError as exc:
        return form_with_error(str(exc), form)

    # --- validate options ----------------------------------------------------
    currency = form.get("currency", "INR")
    income_mode = form.get("income_mode", "two-entry")
    fx_source = form.get("fx_source", "rbi")
    entity = form.get("entity", "").strip() or COMPANY
    ie_flag = form.get("ie_flag", "I").strip()
    doc_series = {
        VOUCHER_JE: form.get("doc_je", "").strip() or DEFAULT_DOC_SERIES[VOUCHER_JE],
        VOUCHER_SALES: form.get("doc_sales", "").strip() or DEFAULT_DOC_SERIES[VOUCHER_SALES],
    }

    igst_pct = to_decimal(form.get("igst_rate", "18"))
    if igst_pct is None or igst_pct < 0:
        return form_with_error(f"'{form.get('igst_rate')}' is not a valid IGST rate.", form)
    igst_rate = igst_pct / Decimal("100")

    # --- build ---------------------------------------------------------------
    try:
        mapping = load_mapping(mapping_path, on_duplicate="resolve")
        statement = load_statement(statement_path)

        manual_rate = None
        typed_rate = form.get("fx_rate", "").strip()
        if typed_rate:
            manual_rate = to_decimal(typed_rate)
            if manual_rate is None or manual_rate <= 0:
                return form_with_error(f"'{typed_rate}' is not a valid exchange rate.", form)

        if currency == "USD":
            fx_provider = FxProvider(rate=Decimal("1"), source="-")
        elif fx_source == "manual":
            if manual_rate is None:
                return form_with_error(
                    "A manual flat rate needs a value in the Manual rate box.", form)
            fx_provider = FxProvider(rate=manual_rate, source="MANUAL")
        else:
            fx_provider = RbiFxProvider(
                cache_path=HERE / "fx_cache.csv", manual_rate=manual_rate,
                offline=bool(form.get("fx_offline")), log=lambda m: None)

        # Shares the registry with the CLI, so numbers issued in either place are
        # never handed out again by the other.
        numberer = DocumentNumberer(
            doc_series, registry_path=DEFAULT_DOC_REGISTRY,
            source=statement_path.name, reset=bool(form.get("reset_docs")))

        reporter = Reporter()
        for key, (chosen, rejected) in mapping.resolved_wallets.items():
            reporter.exception("-", "DUPLICATE_RESOLVED",
                               f"Table B lists '{mapping.wallet_display.get(key, key)}' more "
                               f"than once; used '{chosen}' as the name matches, ignored "
                               f"{', '.join(repr(gl) for gl in rejected)}", severity="INFO")
        for key, candidates in mapping.ambiguous_wallets.items():
            reporter.exception("-", "AMBIGUOUS_ACCOUNT",
                               f"Table B maps '{mapping.wallet_display.get(key, key)}' to "
                               f"conflicting GLs: {', '.join(candidates)}")

        # Ref IDs already journalised, so an overlapping or re-uploaded
        # statement is not posted twice.
        posted = PostedLedger(DEFAULT_POSTED_LEDGER,
                              source=statement_path.name,
                              enabled=not form.get("ignore_posted"))

        journal = build_journal(statement, mapping, igst_rate=igst_rate,
                                currency=currency, fx_provider=fx_provider,
                                income_mode=income_mode, reporter=reporter,
                                numberer=numberer, entity=entity, posted=posted)

        for warning in fx_provider.warnings:
            reporter.exception("-", "FX_FALLBACK", warning, severity="INFO")

        import_frame = build_import_frame(journal, entity=entity, currency=currency,
                                          ie_flag=ie_flag)
        reconciliation = build_reconciliation(journal, currency)
        fx_audit = pd.DataFrame(fx_provider.audit_rows()
                                if hasattr(fx_provider, "audit_rows") else [])
        exceptions = pd.DataFrame(reporter.exceptions)
        all_skipped = pd.DataFrame(reporter.skipped)
        # Duplicates get their own view: they are why a re-uploaded statement
        # produced fewer vouchers, and burying them among zero-amount and
        # no-entry rows makes that hard to see.
        if not all_skipped.empty and "Reason" in all_skipped.columns:
            is_dup = all_skipped["Reason"].astype(str).str.startswith("Already imported")
            duplicate_rows = all_skipped[is_dup].copy()
            skipped = all_skipped[~is_dup].copy()
        else:
            duplicate_rows = pd.DataFrame()
            skipped = all_skipped

        out_xlsx = workdir / "journal.xlsx"
        # The workbook keeps a single Skipped sheet with everything in it.
        out_csv = write_outputs(out_xlsx, journal, reconciliation, exceptions,
                                all_skipped, fx_audit=fx_audit,
                                import_frame=import_frame)
        # Numbers are NOT committed here. Looking at a journal on screen is not
        # posting it, so the registry is only written when the file is actually
        # downloaded -- see `download`. Until then these numbers stay available.
        n_docs = len(numberer.issued)
    except Exception as exc:
        return form_with_error(f"{type(exc).__name__}: {exc}", form)

    run_id = uuid.uuid4().hex
    # Name the downloads after the statement they came from, so a folder of
    # exports says which is which instead of a pile of journal.csv files.
    RUNS[run_id] = {"xlsx": out_xlsx, "csv": out_csv,
                    "numberer": numberer, "posted": posted, "committed": False,
                    "stem": Path(statement_path).stem,
                    "import_frame": import_frame,
                    "rows_read": len(statement),
                    "vouchers": journal["Document Number"].nunique() if not journal.empty else 0,
                    "total_lines": len(journal),
                    "duplicates": len(duplicate_rows),
                    "doc_range": doc_range(journal)}

    # --- present -------------------------------------------------------------
    total_debit = float(journal["Debit"].sum()) if not journal.empty else 0.0
    total_credit = float(journal["Credit"].sum()) if not journal.empty else 0.0
    difference = round(total_debit - total_credit, 2)

    rcm_in = float(journal[journal["Ledger"] == "RCM Input IGST"]["Debit"].sum()) \
        if not journal.empty else 0.0
    rcm_out = float(journal[journal["Ledger"] == "RCM Output IGST"]["Credit"].sum()) \
        if not journal.empty else 0.0

    money_cols = tuple(c for c in reconciliation.columns if "Total" in c) + ("Difference",)

    return render_template_string(
        RESULT_HTML, css=BASE_CSS, run_id=run_id,
        statement_name=statement_path.name,
        currency=currency, igst_pct=f"{float(igst_pct):g}",
        income_mode="2 entries" if income_mode == "two-entry" else "combined",
        fx_note=("manual flat" if fx_source == "manual"
                 else ("n/a" if currency == "USD" else "RBI, per date")),
        fx_warnings=fx_provider.warnings,
        rows_read=len(statement),
        vouchers=journal["Document Number"].nunique() if not journal.empty else 0,
        total_lines=len(journal), n_docs=n_docs,
        total_debit=fmt_money(total_debit), total_credit=fmt_money(total_credit),
        difference=fmt_money(difference), balanced=abs(difference) < 0.005,
        rcm_total=fmt_money(rcm_in), rcm_net=fmt_money(rcm_in - rcm_out) if rcm_in else None,
        errors=len(reporter.errors),
        duplicates=len(duplicate_rows),
        n_fx=len(fx_audit), n_exc=len(exceptions), n_skip=len(skipped),
        import_table=html_table(import_frame, "Nothing to import.",
                                numeric=("Amount", "Debit", "Credit",
                                         "Amount in base currency", "Amount in INR"),
                                voucher_breaks=True),
        journal_table=html_table(journal, "No journal lines were produced.",
                                 numeric=("Debit", "Credit", "Amount USD"),
                                 precise=("FX Rate",), voucher_breaks=True),
        recon_table=html_table(reconciliation, "Nothing posted.", numeric=money_cols),
        fx_audit_table=html_table(fx_audit, "No exchange-rate lookups (USD mode).",
                                  precise=("Rate (INR/USD)",),
                                  links={"Source": fx_source_link}),
        exceptions_table=html_table(exceptions, "No exceptions â€” clean run."),
        skipped_table=html_table(skipped, "Nothing skipped.", numeric=("Amount USD",)),
        duplicates_table=html_table(
            duplicate_rows,
            "No duplicates — every transaction in this statement is new.",
            numeric=("Amount USD",)),
        n_dup=len(duplicate_rows))


@app.route("/fx/<rate_date>")
def fx_verify(rate_date: str):
    """Check one journal rate against what RBI publishes for that date.

    Reached by clicking the Source cell in the FX Audit table. The archive only
    answers a POST, so there is no URL that shows a single day -- this route does
    that POST live and reports whether the figure in the books matches.
    """
    try:
        day = datetime.strptime(rate_date, "%Y-%m-%d").date()
    except ValueError:
        abort(404)

    used = (request.args.get("used") or "").strip()
    txn_date = (request.args.get("txn") or "").strip()
    days_back = 0
    if txn_date:
        try:
            days_back = (datetime.strptime(txn_date, "%Y-%m-%d").date() - day).days
        except ValueError:
            txn_date = ""

    cache = RateCache(HERE / "fx_cache.csv")
    cached = cache.rates.get(day)

    client = RbiArchiveClient(timeout=25, delay=0.0, retries=2, log=lambda m: None)
    fetched, reached = client.fetch_range(day, day)
    published = fetched.get(day)

    if not reached:
        verdict, note = "unreachable", "The archive did not respond."
    elif published is None:
        verdict, note = "nopublish", ""
    elif used and Decimal(used) != published:
        verdict, note = "differ", ""
    else:
        verdict, note = "match", ""

    return render_template_string(
        VERIFY_HTML, css=BASE_CSS, verdict=verdict, note=note,
        rate_date=day.isoformat(), rate_date_long=day.strftime("%d %b %Y"),
        ddmmyyyy=day.strftime("%d/%m/%Y"),
        txn_date=txn_date, days_back=days_back,
        used=used or None, published=published, cached=cached,
        archive_url=ARCHIVE_URL)


def split_zip(import_frame: pd.DataFrame, stem: str) -> io.BytesIO:
    """One archive holding the JE rows and the Sales rows as separate CSVs.

    Some systems want each voucher type imported on its own, and splitting by
    hand after the fact invites mistakes -- so the tool does it, from the same
    frame the combined file comes from.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for kind in (VOUCHER_JE, VOUCHER_SALES):
            part = import_frame[import_frame["Type"] == kind]
            # An empty part still gets a file, headers and all, so a missing
            # voucher type is visible rather than silently absent.
            zf.writestr(f"{stem} - {kind}.csv",
                        part.to_csv(index=False).encode("utf-8-sig"))
    buf.seek(0)
    return buf


HISTORY_HTML = """<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>""" + APP_NAME + """ &mdash; converted statements</title>""" + FONT_LINK + """
<style>{{ css|safe }}</style>
""" + THEME_BOOT + """</head><body>
<div class="shell">
<aside class="side">
  <div class="logo">
    <div class="mark">""" + BOOK_ICON + """</div>
    <div><div class="nm">""" + APP_NAME + """</div>
         <div class="co">""" + COMPANY + """</div></div>
  </div>
  <div class="sblock c1">
    <div class="slabel">Converted so far</div>
    <div class="srow"><span class="k">Statements</span><span class="v big">{{ runs|length }}</span></div>
    <div class="srow"><span class="k">Ref IDs on record</span><span class="v big">{{ n_refs }}</span></div>
    <div class="snote">Only downloaded runs are kept &mdash; the same rule the
      document numbers follow.</div>
  </div>
  <div class="sblock c4">
    <div class="slabel">Go</div>
    <div class="row" style="margin-top:2px">
      <a href="{{ url_for('index') }}" style="width:100%">
        <button type="button" style="width:100%">&larr; New journal</button></a>
    </div>
  </div>
</aside>
<main class="main">
  <div class="top"><h1>Converted statements</h1>
    <span class="s">what has been imported, and the files it produced</span>
    """ + TOPRIGHT.replace('<div class="topright">',
                           '<div class="topright">' + BACK_BTN) + """</div>

  {% if not runs %}
  <div class="panel"><p class="psub" style="margin:0">Nothing converted yet.
    A statement appears here once you download its output.</p></div>
  {% else %}
  {% for r in runs %}
  <div class="panel">
    <div class="step"><span class="n">{{ loop.index }}</span>
      <h2>{{ r.statement }}</h2></div>
    <p class="psub">{{ r.converted_at }} &middot; {{ r.rows_read }} rows &rarr;
      {{ r.vouchers }} vouchers, {{ r.journal_lines }} lines
      {%- if r.duplicates and r.duplicates != '0' %} &middot;
      {{ r.duplicates }} already imported{% endif %}
      {%- if r.doc_numbers %}<br>{{ r.doc_numbers }}{% endif %}</p>
    <div class="row" style="margin-top:0">
      {% for f in r.files %}
      <a href="{{ url_for('history_file', folder=r.folder, name=f) }}">
        <button class="ghost">{{ f }}</button></a>
      {% endfor %}
    </div>
  </div>
  {% endfor %}
  {% endif %}
</main></div>
<script>""" + TAB_JS + """</script></body></html>"""


@app.route("/history")
def history():
    """Statements already converted, with their output files."""
    refs = 0
    if DEFAULT_POSTED_LEDGER.exists():
        try:
            refs = max(0, sum(1 for _ in DEFAULT_POSTED_LEDGER.open(encoding="utf-8")) - 1)
        except OSError:
            refs = 0
    return render_template_string(HISTORY_HTML, css=BASE_CSS,
                                  runs=read_history(), n_refs=refs)


@app.route("/history/<folder>/<name>")
def history_file(folder: str, name: str):
    """Serve one archived output file."""
    base = EXPORTS_DIR.resolve()
    target = (base / folder / name).resolve()
    # Never serve anything outside the exports folder, whatever the URL says.
    if not str(target).startswith(str(base)) or not target.is_file():
        abort(404)
    return send_file(target, as_attachment=True, download_name=target.name)


@app.route("/download/<run_id>/<kind>")
def download(run_id: str, kind: str):
    if kind not in ("xlsx", "csv", "split"):
        abort(404)

    run = RUNS.get(run_id)
    # A run lives in memory until it is downloaded, so a restart loses it.
    # Say so plainly -- and say that nothing was consumed -- rather than
    # showing a bare 404 that looks like the file failed to build.
    if not run or (kind in ("xlsx", "csv") and not run[kind].exists()):
        return render_form(
            "That download expired. The app restarted before you downloaded it, "
            "so the file is gone. Nothing was posted and no document numbers were "
            "used — generate the journal again and it will produce the same "
            "numbers.")

    # Downloading is the moment the numbers leave the tool and become real, so
    # this is where they are committed to the registry. Generating and then
    # discarding a journal costs nothing -- the same numbers come round again.
    if not run["committed"]:
        try:
            run["numberer"].save()
            run["posted"].save()
            run["committed"] = True
            archive_run(run)
        except Exception:
            # A registry that cannot be written must not block the download; the
            # figures are still correct, only the reservation is missing.
            pass

    stem = secure_filename(run.get("stem") or "journal").rstrip("_") or "journal"

    if kind == "split":
        return send_file(split_zip(run["import_frame"], stem), as_attachment=True,
                         mimetype="application/zip",
                         download_name=f"{stem} - import (JE + Sales).zip")

    return send_file(run[kind], as_attachment=True,
                     download_name=f"{stem} - import.{kind}")


def main() -> None:
    global MASTER_MAPPING

    parser = argparse.ArgumentParser(description=f"{APP_NAME} -- local web UI.")
    parser.add_argument("--port", type=int, default=5000)
    parser.add_argument("--host", default="127.0.0.1",
                        help="Bind address; keep the localhost default unless you "
                             "genuinely intend to expose this on your network")
    parser.add_argument("--master", type=Path, default=None,
                        help=f"Accounting master database (default: {MASTER_MAPPING.name})")
    args = parser.parse_args()

    if args.master:
        MASTER_MAPPING = args.master.resolve()

    print(f"\n  {APP_NAME} running at  http://{args.host}:{args.port}")
    print(f"  Sign in as:        {APP_USER} / {'*' * len(APP_PASS)}")
    print(f"  Master database:   {MASTER_MAPPING}")
    print(f"  Document registry: {DEFAULT_DOC_REGISTRY}")
    if APP_PASS == "admin":
        print("\n  ! Default password in use. Before exposing this beyond your own\n"
              "    machine, set VERVI_USER and VERVI_PASS to something private.")
    print()
    app.run(host=args.host, port=args.port, debug=False)


if __name__ == "__main__":
    main()
