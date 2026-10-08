"""Storage for shared BDIT scores.

Normally backed by a Google Sheet (shared across the whole team). Needs two
credentials, read from environment variables (works both locally via a
.env file, loaded with python-dotenv, and on Hugging Face Spaces, where
Space secrets are injected as plain env vars regardless of SDK):

  - SHEET_URL                 the Google Sheet URL, shared with the service
                               account's email as an Editor
  - GCP_SERVICE_ACCOUNT_JSON  the ENTIRE downloaded service-account JSON key
                               file, as one string

If neither is configured (e.g. local preview before Google Cloud setup is
done), falls back to a local JSON file — same persistence model as the
original desktop app — so the app is still fully click-through-able.
"""

import json
import os
import time

from dotenv import load_dotenv

from scoring import KPIS, MONTHS

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SHEET_TAB = "BDIT Scores"
_CACHE_TTL_SECONDS = 20

HEADER = ["User Name", "Year", "Month"]
for _kpi in KPIS:
    HEADER += [f"{_kpi['name']} Value", f"{_kpi['name']} Slab", f"{_kpi['name']} Score"]
HEADER.append("Total Score")

MONTH_ORDER = {m: i for i, m in enumerate(MONTHS)}

_LOCAL_DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 ".local_data", "skoda_scores.json")


def _get_sheet_url():
    return os.environ.get("SHEET_URL") or os.environ.get("sheet_url")


def _get_service_account_info():
    raw = os.environ.get("GCP_SERVICE_ACCOUNT_JSON") or os.environ.get("gcp_service_account_json")
    if not raw:
        return None
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return None


def _using_google_sheets():
    return _get_sheet_url() is not None and _get_service_account_info() is not None


def load_db():
    """Return {user: {year: {month: {kpi_name: {value, pct, label, score}}}}}."""
    if _using_google_sheets():
        return _load_db_sheets()
    return _load_db_local()


def save_entry(username, year, month, scores):
    """Upsert one (user, year, month) entry."""
    if _using_google_sheets():
        _save_entry_sheets(username, year, month, scores)
    else:
        _save_entry_local(username, year, month, scores)


# ── Local JSON fallback (preview / no Google Cloud setup yet) ────────────────

def _load_db_local():
    if os.path.exists(_LOCAL_DATA_FILE):
        try:
            with open(_LOCAL_DATA_FILE) as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_entry_local(username, year, month, scores):
    db = _load_db_local()
    db.setdefault(username, {}).setdefault(str(year), {})[month] = scores
    os.makedirs(os.path.dirname(_LOCAL_DATA_FILE), exist_ok=True)
    with open(_LOCAL_DATA_FILE, "w") as f:
        json.dump(db, f, indent=2)


# ── Google Sheets backend (shared, production) ────────────────────────────────
# Gradio has no built-in resource/data caching like Streamlit, so it's done
# by hand here: the worksheet handle is resolved once per process, and
# load_db() results are cached for a short TTL. Without this, every page
# interaction would re-spend several Sheets API reads and can blow through
# Google's free read-quota (60 reads/minute) almost immediately.

_client_cache = None
_worksheet_cache = None
_load_cache = {"data": None, "ts": 0.0}


def _client():
    global _client_cache
    if _client_cache is None:
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_info(
            _get_service_account_info(), scopes=SCOPES
        )
        _client_cache = gspread.authorize(creds)
    return _client_cache


def _worksheet():
    global _worksheet_cache
    if _worksheet_cache is not None:
        return _worksheet_cache

    import gspread

    sh = _client().open_by_url(_get_sheet_url())
    try:
        ws = sh.worksheet(SHEET_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_TAB, rows=1000, cols=len(HEADER))
        ws.append_row(HEADER)
        _worksheet_cache = ws
        return ws
    if ws.row_values(1) != HEADER:
        ws.update("A1", [HEADER])
    _worksheet_cache = ws
    return ws


def _invalidate_cache():
    _load_cache["data"] = None
    _load_cache["ts"] = 0.0


def _load_db_sheets():
    now = time.time()
    if _load_cache["data"] is not None and (now - _load_cache["ts"]) < _CACHE_TTL_SECONDS:
        return _load_cache["data"]

    rows = _worksheet().get_all_records()
    db = {}
    for row in rows:
        user = str(row.get("User Name", "")).strip()
        year = str(row.get("Year", "")).strip()
        month = str(row.get("Month", "")).strip()
        if not user or not year or not month:
            continue
        scores = {}
        for kpi in KPIS:
            n = kpi["name"]
            scores[n] = {
                "value": row.get(f"{n} Value", ""),
                "label": row.get(f"{n} Slab", ""),
                "score": row.get(f"{n} Score", 0) or 0,
                "pct": _pct_from_label(row.get(f"{n} Slab", "")),
            }
        db.setdefault(user, {}).setdefault(year, {})[month] = scores

    _load_cache["data"] = db
    _load_cache["ts"] = now
    return db


def _pct_from_label(label):
    if label == "Slab 1 — 100%":
        return 1.0
    if label == "Slab 2 — 80%":
        return 0.8
    return 0.0


def _save_entry_sheets(username, year, month, scores):
    ws = _worksheet()
    all_values = ws.get_all_values()

    row_data = [username, str(year), month]
    total = 0
    for kpi in KPIS:
        d = scores[kpi["name"]]
        row_data += [d["value"], d["label"], d["score"]]
        total += d["score"]
    row_data.append(total)

    target_row = None
    for i, r in enumerate(all_values[1:], start=2):
        if len(r) >= 3 and r[0] == username and r[1] == str(year) and r[2] == month:
            target_row = i
            break

    if target_row:
        ws.update(f"A{target_row}", [row_data])
    else:
        ws.append_row(row_data)

    _invalidate_cache()
