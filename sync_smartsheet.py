"""
PMT Dashboard - Smartsheet Sync Script
=======================================
Fetches data from Smartsheet and writes to:
  1. Anthropic shared storage (for Claude Project artifact)
  2. dashboard-data.json (for MS Teams / SharePoint HTML)

Usage:
  python sync_smartsheet.py

Environment variables (set in GitHub Secrets or .env file):
  SMARTSHEET_TOKEN   - Your Smartsheet API access token
  ANTHROPIC_API_KEY  - Your Anthropic API key
  PROJECT_ID         - Your Claude Project ID
"""

import os, json, time, requests
from datetime import datetime, timezone

# ── Configuration ─────────────────────────────────────────────────────────────
SMARTSHEET_TOKEN = os.environ.get("SMARTSHEET_TOKEN", "YOUR_SMARTSHEET_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "YOUR_ANTHROPIC_API_KEY")
PROJECT_ID = os.environ.get("PROJECT_ID", "019d7781-f53c-74f9-8021-688603830be0")

SHEET_NAMES = {
    "wp":   "03. PMT  Workplan",
    "raid": "05. PMT [Project] RAID Log",
    "req":  "03. PMT - Requirements Repository",
    "test": "110. Test Scenarios",
    "cap":  "07. SAP Tech Sprint Capacity Management",
}

SS_BASE = "https://api.smartsheet.com/2.0"
SS_HEADERS = {
    "Authorization": f"Bearer {SMARTSHEET_TOKEN}",
    "Content-Type": "application/json",
}

ANTHROPIC_BASE = "https://api.anthropic.com"
ANTHROPIC_HEADERS = {
    "x-api-key": ANTHROPIC_API_KEY,
    "anthropic-version": "2023-06-01",
    "content-type": "application/json",
}

# Storage keys — must match dashboard JSX
STORAGE_KEYS = {
    "wp":     "pmt3_wp",
    "raid":   "pmt3_raid",
    "req":    "pmt3_req",
    "test":   "pmt3_test",
    "cap":    "pmt3_cap",
    "fnames": "pmt3_fnames",
    "meta":   "pmt3_meta",
}

# ── Columns to keep per sheet (slims payload) ─────────────────────────────────
KEEP_COLS = {
    "wp": [
        "Row ID","Lvl","Parent","Children",
        "Activity Grp - Lvl 1","Activity Grp - Lvl 2","Activity Grp - Lvl 3",
        "Activity Grp - Lvl 4","Activity Grp - Lvl 5","Activity Grp - Lvl 6",
        "Task Name","Default Status","Status","% Complete","Start","Finish","End Date",
        "Workstream","Support","Primary Owner","Secondary Owner","Comments",
    ],
    "raid": [
        "Type","Category","Status","Description","Title","Summary",
        "Primary Owner","Owner","Assignee","Priority","Severity",
        "Component","Workstream","Area","Team","Primary Team",
        "Comment","Comments","Resolution","Due Date","Target Date",
        "ID","RAID ID","Item ID","Experience","Topic","Critical Path",
        "Change Request Analysis","Status of Decision Acceptance (PMO)","Hours Estimate",
    ],
    "req": [
        "User Story","Req Id","Business Requirements","Acceptance Criteria",
        "PM Experience","Status",
        "User Story Review Status (D&A)",
        "Build Cycle (Playback)","Build Cycle","Targeted Closure Sprint",
        "Sub Process",
        "Functional Status Master List","Technical Status Master List",
        "Build Management Comments","Priority",
    ],
    "test": [
        "Scenarios","Scenario Id","SubProcess","Process Step ID","Step Description","Persona",
        "Estimated Test Cases","Primary User Story Ids","SIT Planned Testing",
        "Test Scenario Review SIT Plan","Sprint Build Plan",
        "Review Status (Functional)","Review Status (Technical)",
        "Review Status (Consulting SD)","Review Status (DT)",
        "Review Status (D&A)","Review Status (PMT SD)","Review Status (PM)",
    ],
    "cap": None,  # keep all columns
}

MAX_BYTES = 5 * 1024 * 1024  # 5MB per key

# ── Smartsheet helpers ────────────────────────────────────────────────────────
def list_sheets():
    """Return dict of sheet_name -> sheet_id"""
    r = requests.get(f"{SS_BASE}/sheets", headers=SS_HEADERS, params={"includeAll": True})
    r.raise_for_status()
    return {s["name"]: s["id"] for s in r.json().get("data", [])}

def fetch_sheet(sheet_id):
    """Fetch all rows from a sheet, return as list of dicts"""
    print(f"  Fetching sheet {sheet_id}...")
    r = requests.get(
        f"{SS_BASE}/sheets/{sheet_id}",
        headers=SS_HEADERS,
        params={"include": "rowPermalink", "pageSize": 10000},
    )
    r.raise_for_status()
    data = r.json()

    columns = {col["id"]: col["title"] for col in data.get("columns", [])}
    rows = []
    for row in data.get("rows", []):
        record = {}
        for cell in row.get("cells", []):
            col_name = columns.get(cell["columnId"], str(cell["columnId"]))
            val = cell.get("displayValue", cell.get("value", None))
            record[col_name] = val
        rows.append(record)
    return rows

def slim_rows(rows, keep_cols):
    """Keep only the columns the dashboard needs"""
    if not keep_cols:
        return rows
    return [
        {c: (row.get(c, "") if row.get(c) is not None else "") for c in keep_cols}
        for row in rows
    ]

# ── Anthropic shared storage helpers ─────────────────────────────────────────
def storage_set(key, value, shared=True):
    """Write a value to Anthropic project shared storage"""
    url = f"{ANTHROPIC_BASE}/v1/projects/{PROJECT_ID}/storage"
    payload = {"key": key, "value": value, "shared": shared}
    r = requests.post(url, headers=ANTHROPIC_HEADERS, json=payload)
    if r.status_code not in (200, 201):
        print(f"  WARNING: storage set failed for {key}: {r.status_code} {r.text[:200]}")
        return False
    return True

# ── Main sync ─────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("PMT Dashboard — Smartsheet Sync")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # 1. List all sheets
    print("\n[1/3] Discovering sheets...")
    all_sheets = list_sheets()
    print(f"  Found {len(all_sheets)} sheets in your Smartsheet account")

    # 2. Fetch each sheet
    print("\n[2/3] Fetching sheet data...")
    sheets_data = {}
    fnames = {}
    for key, name in SHEET_NAMES.items():
        if name not in all_sheets:
            print(f"  WARNING: Sheet '{name}' not found — skipping")
            continue
        sheet_id = all_sheets[name]
        rows = fetch_sheet(sheet_id)
        slimmed = slim_rows(rows, KEEP_COLS.get(key))
        sheets_data[key] = slimmed
        fnames[key] = name
        print(f"  ✓ {name}: {len(rows)} rows → {len(slimmed)} kept")

    # 3. Write to Anthropic shared storage
    print("\n[3/3] Writing to shared storage...")
    meta = {
        "lastSync": datetime.now(timezone.utc).isoformat(),
        "rowCounts": {k: len(v) for k, v in sheets_data.items()},
    }

    for key, rows in sheets_data.items():
        sheet_json = json.dumps({key: rows})
        size_kb = len(sheet_json.encode()) / 1024
        if len(sheet_json.encode()) > MAX_BYTES:
            print(f"  WARNING: {key} is {size_kb:.0f}KB — exceeds 5MB limit, skipping")
            continue
        ok = storage_set(STORAGE_KEYS[key], sheet_json)
        status = "✓" if ok else "✗"
        print(f"  {status} {key}: {size_kb:.0f}KB written to shared storage")

    storage_set(STORAGE_KEYS["fnames"], json.dumps(fnames))
    storage_set(STORAGE_KEYS["meta"], json.dumps(meta))

    # Also write combined JSON for MS Teams / SharePoint HTML
    combined = {
        "meta": meta,
        "fnames": fnames,
        **sheets_data,
    }
    combined_json = json.dumps(combined)
    size_kb = len(combined_json.encode()) / 1024
    with open("dashboard-data.json", "w") as f:
        f.write(combined_json)
    print(f"\n  ✓ dashboard-data.json written ({size_kb:.0f}KB) — upload this to SharePoint")

    print("\n" + "=" * 60)
    print("✅ Sync complete!")
    print(f"   Sheets synced: {list(sheets_data.keys())}")
    print(f"   Last sync: {meta['lastSync']}")
    print("=" * 60)

if __name__ == "__main__":
    main()
