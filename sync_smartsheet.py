"""
PMT Dashboard - Smartsheet Sync Script
=======================================
Fetches data from Smartsheet and writes dashboard-data.json
to the GitHub repo so GitHub Pages serves it to the dashboard.

Environment variables (set in GitHub Secrets):
  SMARTSHEET_TOKEN   - Your Smartsheet API access token
  GITHUB_TOKEN       - Personal Access Token with repo Contents write access
  GITHUB_USER        - GitHub username (default: sarmittal)
  GITHUB_REPO        - GitHub repo name (default: pmt-dashboard-sync)
"""

import os, json, base64, requests
from datetime import datetime, timezone

# ── Configuration ─────────────────────────────────────────────────────────────
SMARTSHEET_TOKEN = os.environ.get("SMARTSHEET_TOKEN", "YOUR_SMARTSHEET_TOKEN")
GITHUB_TOKEN     = os.environ.get("GITHUB_TOKEN", "YOUR_GITHUB_TOKEN")
GITHUB_USER      = os.environ.get("GITHUB_USER", "sarmittal")
GITHUB_REPO      = os.environ.get("GITHUB_REPO", "pmt-dashboard-sync")
OUTPUT_FILE      = "dashboard-data.json"

SHEET_NAMES = {
    "wp":   "03. PMT  Workplan",
    "raid": "05. PMT [Project] RAID Log",
    "req":  "03. PMT - Requirements Repository",
    "test": "110. Test Scenarios",
    "cap":  "07. SAP Tech Sprint Capacity Management",
}

SS_BASE    = "https://api.smartsheet.com/2.0"
SS_HEADERS = {"Authorization": f"Bearer {SMARTSHEET_TOKEN}", "Content-Type": "application/json"}

GH_BASE    = "https://api.github.com"
GH_HEADERS = {"Authorization": f"Bearer {GITHUB_TOKEN}", "Accept": "application/vnd.github+json", "Content-Type": "application/json"}

# Columns to keep per sheet (keeps payload small)
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
        "PM Experience","Status","User Story Review Status (D&A)",
        "Build Cycle (Playback)","Build Cycle","Targeted Closure Sprint",
        "Sub Process","Functional Status Master List","Technical Status Master List",
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

# ── Smartsheet helpers ────────────────────────────────────────────────────────
def list_sheets():
    r = requests.get(f"{SS_BASE}/sheets", headers=SS_HEADERS, params={"includeAll": True})
    r.raise_for_status()
    return {s["name"]: s["id"] for s in r.json().get("data", [])}

def fetch_sheet(sheet_id):
    r = requests.get(f"{SS_BASE}/sheets/{sheet_id}", headers=SS_HEADERS, params={"pageSize": 10000})
    r.raise_for_status()
    data = r.json()
    columns = {col["id"]: col["title"] for col in data.get("columns", [])}
    rows = []
    for row in data.get("rows", []):
        record = {}
        for cell in row.get("cells", []):
            col_name = columns.get(cell["columnId"], str(cell["columnId"]))
            record[col_name] = cell.get("displayValue", cell.get("value", "")) or ""
        rows.append(record)
    return rows

def slim_rows(rows, keep_cols):
    if not keep_cols:
        return rows
    return [{c: row.get(c, "") or "" for c in keep_cols} for row in rows]

# ── GitHub helpers ────────────────────────────────────────────────────────────
def get_file_sha(path):
    """Get current SHA of file in repo (needed to update it)"""
    url = f"{GH_BASE}/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/{path}"
    r = requests.get(url, headers=GH_HEADERS)
    if r.status_code == 404:
        return None  # file doesn't exist yet
    r.raise_for_status()
    return r.json().get("sha")

def write_file_to_github(path, content_str, commit_message):
    """Create or update a file in the GitHub repo"""
    url = f"{GH_BASE}/repos/{GITHUB_USER}/{GITHUB_REPO}/contents/{path}"
    sha = get_file_sha(path)
    content_b64 = base64.b64encode(content_str.encode()).decode()
    payload = {
        "message": commit_message,
        "content": content_b64,
    }
    if sha:
        payload["sha"] = sha  # required for updates
    r = requests.put(url, headers=GH_HEADERS, json=payload)
    r.raise_for_status()
    return r.json()

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("PMT Dashboard — Smartsheet Sync")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # 1. List sheets
    print("\n[1/3] Discovering Smartsheet sheets...")
    all_sheets = list_sheets()
    print(f"  Found {len(all_sheets)} sheets")

    # 2. Fetch each sheet
    print("\n[2/3] Fetching sheet data...")
    sheets_data = {}
    for key, name in SHEET_NAMES.items():
        if name not in all_sheets:
            print(f"  WARNING: '{name}' not found — skipping")
            continue
        rows = fetch_sheet(all_sheets[name])
        sheets_data[key] = slim_rows(rows, KEEP_COLS.get(key))
        print(f"  ✓ {name}: {len(rows)} rows")

    # 3. Write dashboard-data.json to GitHub repo
    print("\n[3/3] Writing dashboard-data.json to GitHub...")
    meta = {
        "lastSync": datetime.now(timezone.utc).isoformat(),
        "rowCounts": {k: len(v) for k, v in sheets_data.items()},
    }
    combined = {"meta": meta, **sheets_data}
    combined_json = json.dumps(combined)
    size_kb = len(combined_json.encode()) / 1024
    print(f"  Payload size: {size_kb:.0f}KB")

    result = write_file_to_github(
        OUTPUT_FILE,
        combined_json,
        f"chore: sync dashboard data {meta['lastSync'][:10]}"
    )
    print(f"  ✓ dashboard-data.json committed to repo")
    print(f"  ✓ Will be live at: https://{GITHUB_USER}.github.io/{GITHUB_REPO}/{OUTPUT_FILE}")

    print("\n" + "=" * 60)
    print("✅ Sync complete!")
    print(f"   Sheets: {list(sheets_data.keys())}")
    print(f"   Last sync: {meta['lastSync']}")
    print("=" * 60)

if __name__ == "__main__":
    main()
