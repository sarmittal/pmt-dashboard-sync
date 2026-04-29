/**
 * Smartsheet API client — fetches all 5 PMT sheets and returns
 * the same JSON shape the dashboard expects:
 *   { meta, wp, raid, req, test, cap, ec }
 *
 * Write-back:
 *   updateRow(sheetKey, rowId, updates, token)
 *   updates = { "Column Name": value, ... }
 */

import fetch from "node-fetch";

const BASE = "https://api.smartsheet.com/2.0";

// Sheet IDs (confirmed)
export const SHEETS = {
  wp:   "1792763851919236",
  raid: "491793142468484",
  req:  "1761775495106436",
  test: "2362069488717700",
  cap:  "1662804185534340",
  ec:   "1093126390239108",
};

// Columns allowed for write-back per sheet.
// Add column names here to enable inline editing in the UI.
export const EDITABLE = {
  wp: [
    "% Complete",
    "Comments",
  ],
  raid: [
    "Description",
    "Comments/Resolution History",
    "Critical Path",
    "RAID Due Date",
    "Tag",
    "Targeted Build Sprint",
  ],
};

// Columns to keep per sheet (keeps payload small, matches dashboard parser expectations)
const KEEP = {
  wp: [
    "Row ID","Lvl","Parent","Children",
    "Activity Grp - Lvl 0","Activity Grp - LVL 0",
    "Activity Grp - Lvl 1","Activity Grp - Lvl 2","Activity Grp - Lvl 3",
    "Activity Grp - Lvl 4","Activity Grp - Lvl 5","Activity Grp - Lvl 6",
    "Task Name","Default Status","Status","% Complete","Start","Finish","End Date",
    "Workstream","Support","Primary Owner","Secondary Owner","Comments",
  ],
  raid: null, // keep all columns — Tag and other fields vary by sheet configuration
  req: [
    "User Story", "Req Id", "Business Requirements", "Acceptance Criteria",
    "PM Experience", "User Story Review Status (D&A)",
    "Build Cycle (Playback)", "Targeted Closure Sprint", "Sub Process",
    "Functional Build Status", "Tech Build Status",
    "Build Management Comments", "User Story Derived Status", "Priority",
    "Test Script/Test Scenario",
    // Traceability tab — build approach & tags
    "Tags",
    "SF", "SF Design POV", "SF Configuration Workbook Reference",
    "BTP", "BTP Design POV",
    "CI", "CI/API POV",
    "AI", "AI POV",
    "Tech lean Spec Reference",
    "SNOW POV",
  ],
  test: null, // keep all — many reviewer/feedback/due-date columns needed for redesigned tab
  cap:  null, // keep all
  ec:   null, // keep all — EC Classes sheet, small number of rows
};

async function fetchRowAttachments(sheetId, token) {
  try {
    const res = await fetch(`${BASE}/sheets/${sheetId}/attachments?pageSize=500`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) return {};
    const data = await res.json();
    const map = {};
    for (const att of data.data || []) {
      // Accept any non-FILE attachment that has a url (LINK, ONEDRIVE, GOOGLE_DRIVE, etc.)
      if (att.parentType === "ROW" && att.url && att.attachmentType !== "FILE" && !map[att.parentId]) {
        map[att.parentId] = att.url;
      }
    }
    return map;
  } catch (e) {
    console.warn("[smartsheet] attachments fetch failed:", e.message);
    return {};
  }
}

async function fetchSheet(sheetId, token, fetchAttachments = false) {
  const url = `${BASE}/sheets/${sheetId}?pageSize=10000&include=rowPermalink`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Smartsheet ${res.status} for sheet ${sheetId}: ${body.slice(0, 200)}`);
  }
  const data = await res.json();

  // id→name for reading; name→id stored for write-back; PICKLIST options for dropdowns
  const columnNameById = {};
  const columnIdByName = {};
  const columnOptions  = {};
  for (const col of data.columns || []) {
    columnNameById[col.id] = col.title;
    columnIdByName[col.title] = col.id;
    if (col.type === "PICKLIST" && Array.isArray(col.options) && col.options.length) {
      columnOptions[col.title] = col.options;
    }
  }

  const attachMap = fetchAttachments ? await fetchRowAttachments(sheetId, token) : {};

  const rows = [];
  for (const row of data.rows || []) {
    const record = {
      _rowId:         row.id,
      _permalink:     row.permalink || "",
      _attachmentUrl: attachMap[row.id] || "",
    };
    for (const cell of row.cells || []) {
      const name = columnNameById[cell.columnId] || String(cell.columnId);
      record[name] = cell.displayValue ?? cell.value ?? "";
    }
    rows.push(record);
  }
  return { rows, columnIdByName, columnOptions };
}

function slim(rows, cols) {
  if (!cols) return rows;
  return rows.map(row => {
    const out = {};
    // Always carry internal _ fields (rowId, permalink, attachmentUrl)
    for (const [k, v] of Object.entries(row)) {
      if (k.startsWith("_")) out[k] = v;
    }
    for (const c of cols) out[c] = row[c] ?? "";
    return out;
  });
}

// For the capacity sheet, remove "Sprint N Actual" and "Sprint N." columns so the
// client parser always maps sprint numbers to the plain "Sprint N" column.
function cleanCapacityColumns(rows) {
  if (!rows?.length) return rows;
  return rows.map(row => {
    const out = {};
    for (const [k, v] of Object.entries(row)) {
      const isSprint = /sprint/i.test(k);
      if (isSprint && (/actual/i.test(k) || k.trimEnd().endsWith("."))) continue;
      out[k] = v;
    }
    return out;
  });
}

// ── Write-back ────────────────────────────────────────────────────────────────

/**
 * Update a single row in Smartsheet.
 * @param {string} sheetKey - Key from SHEETS (e.g. "wp", "raid")
 * @param {number} rowId    - Smartsheet row ID
 * @param {Object} updates  - { "Column Name": newValue, ... }
 * @param {Object} colMaps  - columnMaps from cache ({ sheetKey: { colName: colId } })
 * @param {string} token    - Bearer token
 */
export async function updateRow(sheetKey, rowId, updates, colMaps, token) {
  const sheetId = SHEETS[sheetKey];
  if (!sheetId) throw new Error(`Unknown sheet key: ${sheetKey}`);

  const colMap = colMaps[sheetKey];
  if (!colMap) throw new Error(`No column map for sheet: ${sheetKey}`);

  const allowed = EDITABLE[sheetKey] || [];
  const cells = [];
  for (const [colName, value] of Object.entries(updates)) {
    if (!allowed.includes(colName)) throw new Error(`Column "${colName}" is not editable on sheet "${sheetKey}"`);
    const columnId = colMap[colName];
    if (!columnId) throw new Error(`Column "${colName}" not found in sheet "${sheetKey}"`);
    cells.push({ columnId, value: value === "" ? null : value });
  }

  if (!cells.length) return { ok: true, skipped: true };

  const res = await fetch(`${BASE}/sheets/${sheetId}/rows`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify([{ id: rowId, cells }]),
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Smartsheet update failed ${res.status}: ${body.slice(0, 200)}`);
  }
  return res.json();
}

// ── Read all sheets ───────────────────────────────────────────────────────────

export async function fetchAllSheets(token) {
  if (!token) throw new Error("SMARTSHEET_TOKEN is not set");

  const results = {};
  const columnMaps = {};
  const allColumnOptions = {};
  const errors = [];

  await Promise.allSettled(
    Object.entries(SHEETS).map(async ([key, id]) => {
      try {
        const { rows, columnIdByName, columnOptions } = await fetchSheet(id, token, key === "raid");
        const slimmed = slim(rows, KEEP[key]);
        results[key] = key === "cap" ? cleanCapacityColumns(slimmed) : slimmed;
        columnMaps[key] = columnIdByName;
        if (Object.keys(columnOptions).length) allColumnOptions[key] = columnOptions;
        console.log(`[smartsheet] ${key}: ${rows.length} rows`);
      } catch (err) {
        errors.push(`${key}: ${err.message}`);
        console.error(`[smartsheet] ${key} failed:`, err.message);
      }
    })
  );

  if (errors.length === Object.keys(SHEETS).length) {
    throw new Error(`All sheet fetches failed:\n${errors.join("\n")}`);
  }

  return {
    data: {
      meta: {
        lastSync: new Date().toISOString(),
        rowCounts: Object.fromEntries(Object.entries(results).map(([k, v]) => [k, v.length])),
        errors: errors.length ? errors : undefined,
      },
      ...results,
      columnOptions: allColumnOptions,
    },
    columnMaps,
  };
}
