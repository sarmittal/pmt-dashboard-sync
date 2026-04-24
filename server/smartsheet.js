/**
 * Smartsheet API client — fetches all 5 PMT sheets and returns
 * the same JSON shape the dashboard expects:
 *   { meta, wp, raid, req, test, cap }
 */

import fetch from "node-fetch";

const BASE = "https://api.smartsheet.com/2.0";

// Sheet IDs (confirmed)
const SHEETS = {
  wp:   "1792763851919236",
  raid: "491793142468484",
  req:  "1761775495106436",
  test: "2362069488717700",
  cap:  "1662804185534340",
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
  ],
  test: [
    "Scenarios","Scenario Id","SubProcess","Process Step ID","Step Description",
    "Persona","Estimated Test Cases","Primary User Story Ids","SIT Planned Testing",
    "Test Scenario Review SIT Plan","Sprint Build Plan",
    "Review Status (Functional)","Review Status (Technical)",
    "Review Status (Consulting SD)","Review Status (DT)","Review Status (D&A)",
    "Review Status (PMT SD)","Review Status (PM)",
  ],
  cap: null, // keep all
};

async function fetchRowAttachments(sheetId, token) {
  // Fetch all sheet attachments and return { map, raw } where:
  //   map: rowId → first non-FILE attachment URL for that row
  //   raw: the first 20 attachment objects (for diagnostics)
  try {
    const res = await fetch(`${BASE}/sheets/${sheetId}/attachments?pageSize=500`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) {
      console.warn("[smartsheet] attachments HTTP", res.status, "for sheet", sheetId);
      return { map: {}, raw: [] };
    }
    const data = await res.json();
    const map = {};
    for (const att of data.data || []) {
      // Accept any non-FILE attachment that has a url (LINK, ONEDRIVE, GOOGLE_DRIVE, etc.)
      // FILE attachments expose url only via GET /attachments/{id}, not in the list response.
      if (att.parentType === "ROW" && att.url && att.attachmentType !== "FILE" && !map[att.parentId]) {
        map[att.parentId] = att.url;
      }
    }
    console.log(`[smartsheet] attachments for ${sheetId}: ${(data.data||[]).length} total, ${Object.keys(map).length} row URLs mapped`);
    return { map, raw: (data.data || []).slice(0, 20) };
  } catch (e) {
    console.warn("[smartsheet] attachments fetch failed:", e.message);
    return { map: {}, raw: [] };
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
  const columns = {};
  for (const col of data.columns || []) columns[col.id] = col.title;

  // Fetch row URL attachments (e.g. SharePoint links attached to rows)
  const { map: attachMap, raw: attachRaw } = fetchAttachments
    ? await fetchRowAttachments(sheetId, token)
    : { map: {}, raw: [] };

  const rows = [];
  for (const row of data.rows || []) {
    const record = {
      _permalink:     row.permalink || "",
      _attachmentUrl: attachMap[row.id] || "",
    };
    for (const cell of row.cells || []) {
      const name = columns[cell.columnId] || String(cell.columnId);
      record[name] = cell.displayValue ?? cell.value ?? "";
    }
    rows.push(record);
  }
  return rows;
}

function slim(rows, cols) {
  if (!cols) return rows;
  return rows.map(row => {
    const out = {};
    for (const c of cols) out[c] = row[c] ?? "";
    return out;
  });
}

export async function fetchAllSheets(token) {
  if (!token) throw new Error("SMARTSHEET_TOKEN is not set");

  const results = {};
  const errors = [];

  await Promise.allSettled(
    Object.entries(SHEETS).map(async ([key, id]) => {
      try {
        // Fetch row attachments for RAID sheet only (SharePoint links)
        const rows = await fetchSheet(id, token, key === "raid");
        results[key] = slim(rows, KEEP[key]);
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
    meta: {
      lastSync: new Date().toISOString(),
      rowCounts: Object.fromEntries(Object.entries(results).map(([k, v]) => [k, v.length])),
      errors: errors.length ? errors : undefined,
    },
    ...results,
  };
}
