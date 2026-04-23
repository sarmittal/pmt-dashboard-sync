import React, { useState, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── THEME ───────────────────────────────────────────────────────────────────
const C = {
  bg: "#f0f2f5", white: "#ffffff", border: "#dde3ea",
  navy: "#162f50", navyLight: "#2a5298", accent: "#1565c0",
  gold: "#f5a623", onTrack: "#f5a623", delayed: "#c0392b",
  complete: "#1a73e8", inProgress: "#27ae60", notStarted: "#bdc3c7",
  blocked: "#8e44ad", green: "#27ae60", red: "#c0392b",
  yellow: "#e67e22", text: "#1a1a2e", muted: "#6b7280", headerBg: "#162f50",
};
const SC = {
  "Off Track": C.delayed, "On Track": C.onTrack, "Complete": C.complete,
  "Not Started": C.notStarted, "Not started": C.notStarted,
  "At Risk": C.yellow, "Blocked": C.blocked, "In Progress": C.inProgress,
  "Open": C.yellow, "Closed": C.green, "High": C.red, "Medium": C.yellow, "Low": C.green,
  "Delayed": C.delayed,
};

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const readXlsx = (file) => new Promise((res, rej) => {
  const r = new FileReader();
  r.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array" });
      const sheets = {};
      wb.SheetNames.forEach(n => { sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: "" }); });
      res(sheets);
    } catch (err) { rej(err); }
  };
  r.onerror = rej; r.readAsArrayBuffer(file);
});

const pct = (v) => {
  if (v == null || v === "") return null;
  const s = String(v).replace("%","").trim();
  if (s === "" || isNaN(Number(s))) return null;
  const n = Number(s);
  return n <= 1 ? Math.round(n * 100) : Math.round(n);
};
const today = new Date();
const daysUntil = (d) => { if (!d) return null; try { const dt = typeof d === "number" ? new Date((d - 25569) * 86400000) : new Date(d); return Math.round((dt - today) / 86400000); } catch { return null; } };
const fmtDate = (d) => { if (!d) return "—"; try { return (typeof d === "number" ? new Date((d - 25569) * 86400000) : new Date(d)).toLocaleDateString("en-US", { month: "short", day: "numeric" }); } catch { return "—"; } };

// ─── WORKSTREAM NAME ALIASES ─────────────────────────────────────────────────
// Merge rows with different spellings of the same workstream into one canonical name.
// Keys are lowercase; value is the canonical display name.
const WS_ALIASES = {
  "expectation framework":  "Expectations Framework",
  "expectations framework": "Expectations Framework",
};
const normaliseWs = name => {
  const key = String(name || "").toLowerCase().trim();
  return WS_ALIASES[key] || name;
};

// ─── PARSERS ─────────────────────────────────────────────────────────────────
function parseWorkplan(sheets) {
  const key = Object.keys(sheets).find(k => k.toLowerCase().includes("workplan")) || Object.keys(sheets)[0];
  const rows = sheets[key]; if (!rows?.length) return null;

  // Filter rows that have a valid Lvl 1 value (exclude blank/header rows)
  // Detect the actual Comments column key (handles spaces, case variations)
  const firstRow = rows[0] || {};
  const commentsKey = Object.keys(firstRow).find(k => k.trim().toLowerCase() === "comments") 
                   || Object.keys(firstRow).find(k => /comment/i.test(k))
                   || "Comments";
  console.log("[PMT] Comments key detected:", commentsKey, "sample:", String(rows.find(r => r[commentsKey])?.  [commentsKey] || "").slice(0,40));

  const validRows = rows.filter(r => {
    const l = String(r["Activity Grp - Lvl 1"] || r["Workstream"] || "").trim();
    return l.length > 0;
  }).map(r => ({
    ...r,
    // Normalise Comments — try detected key, then fallback to any key matching "comment"
    "Comments": (() => {
      const raw = r[commentsKey] ?? r["Comments"] ?? 
                  Object.entries(r).find(([k]) => /comment/i.test(k))?.[1] ?? "";
      const v = String(raw).trim();
      const EMPTY = ["nan", "NaN", "null", "undefined", "0"];
      return EMPTY.includes(v) ? "" : v;
    })()
  }));

  // Tech + testing rows for workplan/testing-specific tabs
  const tech = validRows.filter(r => {
    const l = String(r["Activity Grp - Lvl 1"] || r["Workstream"] || "").toLowerCase();
    return l.includes("technology") || l.includes("testing");
  });
  const testRows = validRows.filter(r => String(r["Activity Grp - Lvl 1"] || "").toLowerCase().includes("testing"));

  const getS = r => (r["Default Status"] || r["Status"] || "").toLowerCase();
  const isLeaf = r => !r["Children"] || Number(r["Children"]) === 0;
  const techLeaves = tech.filter(isLeaf);
  const offTrack  = techLeaves.filter(r => getS(r).includes("off track"));
  const onTrack   = techLeaves.filter(r => getS(r).includes("on track"));
  const complete  = techLeaves.filter(r => getS(r).includes("complete"));
  const notStarted= techLeaves.filter(r => getS(r).includes("not start"));
  const dueSoon = tech.filter(r => {
    const d = daysUntil(r["Finish"] || r["End Date"]);
    return d != null && d >= 0 && d <= 14 && !getS(r).includes("complete");
  }).sort((a, b) => daysUntil(a["Finish"]) - daysUntil(b["Finish"]));

  // componentMap built from LEAF rows only (Children=0) so counts match the heatmap
  const componentMap = {};
  validRows.forEach(r => {
    const isLeaf = !r["Children"] || Number(r["Children"]) === 0;
    if (!isLeaf) return; // skip parent/header rows
    const c = normaliseWs(String(r["Activity Grp - Lvl 1"] || r["Workstream"] || "Unknown").trim());
    if (!componentMap[c]) componentMap[c] = { total: 0, offTrack: 0, complete: 0, onTrack: 0, notStarted: 0, rows: [] };
    const s = getS(r);
    componentMap[c].total++;
    componentMap[c].rows.push(r);
    if (s.includes("off track")) componentMap[c].offTrack++;
    else if (s.includes("complete")) componentMap[c].complete++;
    else if (s.includes("on track")) componentMap[c].onTrack++;
    else componentMap[c].notStarted++;
  });

  const subMap = {};
  [...offTrack, ...onTrack].forEach(r => {
    const k = r["Activity Grp - Lvl 3"] || r["Activity Grp - Lvl 2"] || "Other";
    if (!subMap[k]) subMap[k] = { onTrack: 0, delayed: 0, rows: [] };
    subMap[k].rows.push(r);
    if (getS(r).includes("off track")) subMap[k].delayed++; else subMap[k].onTrack++;
  });

  return { total: validRows.length, offTrack, onTrack, complete, notStarted, dueSoon, testRows, componentMap, subMap, allRows: validRows };
}

function parseRaid(sheets) {
  const key = Object.keys(sheets)[0];
  const rows = sheets[key]; if (!rows?.length) return null;
  const ks = Object.keys(rows[0]);
  const K = {
    type:      ks.find(k => /^type$|category/i.test(k)) || ks[0],
    status:    ks.find(k => /status/i.test(k)),
    desc:      ks.find(k => /desc|title|summary/i.test(k)),
    owner:     ks.find(k => k === "Primary Owner") || ks.find(k => /primary.?owner/i.test(k)) || ks.find(k => /owner|assignee/i.test(k)),
    priority:  ks.find(k => /priority|severity/i.test(k)),
    component: ks.find(k => k === "Component") || ks.find(k => /component|workstream|area/i.test(k)),
    team:      ks.find(k => k === "Primary Team (Owner)") || ks.find(k => k === "Primary Team") || ks.find(k => k === "Team") || ks.find(k => /primary.?team|^team$/i.test(k)),
    comment:   ks.find(k => k === "Comments/Resolution History") || ks.find(k => k === "Comments/ Resolution History") || ks.find(k => k === "Comment") || ks.find(k => k === "Comments") || ks.find(k => k === "Resolution") || ks.find(k => /comment|resolution/i.test(k)),
    id:        ks.find(k => k === "RAID ID") || ks.find(k => k === "Item ID") || ks.find(k => k === "ID") || ks.find(k => /raid.?id|item.?id/i.test(k)) || ks.find(k => /^id$/i.test(k)),
    experience:ks.find(k => /experience/i.test(k)),
    topic:     ks.find(k => /topic/i.test(k)),
    critPath:  ks.find(k => /critical.?path/i.test(k)),
    date:      ks.find(k => k === "Due Date") || ks.find(k => /due.?date|target.?date/i.test(k)),
    crAnalysis: ks.find(k => k === "Change Request Analysis") || ks.find(k => /change.?request.?analysis/i.test(k)),
    crStatus:   ks.find(k => k === "Status of Decision Acceptance (PMO)") || ks.find(k => /status.?of.?decision|decision.?acceptance/i.test(k)) || ks.find(k => /pmo.?status/i.test(k)),
    crHours:    ks.find(k => k === "Hours Estimate") || ks.find(k => /hours?.?estimate/i.test(k)),
  };
  const byPriority = {}, byComponent = {}, byTeam = {};
  rows.forEach(r => {
    const s = String(r[K.status] || "").toLowerCase();
    // Skip Complete and Deferred items from charts
    if (s === "complete" || s === "deferred") return;
    const isD = s === "delayed";
    const grp = (map, key) => { const k = r[key] || "Unknown"; if (!map[k]) map[k] = { onTrack: 0, delayed: 0, rows: [] }; isD ? map[k].delayed++ : map[k].onTrack++; map[k].rows.push(r); };
    grp(byPriority, K.priority); grp(byComponent, K.component); grp(byTeam, K.team);
  });
  // RAID Status values: "On Track", "Delayed", "Complete"
  // Open   = Status != "Complete"  (On Track + Delayed)
  // Delayed = Status == "Delayed"
  const isComplete = r => String(r[K.status] || "").toLowerCase() === "complete";
  const isDeferred = r => String(r[K.status] || "").toLowerCase() === "deferred";
  const isDelayed  = r => String(r[K.status] || "").toLowerCase() === "delayed";
  const isOpen     = r => !isComplete(r) && !isDeferred(r);

  const open    = rows.filter(isOpen);
  const delayed = rows.filter(r => isDelayed(r) && !isDeferred(r));
  const deferred = rows.filter(isDeferred);

  // Open Issues = Type contains "Issue" AND Status != Complete AND not Deferred
  // Open Risks  = Type contains "Risk"  AND Status != Complete AND not Deferred
  const openIssues  = rows.filter(r => isOpen(r) && String(r[K.type]||"").toLowerCase().includes("issue"));
  const openRisks   = rows.filter(r => isOpen(r) && String(r[K.type]||"").toLowerCase().includes("risk"));

  const statusValues = Array.from(new Set(rows.map(r => String(r[K.status] || "")))).sort();

  // ── Change Request buckets ──────────────────────────────────────────────────
  // CR-warranted rows = those whose Change Request Analysis contains one of the trigger values
  const CR_ANALYSIS_VALUES = [
    "tech reviewed - change request needed (inform)",
    "tech reviewed - change request needed (for decisioning)",
    "sd reviewed - change request needed (inform)",
    "sd reviewed - change request needed (for decisioning)",
    "ocm reviewed - change request needed (inform)",
    "ocm reviewed - change request needed (for decisioning)",
  ];
  const isCR = r => { const v = String(r[K.crAnalysis] || "").toLowerCase().trim(); return CR_ANALYSIS_VALUES.some(t => v.includes(t.slice(0, 20))); };
  const crRows = rows.filter(isCR);

  const getCrStatus = r => String(r[K.crStatus] || "").trim();
  const getCrHours  = r => { const v = String(r[K.crHours] || "").replace(/[^0-9.]/g, ""); const n = parseFloat(v); return isNaN(n) ? 0 : Math.round(n); };
  const sumHours    = arr => arr.reduce((s, r) => s + getCrHours(r), 0);

  // Status buckets — Approved = "Approved" + "Inform-Accepted (Reviewed)"
  const crApproved  = crRows.filter(r => { const s = getCrStatus(r).toLowerCase(); return s.includes("approved") || s.includes("inform-accepted (reviewed)"); });
  const crRejected  = crRows.filter(r => getCrStatus(r).toLowerCase().includes("rejected"));
  const crDeferred  = crRows.filter(r => getCrStatus(r).toLowerCase().includes("deferred"));
  const crPending   = crRows.filter(r => { const s = getCrStatus(r).toLowerCase(); return s.includes("pending") || s.includes("inform-accepted(not reviewed)") || s.includes("inform-accepted (not reviewed)"); });

  const cr = {
    all: crRows,
    approved: crApproved, approvedHours: sumHours(crApproved),
    rejected: crRejected, rejectedHours: sumHours(crRejected),
    deferred: crDeferred, deferredHours: sumHours(crDeferred),
    pending:  crPending,  pendingHours:  sumHours(crPending),
    totalHours: sumHours(crRows),
  };

  return { total: rows.length, open, delayed, deferred, openIssues, openRisks, byPriority, byComponent, byTeam, items: rows, keys: K, statusValues, cr };
}

function parseRequirements(sheets) {
  const key = Object.keys(sheets)[0];
  const rows = sheets[key]; if (!rows?.length) return null;
  const ks = Object.keys(rows[0]);

  // Log detected columns for debugging
  console.log("[PMT] Req columns:", ks.slice(0, 20));

  const K = {
    story:           ks.find(k => k === "User Story") || ks.find(k => /^user.?story$/i.test(k)) || ks[0],
    reqId:           ks.find(k => k === "Req Id") || ks.find(k => /^req.?id$/i.test(k)) || ks.find(k => /requirement.?id|req.?#/i.test(k)),
    bizReq:          ks.find(k => k === "Business Requirements") || ks.find(k => /business.?req/i.test(k)),
    acceptance:      ks.find(k => k === "Acceptance Criteria") || ks.find(k => /acceptance.?criteria/i.test(k)),
    pmExperience:    ks.find(k => k === "PM Experience") || ks.find(k => /pm.?experience|experience/i.test(k)),
    status:          ks.find(k => k === "Status") || ks.find(k => /^status$/i.test(k)),
    // "User Story Review Status (D&A)" — used to filter Deprecated/Deferred rows
    derivedStatus:   ks.find(k => k === "User Story Review Status (D&A)")
                  || ks.find(k => /user.?story.?review.?status/i.test(k))
                  || ks.find(k => /review.?status.*d.?a/i.test(k))
                  || ks.find(k => /derived.?status|req.?review/i.test(k)),
    // "Build Cycle (Playback)" or "Build Cycle"
    sprint:          ks.find(k => k === "Build Cycle (Playback)") || ks.find(k => k === "Build Cycle") || ks.find(k => /build.?cycle/i.test(k)) || ks.find(k => /^sprint$/i.test(k)) || ks.find(k => /sprint/i.test(k)),
    // "Targeted Closure Sprint"
    closureSprint:   ks.find(k => k === "Targeted Closure Sprint") || ks.find(k => /target.?closure|closure.?sprint/i.test(k)),
    // Component = "Sub Process"
    component:       ks.find(k => k === "Sub Process") || ks.find(k => /sub.?process/i.test(k)) || ks.find(k => /component|module|feature/i.test(k)),
    // "Functional Build Status" (API name) or legacy "Functional Status Master List" (XLSX export name)
    funcBuildStatus: ks.find(k => k === "Functional Build Status") || ks.find(k => k === "Functional Status Master List") || ks.find(k => /functional.?build/i.test(k)) || ks.find(k => /functional.?status/i.test(k)),
    // "Tech Build Status" (API name) or legacy "Technical Status Master List" (XLSX export name)
    techBuildStatus: ks.find(k => k === "Tech Build Status") || ks.find(k => k === "Technical Status Master List") || ks.find(k => /tech.?build/i.test(k)) || ks.find(k => /technical.?status/i.test(k)),
    assignee:        ks.find(k => /assign|owner/i.test(k)),
    priority:        ks.find(k => /priority/i.test(k)),
    // "Build Management Comments" or similar
    buildComment:    ks.find(k => k === "Build Management Comments")
                  || ks.find(k => /build.?management.?comment/i.test(k))
                  || ks.find(k => /build.?mgmt.?comment/i.test(k))
                  || ks.find(k => /build.?comment/i.test(k)),
  };

  console.log("[PMT] Req key mapping:", K);

  // Exclude rows where User Story Review Status (D&A) is blank, "5. Deprecated" or "6. Deferred"
  const EXCLUDED = ["5. deprecated", "6. deferred", "deferred"];
  const isExcluded = r => {
    const v = String(r[K.derivedStatus] || "").toLowerCase().trim();
    if (!v || v === "nan" || v === "null" || v === "undefined") return true;
    return EXCLUDED.some(e => v.includes(e));
  };
  const activeRows = rows.filter(r => !isExcluded(r));
  console.log("[PMT] Req active rows:", activeRows.length, "of", rows.length);

  // Status buckets — use the worst of funcBuildStatus and techBuildStatus
  // (mirrors the Smartsheet consolidated status formula)
  // Fallback to Status column if func/tech not available
  const statusToBucket = s => {
    const v = String(s || "").toLowerCase().trim();
    if (!v || v === "nan" || v === "") return null;
    if (v.includes("block"))                                   return "blocked";
    if (v.includes("in progress") || v.includes("progress"))  return "inProgress";
    if (v.includes("partial"))                                 return "partial";
    if (v.includes("complete") && !v.includes("partial"))     return "complete";
    if (v.includes("n/a") || v === "na")                      return "na";
    if (v.includes("not started") || v.includes("not start")) return "notStarted";
    return "notStarted";
  };

  const BUCKET_RANK = { blocked:5, inProgress:4, partial:3, notStarted:2, complete:1, na:0 };

  const getStatusBucket = r => {
    const fb = statusToBucket(r[K.funcBuildStatus]);
    const tb = statusToBucket(r[K.techBuildStatus]);
    // Pick worst of func and tech
    if (fb && tb) return (BUCKET_RANK[fb]||0) >= (BUCKET_RANK[tb]||0) ? fb : tb;
    if (fb) return fb;
    if (tb) return tb;
    // Fallback to Status column
    return statusToBucket(r[K.status]) || "notStarted";
  };

  // Build sprint × component matrix
  const bySprint = {}, byComponent = {};
  const allSprints = new Set();

  activeRows.forEach(r => {
    const sp = String(r[K.sprint] || "No Sprint").trim();
    const c  = String(r[K.component] || "Other").trim();
    const bucket = getStatusBucket(r);
    allSprints.add(sp);

    if (!bySprint[sp]) bySprint[sp] = { complete:0, partial:0, inProgress:0, notStarted:0, blocked:0, na:0, total:0, rows:[] };
    if (!byComponent[c]) byComponent[c] = { complete:0, partial:0, inProgress:0, notStarted:0, blocked:0, na:0, total:0, rows:[] };

    bySprint[sp].rows.push(r);
    bySprint[sp].total++;
    bySprint[sp][bucket]++;

    byComponent[c].rows.push(r);
    byComponent[c].total++;
    byComponent[c][bucket]++;
  });

  // Per-component sprint breakdown: { compName: { sprintName: { complete, partial, ... } } }
  const compBySprint = {};
  activeRows.forEach(r => {
    const c  = String(r[K.component] || "Other").trim();
    const sp = String(r[K.sprint] || "No Sprint").trim();
    const bucket = getStatusBucket(r);
    if (!compBySprint[c]) compBySprint[c] = {};
    if (!compBySprint[c][sp]) compBySprint[c][sp] = { complete:0, partial:0, inProgress:0, notStarted:0, blocked:0, na:0, total:0 };
    compBySprint[c][sp].total++;
    compBySprint[c][sp][bucket]++;
  });

  // Per-component build status distributions
  const compBuildStatus = {};
  activeRows.forEach(r => {
    const c  = String(r[K.component] || "Other").trim();
    if (!compBuildStatus[c]) compBuildStatus[c] = { func:{}, tech:{} };
    const fb = String(r[K.funcBuildStatus] || "").trim();
    const tb = String(r[K.techBuildStatus] || "").trim();
    const EMPTY = ["", "nan", "NaN", "null", "undefined"];
    if (fb && !EMPTY.includes(fb)) compBuildStatus[c].func[fb] = (compBuildStatus[c].func[fb] || 0) + 1;
    if (tb && !EMPTY.includes(tb)) compBuildStatus[c].tech[tb] = (compBuildStatus[c].tech[tb] || 0) + 1;
  });

  // Overall totals by sprint (for Build Management tab)
  const sprintOrder = Array.from(allSprints).sort((a,b) => String(a).localeCompare(String(b)));

  const done       = activeRows.filter(r => getStatusBucket(r) === "complete");
  const partial    = activeRows.filter(r => getStatusBucket(r) === "partial");
  const inProg     = activeRows.filter(r => getStatusBucket(r) === "inProgress");
  const blocked    = activeRows.filter(r => getStatusBucket(r) === "blocked");
  const notStarted = activeRows.filter(r => getStatusBucket(r) === "notStarted");

  return {
    total: activeRows.length, done, partial, inProg, blocked, notStarted,
    bySprint, byComponent, compBySprint, compBuildStatus, sprintOrder,
    items: activeRows, keys: K
  };
}

function parseTestScenarios(sheets) {
  const key = Object.keys(sheets)[0];
  const rows = sheets[key]; if (!rows?.length) return null;
  const ks = Object.keys(rows[0]);

  const K = {
    id:          ks.find(k => /scenario.?id/i.test(k)),
    name:        ks.find(k => k === "Scenarios") || ks.find(k => /^scenarios?$/i.test(k)),
    subprocess:  ks.find(k => k === "SubProcess") || ks.find(k => /sub.?process/i.test(k)),
    processStep: ks.find(k => /process.?step.?id/i.test(k)),
    stepDesc:    ks.find(k => /step.?desc/i.test(k)),
    persona:     ks.find(k => /persona/i.test(k)),
    estCases:    ks.find(k => /estimated.?test.?cases/i.test(k)),
    storyIds:    ks.find(k => /primary.?user.?story/i.test(k)),
    sitPlan:     ks.find(k => k === "Test Scenario Review SIT Plan") || ks.find(k => /sit.?plan|sit.?review/i.test(k)),
    sprintPlan:  ks.find(k => k === "Sprint Build Plan") || ks.find(k => /sprint.?build/i.test(k)),
    funcStatus:  ks.find(k => k === "Review Status (Functional)") || ks.find(k => /review.?status.*functional/i.test(k)),
    techStatus:  ks.find(k => k === "Review Status (Technical)") || ks.find(k => /review.?status.*technical/i.test(k)),
    sdStatus:    ks.find(k => k === "Review Status (Consulting SD)") || ks.find(k => /review.?status.*consulting/i.test(k)),
    dtStatus:    ks.find(k => k === "Review Status (DT)") || ks.find(k => /review.?status.*\bdt\b/i.test(k)),
    daStatus:    ks.find(k => k === "Review Status (D&A)") || ks.find(k => /review.?status.*d.?a/i.test(k)),
    pmtStatus:   ks.find(k => k === "Review Status (PMT SD)") || ks.find(k => /review.?status.*pmt/i.test(k)),
    pmStatus:    ks.find(k => k === "Review Status (PM)") || ks.find(k => /review.?status.*\bpm\b/i.test(k)),
  };

  // Exclude deprecated / deferred / duplicate
  const EXCLUDED = ["5. deprecated", "6. deferred", "7. duplicate"];
  const isExcluded = r => {
    const v = String(r[K.funcStatus] || "").toLowerCase().trim();
    return EXCLUDED.some(e => v.includes(e));
  };
  const activeRows = rows.filter(r => !isExcluded(r));

  // Status normaliser — strip leading "N. " prefix
  const cleanStatus = s => {
    if (!s) return null;
    const v = String(s).trim();
    if (!v || v === "None" || v === "nan") return null;
    return v.replace(/^\d+\.\s*/, "").trim();
  };

  // Bucket func review status
  const STATUS_BUCKETS = {
    "reviewed":         "Reviewed",
    "ready for review": "Ready for Review",
    "not applicable":   "Not Applicable",
    "duplicate":        "Duplicate",
    "deprecated":       "Deprecated",
    "deferred":         "Deferred",
  };
  const bucketStatus = s => {
    if (!s) return "Unknown";
    const v = s.toLowerCase();
    for (const [k, b] of Object.entries(STATUS_BUCKETS)) { if (v.includes(k)) return b; }
    return s;
  };

  // Sprint label extraction — "5. S5 + PB..." -> "S5"
  const extractSprint = s => {
    if (!s) return null;
    const m = String(s).match(/s(\d+)/i);
    return m ? `S${m[1]}` : null;
  };

  // Group by SubProcess
  const bySubprocess = {};
  activeRows.forEach(r => {
    const sp = String(r[K.subprocess] || "Unknown").trim();
    if (!bySubprocess[sp]) bySubprocess[sp] = [];
    bySubprocess[sp].push(r);
  });

  // Group by Sprint
  const bySprint = {};
  activeRows.forEach(r => {
    const raw = String(r[K.sprintPlan] || "").trim();
    // may have multiple sprints in one cell
    const sprints = raw.split(/\n/).map(s => extractSprint(s)).filter(Boolean);
    const labels = sprints.length ? [...new Set(sprints)] : ["Unassigned"];
    labels.forEach(lbl => {
      if (!bySprint[lbl]) bySprint[lbl] = [];
      bySprint[lbl].push(r);
    });
  });

  // KPI counts
  const funcReviewed    = activeRows.filter(r => bucketStatus(cleanStatus(r[K.funcStatus])) === "Reviewed").length;
  const funcReadyReview = activeRows.filter(r => bucketStatus(cleanStatus(r[K.funcStatus])) === "Ready for Review").length;
  const totalEstCases   = activeRows.reduce((s, r) => s + (Number(r[K.estCases]) || 0), 0);

  // SIT plan distribution
  const sitCounts = {};
  activeRows.forEach(r => {
    const v = String(r[K.sitPlan] || "").trim();
    if (v && v !== "None" && v !== "nan") {
      v.split(/\n|,/).map(s => s.trim()).filter(Boolean).forEach(s => {
        sitCounts[s] = (sitCounts[s] || 0) + 1;
      });
    }
  });

  return {
    total: activeRows.length,
    totalEstCases: Math.round(totalEstCases),
    funcReviewed, funcReadyReview,
    bySubprocess, bySprint, sitCounts,
    activeRows, keys: K,
  };
}

function parseCapacity(sheets) {
  const key = Object.keys(sheets)[0];
  const rows = sheets[key]; if (!rows?.length) return null;
  const ks = Object.keys(rows[0]);
  const K = {
    resource: ks.find(k => /resource|person|name|team.?member/i.test(k)) || ks[0],
    sprint: ks.find(k => /sprint/i.test(k)),
    available: ks.find(k => /available|capacity/i.test(k)),
    planned: ks.find(k => /planned|allocated/i.test(k)),
    workstream: ks.find(k => /workstream|team|component/i.test(k)),
  };
  const bySprint = {};
  rows.forEach(r => {
    const sp = r[K.sprint] || "Overall";
    const avail = Number(r[K.available] || 0), planned = Number(r[K.planned] || 0);
    if (!bySprint[sp]) bySprint[sp] = { available: 0, planned: 0, rows: [] };
    bySprint[sp].available += avail; bySprint[sp].planned += planned; bySprint[sp].rows.push(r);
  });
  const sprintChart = Object.entries(bySprint).map(([name, d]) => ({ name, ...d, diff: d.available - d.planned }));
  return { total: rows.length, bySprint, sprintChart, items: rows, keys: K };
}

// ─── SHARED UI ───────────────────────────────────────────────────────────────
function UploadZone({ label, icon, loaded, onFile, hint, filename }) {
  const [drag, setDrag] = useState(false);
  const onDrop = useCallback(e => { e.preventDefault(); setDrag(false); const f = e.dataTransfer.files[0]; if (f) onFile(f); }, [onFile]);
  return (
    <label onDragOver={e => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)} onDrop={onDrop}
      style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 5, padding: "12px 10px",
        border: `2px dashed ${loaded ? C.green : drag ? C.accent : C.border}`, borderRadius: 7, cursor: "pointer", transition: "all .2s",
        background: loaded ? "#f0fdf4" : drag ? "#eff6ff" : C.white, minHeight: 82, flex: 1 }}>
      <input type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => { const f = e.target.files[0]; if (f) { onFile(f); e.target.value = ""; } }} />
      <span style={{ fontSize: 20 }}>{loaded ? "✅" : icon}</span>
      <span style={{ color: loaded ? C.green : C.text, fontWeight: 700, fontSize: 11, textAlign: "center" }}>{label}</span>
      <span style={{ color: C.muted, fontSize: 10, textAlign: "center" }}>{loaded ? (filename ? filename.slice(0, 28) : "Loaded") : hint}</span>
    </label>
  );
}

function KpiCard({ label, value, color, sub, onClick }) {
  const [hover, setHover] = useState(false);
  return (
    <div onClick={onClick} onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}
      style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 7, padding: "14px 16px",
        borderTop: `3px solid ${color}`, boxShadow: hover && onClick ? "0 4px 14px rgba(0,0,0,0.12)" : "0 1px 3px rgba(0,0,0,0.06)",
        cursor: onClick ? "pointer" : "default", transition: "box-shadow .15s" }}>
      <div style={{ color: C.muted, fontSize: 10, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 3 }}>{label}</div>
      <div style={{ color, fontSize: 26, fontWeight: 800, lineHeight: 1 }}>{value}</div>
      {sub && <div style={{ color: C.muted, fontSize: 10, marginTop: 3 }}>{sub}</div>}
      {onClick && <div style={{ color: C.accent, fontSize: 10, marginTop: 3 }}>Click for details →</div>}
    </div>
  );
}

function Card({ children, style }) {
  return <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 8, padding: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.06)", ...style }}>{children}</div>;
}

function SecTitle({ title, color }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 12 }}>
      <div style={{ width: 3, height: 14, background: color || C.accent, borderRadius: 2 }} />
      <span style={{ color: C.text, fontWeight: 700, fontSize: 12, textDecoration: "underline", textDecorationColor: color || C.accent }}>{title}</span>
    </div>
  );
}

function Leg({ items }) {
  return (
    <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginTop: 8 }}>
      {items.map(({ label, color }) => (
        <span key={label} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: C.muted }}>
          <span style={{ width: 9, height: 9, borderRadius: "50%", background: color, display: "inline-block" }} />{label}
        </span>
      ))}
    </div>
  );
}

// Horizontal stacked bars — Smartsheet style
// Single bar row — extracted as component so hooks are legal
function HSBarRow({ row, i, valueKeys, colors, max, onBarClick }) {
  const [hover, setHover] = useState(false);
  const total = valueKeys.reduce((s, k) => s + (Number(row[k]) || 0), 0);
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}
      onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}>
      <div style={{ minWidth: 155, maxWidth: 170, color: C.text, fontSize: 11, textAlign: "right",
        overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={row.name}>{row.name}</div>
      <div style={{ flex: 1, display: "flex", height: 20, borderRadius: 3, overflow: "hidden",
        cursor: onBarClick ? "pointer" : "default",
        outline: hover && onBarClick ? `2px solid ${C.accent}` : "none", transition: "outline .1s" }}
        onClick={() => onBarClick && onBarClick(row)}
        title={onBarClick ? "Click to drill down" : ""}>
        {valueKeys.map((k, ki) => (Number(row[k]) || 0) > 0 && (
          <div key={ki} title={`${k}: ${row[k]}`}
            style={{ width: `${((Number(row[k]) || 0) / max) * 100}%`, background: colors[ki],
              display: "flex", alignItems: "center", justifyContent: "center", transition: "width .3s" }}>
            {(Number(row[k]) || 0) > 1 && <span style={{ color: "#fff", fontSize: 10, fontWeight: 700, userSelect: "none" }}>{row[k]}</span>}
          </div>
        ))}
      </div>
      <div style={{ minWidth: 24, color: C.muted, fontSize: 11 }}>{total}</div>
    </div>
  );
}

function HSBar({ data, valueKeys, colors, onBarClick }) {
  const max = Math.max(...data.map(d => valueKeys.reduce((s, k) => s + (Number(d[k]) || 0), 0)), 1);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      {data.map((row, i) => (
        <HSBarRow key={i} row={row} i={i} valueKeys={valueKeys} colors={colors} max={max} onBarClick={onBarClick} />
      ))}
    </div>
  );
}

// Drill-Down Modal
function Modal({ title, rows, columns, onClose }) {
  if (!rows?.length) return null;
  const cols = columns || Object.keys(rows[0]).slice(0, 7);
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}
      onClick={onClose}>
      <div style={{ background: C.white, borderRadius: 10, width: "92%", maxWidth: 1000, maxHeight: "82vh", display: "flex", flexDirection: "column", boxShadow: "0 24px 60px rgba(0,0,0,0.35)" }}
        onClick={e => e.stopPropagation()}>
        <div style={{ background: C.headerBg, padding: "13px 20px", borderRadius: "10px 10px 0 0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ color: "#fff", fontWeight: 700, fontSize: 14 }}>{title} <span style={{ opacity: .6, fontWeight: 400 }}>({rows.length} items)</span></div>
          <button onClick={onClose} style={{ background: "rgba(255,255,255,0.15)", border: "none", color: "#fff", borderRadius: 5, padding: "4px 12px", cursor: "pointer", fontSize: 13, fontWeight: 600 }}>✕</button>
        </div>
        <div style={{ overflowY: "auto", flex: 1 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead style={{ position: "sticky", top: 0, background: "#f0f4f8", zIndex: 1 }}>
              <tr>{cols.map(c => <th key={c} style={{ textAlign: "left", padding: "9px 12px", color: C.muted, fontSize: 11, fontWeight: 700, borderBottom: `2px solid ${C.border}`, whiteSpace: "nowrap" }}>{c}</th>)}</tr>
            </thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={i} style={{ background: i % 2 === 0 ? C.white : "#f9fafb", borderBottom: `1px solid ${C.border}` }}>
                  {cols.map(c => {
                    const v = String(r[c] || "—"); const col = /status/i.test(c) ? SC[v] : null;
                    return (
                      <td key={c} style={{ padding: "7px 12px", color: C.text, maxWidth: 260, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={v}>
                        {col ? <span style={{ background: col + "20", color: col, border: `1px solid ${col}40`, borderRadius: 4, padding: "2px 7px", fontSize: 10, fontWeight: 700 }}>{v}</span> : v.slice(0, 65)}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function Empty({ label }) {
  return (
    <div style={{ textAlign: "center", padding: "60px", color: C.muted }}>
      <div style={{ fontSize: 40, marginBottom: 10 }}>📂</div>
      <div style={{ fontSize: 13 }}>{label}</div>
    </div>
  );
}

const WP_COLS = ["Task Name", "Activity Grp - Lvl 1", "Activity Grp - Lvl 3", "Default Status", "% Complete", "Finish", "Primary Owner"];

function ActivityTable({ rows }) {
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <thead><tr style={{ background: "#f0f4f8" }}>
          {WP_COLS.map(c => <th key={c} style={{ textAlign: "left", padding: "8px 10px", color: C.muted, fontSize: 11, fontWeight: 700, borderBottom: `2px solid ${C.border}`, whiteSpace: "nowrap" }}>{c}</th>)}
        </tr></thead>
        <tbody>
          {rows.slice(0, 60).map((r, i) => {
            const s = r["Default Status"] || r["Status"] || ""; const sc = SC[s] || C.muted;
            const p = pct(r["% Complete"] ?? r["% complete"]);
            return (
              <tr key={i} style={{ borderBottom: `1px solid ${C.border}`, background: i % 2 === 0 ? C.white : "#f9fafb" }}>
                <td style={{ padding: "7px 10px", color: C.text, maxWidth: 280 }} title={r["Task Name"]}>{String(r["Task Name"] || "—").slice(0, 65)}</td>
                <td style={{ padding: "7px 10px", color: C.muted, fontSize: 11 }}>{String(r["Activity Grp - Lvl 1"] || "").replace("Technology - ", "").slice(0, 22)}</td>
                <td style={{ padding: "7px 10px", color: C.muted, fontSize: 11 }}>{String(r["Activity Grp - Lvl 3"] || r["Activity Grp - Lvl 2"] || "").slice(0, 30)}</td>
                <td style={{ padding: "7px 10px" }}><span style={{ background: sc + "20", color: sc, border: `1px solid ${sc}40`, borderRadius: 4, padding: "2px 7px", fontSize: 10, fontWeight: 700 }}>{s || "—"}</span></td>
                <td style={{ padding: "7px 10px", color: C.text }}>{p != null ? `${p}%` : "—"}</td>
                <td style={{ padding: "7px 10px", color: C.muted }}>{fmtDate(r["Finish"] || r["End Date"])}</td>
                <td style={{ padding: "7px 10px", color: C.muted }}>{String(r["Primary Owner"] || "—").slice(0, 20)}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
      {rows.length > 60 && <div style={{ color: C.muted, fontSize: 11, textAlign: "center", padding: 6 }}>Showing 60 of {rows.length}</div>}
    </div>
  );
}

// ─── TABS ────────────────────────────────────────────────────────────────────
const TABS = [
  { id: "executive",   label: "Executive Summary" },
  { id: "workplan",    label: "Workplan" },
  { id: "raid",        label: "RAID Analysis" },
  { id: "cr",          label: "Change Requests" },
  { id: "backlog",     label: "Backlog" },
  { id: "overview",    label: "Program Health Overview" },
  { id: "scorecard",   label: "Component Scorecard" },
  { id: "testing",     label: "Test Scenarios" },
];

// ─── STORAGE ─────────────────────────────────────────────────────────────────
const KEYS = { wp: "pmt3_wp", raid: "pmt3_raid", req: "pmt3_req", cap: "pmt3_cap", test: "pmt3_test", fnames: "pmt3_fnames", meta: "pmt3_meta" };

const WP_COLS_KEEP = [
  "Row ID","Lvl","Parent","Children",
  "Activity Grp - Lvl 1","Activity Grp - Lvl 2","Activity Grp - Lvl 3","Activity Grp - Lvl 4",
  "Activity Grp - Lvl 5","Activity Grp - Lvl 6",
  "Task Name","Default Status","Status","% Complete","Start","Finish","End Date",
  "Workstream","Support","Primary Owner","Secondary Owner","Comments"
];

const RAID_COLS_KEEP = [
  "Type","Category","Status","Description","Title","Summary",
  "Primary Owner","Owner","Assignee","Priority","Severity",
  "Component","Workstream","Area","Team","Primary Team",
  "Comment","Comments","Resolution","Due Date","Target Date",
  "ID","RAID ID","Item ID","Experience","Topic","Critical Path",
  "Change Request Analysis","Status of Decision Acceptance (PMO)","Hours Estimate",
];

const TEST_COLS_KEEP = [
  "Scenarios","Scenario Id","SubProcess","Process Step ID","Step Description","Persona",
  "Estimated Test Cases","Primary User Story Ids","SIT Planned Testing",
  "Test Scenario Review SIT Plan","Sprint Build Plan",
  "Review Status (Functional)","Review Status (Technical)","Review Status (Consulting SD)",
  "Review Status (DT)","Review Status (D&A)","Review Status (PMT SD)","Review Status (PM)",
];

// Requirements columns the dashboard actually uses — slim aggressively to stay under 5MB
const REQ_COLS_KEEP = [
  "User Story","Req Id","Business Requirements","Acceptance Criteria",
  "PM Experience","Status",
  "User Story Review Status (D&A)",
  "Build Cycle (Playback)","Build Cycle","Targeted Closure Sprint",
  "Sub Process",
  "Functional Status Master List","Technical Status Master List",
  "Build Management Comments",
  "Priority",
];

function slimRows(sheets, keepCols) {
  const out = {};
  Object.entries(sheets).forEach(([name, rows]) => {
    out[name] = (rows || []).map(row => {
      if (!keepCols) return row;
      const s = {};
      keepCols.forEach(c => {
        const v = (c in row) ? row[c] : "";
        // Convert NaN/null/undefined to empty string
        s[c] = (v === null || v === undefined || (typeof v === "number" && isNaN(v))) ? "" : v;
      });
      return s;
    });
  });
  return out;
}

// Try window.storage first, fall back to sessionStorage
async function persist(key, sheets, keepCols) {
  try {
    const data = slimRows(sheets, keepCols);
    const json = JSON.stringify(data);
    const kb = Math.round(json.length / 1024);
    console.log("[PMT] Saving", key, kb + "KB");
    if (json.length > 5 * 1024 * 1024) {
      console.warn("[PMT] Skipping storage for", key, "— too large:", kb + "KB (> 5MB)");
      return;
    }
    if (window.storage && typeof window.storage.set === "function") {
      try {
        const r = await window.storage.set(key, json);
        console.log("[PMT] window.storage.set result:", key, r ? "ok" : "failed");
      } catch(e) { console.warn("[PMT] window.storage.set failed for", key, e?.message); }
    }
    try { sessionStorage.setItem(key, json); } catch(e) { console.warn("[PMT] sessionStorage failed for", key); }
  } catch(e) { console.error("[PMT] persist error:", key, e); }
}

async function restore(key) {
  try {
    // Try window.storage first
    if (window.storage && typeof window.storage.get === "function") {
      const r = await window.storage.get(key);
      if (r && r.value) {
        console.log("[PMT] window.storage restored", key, Math.round(r.value.length/1024) + "KB");
        return JSON.parse(r.value);
      }
    }
    // Fall back to sessionStorage
    const s = sessionStorage.getItem(key);
    if (s) {
      console.log("[PMT] sessionStorage restored", key, Math.round(s.length/1024) + "KB");
      return JSON.parse(s);
    }
    console.log("[PMT] Nothing found for", key);
    return null;
  } catch(e) { console.error("[PMT] restore error:", key, e); return null; }
}

async function clearAll() {
  try { if (window.storage) { Object.values(KEYS).forEach(k => window.storage.delete(k).catch(()=>{})); } } catch(e) {}
  Object.values(KEYS).forEach(k => { try { sessionStorage.removeItem(k); } catch(e) {} });
}

// ─── APP ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("executive");
  const [modal, setModal] = useState(null);
  const [wp, setWp] = useState(null);
  const [raid, setRaid] = useState(null);
  const [req, setReq] = useState(null);
  const [cap, setCap] = useState(null);
  const [test, setTest] = useState(null);
  const [fnames, setFnames] = useState({});
  const [rawSheets, setRawSheets] = useState({});
  const [storageLoaded, setStorageLoaded] = useState(false);
  const [snapshotText, setSnapshotText] = useState(null);
  const [syncMeta, setSyncMeta] = useState(null); // { lastSync, source }
  const [isLoading, setIsLoading] = useState(false);
  const [refreshing, setRefreshing] = useState(false);

  const [storageStatus, setStorageStatus] = useState({});

  // ── applyApiData: populate state from /api/data response ────────────────
  const applyApiData = useCallback((apiJson) => {
    if (apiJson.wp?.length)   { const s={"03. PMT  Workplan":apiJson.wp};            setWp(parseWorkplan(s));           setRawSheets(p=>({...p,wp:s})); }
    if (apiJson.raid?.length) { const s={"05. PMT [Project] RAID Log":apiJson.raid}; setRaid(parseRaid(s));             setRawSheets(p=>({...p,raid:s})); }
    if (apiJson.req?.length)  { const s={"03. PMT - Requirements Repository":apiJson.req}; setReq(parseRequirements(s)); setRawSheets(p=>({...p,req:s})); }
    if (apiJson.cap?.length)  { const s={"07. SAP Tech Sprint Capacity Management":apiJson.cap}; setCap(parseCapacity(s)); setRawSheets(p=>({...p,cap:s})); }
    if (apiJson.test?.length) { const s={"110. Test Scenarios":apiJson.test}; setTest(parseTestScenarios(s)); setRawSheets(p=>({...p,test:s})); }
    if (apiJson.meta)         setSyncMeta({ lastSync: apiJson.meta.lastSync, source: "smartsheet" });
    console.log("[PMT] API data applied:", apiJson.meta);
  }, []);

  // ── On mount: restore persisted data ──────────────────────────────────────
  useEffect(() => {
    (async () => {
      // Try backend API first
      setIsLoading(true);
      try {
        const res = await fetch("/api/data");
        if (res.ok) {
          const apiJson = await res.json();
          applyApiData(apiJson);
          setFnames({ wp:"Smartsheet", raid:"Smartsheet", req:"Smartsheet", test:"Smartsheet", cap:"Smartsheet" });
          setStorageLoaded(true);
          setIsLoading(false);
          return;
        }
      } catch (e) {
        console.warn("[PMT] API unavailable, falling back to sessionStorage:", e.message);
      } finally {
        setIsLoading(false);
      }

      // Fall back to sessionStorage restore
      console.log("[PMT] Starting restore on mount...");
      const status = {};

      const wpRaw     = await restore(KEYS.wp);
      const raidRaw   = await restore(KEYS.raid);
      const reqRaw    = await restore(KEYS.req);
      const capRaw    = await restore(KEYS.cap);
      const testRaw   = await restore(KEYS.test);
      const fnamesRaw = await restore(KEYS.fnames);

      status.wp    = wpRaw    ? `✅ ${Object.values(wpRaw)[0]?.length || 0} rows`    : "❌ not found";
      status.raid  = raidRaw  ? `✅ ${Object.values(raidRaw)[0]?.length || 0} rows`  : "❌ not found";
      status.req   = reqRaw   ? `✅ ${Object.values(reqRaw)[0]?.length || 0} rows`   : "❌ not found";
      status.cap   = capRaw   ? `✅ ${Object.values(capRaw)[0]?.length || 0} rows`   : "❌ not found";
      status.test  = testRaw  ? `✅ ${Object.values(testRaw)[0]?.length || 0} rows`  : "❌ not found";

      if (wpRaw)   { setWp(parseWorkplan(wpRaw));      setRawSheets(p=>({...p,wp:wpRaw})); }
      if (raidRaw) { setRaid(parseRaid(raidRaw));       setRawSheets(p=>({...p,raid:raidRaw})); }
      if (reqRaw)  { setReq(parseRequirements(reqRaw)); setRawSheets(p=>({...p,req:reqRaw})); }
      if (capRaw)  { setCap(parseCapacity(capRaw));         setRawSheets(p=>({...p,cap:capRaw})); }
      if (testRaw) { setTest(parseTestScenarios(testRaw));  setRawSheets(p=>({...p,test:testRaw})); }
      if (fnamesRaw) setFnames(typeof fnamesRaw === "object" ? fnamesRaw : JSON.parse(fnamesRaw));

      setStorageStatus(status);
      setStorageLoaded(true);
      console.log("[PMT] Restore complete:", status);
    })();
  }, []);

  const openModal = (title, rows, columns) => { if (rows?.length) setModal({ title, rows, columns }); };

  const load = useCallback(async (type, file) => {
    console.log("[PMT] Loading file:", file.name, "type:", type);
    const sheets = await readXlsx(file);
    console.log("[PMT] Read sheets:", Object.keys(sheets));

    if (type === "wp")        setWp(parseWorkplan(sheets));
    else if (type === "raid") setRaid(parseRaid(sheets));
    else if (type === "req")  setReq(parseRequirements(sheets));
    else if (type === "cap")  setCap(parseCapacity(sheets));
    else if (type === "test") setTest(parseTestScenarios(sheets));

    setRawSheets(prev => ({ ...prev, [type]: sheets }));

    const keepCols = type === "wp"   ? WP_COLS_KEEP
                   : type === "req"  ? REQ_COLS_KEEP
                   : type === "raid" ? RAID_COLS_KEEP
                   : type === "test" ? TEST_COLS_KEEP
                   : null;
    await persist(KEYS[type], sheets, keepCols);

    // Use functional update so we never capture stale fnames
    setFnames(prev => {
      const newFnames = { ...prev, [type]: file.name };
      try { window.storage?.set(KEYS.fnames, JSON.stringify(newFnames)); } catch(e) {}
      try { sessionStorage.setItem(KEYS.fnames, JSON.stringify(newFnames)); } catch(e) {}
      return newFnames;
    });
    console.log("[PMT] Load complete for", type);
  }, []); // no deps needed — all setters are stable

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'Segoe UI', Arial, sans-serif", color: C.text }}>
      {/* Header */}
      <div style={{ background: "#000", padding: "0 24px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 50 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 26, height: 26, background: C.gold, borderRadius: 5, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13 }}>⚡</div>
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 15 }}>Performance Management for TAM Dashboard</span>
        </div>
        <span style={{ color: "rgba(255,255,255,0.55)", fontSize: 12 }}>{today.toLocaleDateString("en-US", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}</span>
      </div>

      {/* Upload Bar */}
      <div style={{ background: "#eaecf2", borderBottom: `1px solid ${C.border}`, padding: "10px 24px" }}>
        <div style={{ maxWidth: 1400, margin: "0 auto" }}>

          {/* Top row — status + actions */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 8, marginBottom: 8 }}>

            {/* Left — data status */}
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              {syncMeta ? (
                <>
                  <span style={{ fontSize: 10, fontWeight: 700, color: "#16a34a", textTransform: "uppercase", letterSpacing: "0.06em" }}>
                    ✓ Smartsheet Data Loaded
                  </span>
                  <span style={{ fontSize: 10, color: C.muted }}>
                    {syncMeta.source === "smartsheet" ? "via Smartsheet sync" : "via snapshot"} · {syncMeta.lastSync?.slice(0,10)}
                  </span>
                  {[["wp","Workplan",wp],["raid","RAID",raid],["req","Requirements",req],["test","Test Scenarios",test],["cap","Capacity",cap]].map(([key,label,state]) => (
                    <span key={key} style={{ display:"flex", alignItems:"center", gap:3, fontSize:10 }}>
                      <span style={{ width:6, height:6, borderRadius:"50%", background: state ? "#16a34a" : "#94a3b8", display:"inline-block" }} />
                      <span style={{ color: state ? C.text : C.muted }}>{label}</span>
                    </span>
                  ))}
                </>
              ) : (
                <>
                  <span style={{ fontSize: 11, color: C.gold, fontWeight: 600 }}>
                    ⚠ No data loaded — upload Smartsheet JSON or individual XLSX files below
                  </span>
                  {isLoading && <span style={{ fontSize: 11, color: C.accent }}>⏳ Loading from Smartsheet…</span>}
                </>
              )}
            </div>

            {/* Right — action buttons */}
            <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>

              {/* Refresh from Smartsheet (API) */}
              <button
                disabled={refreshing}
                onClick={async () => {
                  setRefreshing(true);
                  try {
                    const r = await fetch("/api/refresh", { method: "POST" });
                    if (!r.ok) { alert("Refresh failed: " + r.status); return; }
                    const dataRes = await fetch("/api/data");
                    if (dataRes.ok) {
                      const apiJson = await dataRes.json();
                      applyApiData(apiJson);
                      setFnames({ wp:"Smartsheet", raid:"Smartsheet", req:"Smartsheet", test:"Smartsheet", cap:"Smartsheet" });
                      alert("✅ Refreshed from Smartsheet!");
                    }
                  } catch(e) { alert("Refresh error: " + e.message); }
                  finally { setRefreshing(false); }
                }}
                style={{ padding: "7px 14px", background: refreshing ? "#e2e8f0" : "#16a34a",
                  border: "none", borderRadius: 6, cursor: refreshing ? "wait" : "pointer",
                  color: refreshing ? C.muted : "#fff", fontSize: 11, fontWeight: 700 }}>
                {refreshing ? "⏳ Refreshing…" : "🔄 Refresh from Smartsheet"}
              </button>
            </div>
          </div>

        </div>
      </div>

      {/* Tab Bar */}
      <div style={{ background: "#595959", borderBottom: `1px solid #444` }}>
        <div style={{ maxWidth: 1400, margin: "0 auto", padding: "0 24px", display: "flex" }}>
          {TABS.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{
              padding: "12px 16px", border: "none", background: "transparent", cursor: "pointer",
              color: tab === t.id ? "#fff" : "rgba(255,255,255,0.7)",
              borderBottom: `3px solid ${tab === t.id ? "#fff" : "transparent"}`,
              fontWeight: tab === t.id ? 700 : 500, fontSize: 13, transition: "all .12s",
            }}>{t.label}</button>
          ))}
        </div>
      </div>

      {/* Content */}
      <div style={{ maxWidth: "100%", margin: "0 auto", padding: "20px 16px" }}>
        {tab === "executive"    && <ExecutiveSummaryTab wp={wp} raid={raid} req={req} cap={cap} openModal={openModal} />}
        {tab === "workplan"     && <WorkplanTab wp={wp} raid={raid} openModal={openModal} />}
        {tab === "raid"         && <RaidAnalysisTab raid={raid} />}
        {tab === "cr"           && <ChangeRequestTab raid={raid} />}
        {tab === "backlog"      && <BacklogTab raid={raid} />}
        {tab === "overview"     && <OverviewTab wp={wp} raid={raid} req={req} cap={cap} openModal={openModal} />}
        {tab === "scorecard"    && <ScorecardTab wp={wp} raid={raid} req={req} openModal={openModal} />}
        {tab === "testing"      && <TestScenariosTab data={test} wp={wp} />}
      </div>

      {modal && <Modal title={modal.title} rows={modal.rows} columns={modal.columns} onClose={() => setModal(null)} />}

      {/* Snapshot text modal — copy this and save as .json */}
      {snapshotText && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.6)", zIndex:2000, display:"flex", alignItems:"center", justifyContent:"center" }}
          onClick={() => setSnapshotText(null)}>
          <div style={{ background:"#fff", borderRadius:10, width:"80%", maxWidth:800, maxHeight:"80vh", display:"flex", flexDirection:"column", boxShadow:"0 24px 60px rgba(0,0,0,0.4)" }}
            onClick={e => e.stopPropagation()}>
            <div style={{ background:C.headerBg, padding:"12px 20px", borderRadius:"10px 10px 0 0", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
              <div style={{ color:"#fff", fontWeight:700, fontSize:13 }}>📋 Snapshot Data — Select All &amp; Copy, then save as <code style={{background:"rgba(255,255,255,0.2)",padding:"1px 6px",borderRadius:3}}>pmt_snapshot.json</code></div>
              <button onClick={() => setSnapshotText(null)} style={{ background:"rgba(255,255,255,0.15)", border:"none", color:"#fff", borderRadius:5, padding:"4px 12px", cursor:"pointer", fontSize:13, fontWeight:600 }}>✕</button>
            </div>
            <div style={{ padding:12, background:"#f8fafc", flex:1, display:"flex", flexDirection:"column", gap:8 }}>
              <div style={{ display:"flex", gap:8 }}>
                <button onClick={() => { const ta = document.getElementById("snapshot-ta"); ta.select(); document.execCommand("copy"); alert("✅ Copied! Now paste into a text editor and save as pmt_snapshot.json"); }}
                  style={{ padding:"6px 16px", background:C.navyLight, color:"#fff", border:"none", borderRadius:6, cursor:"pointer", fontWeight:700, fontSize:12 }}>
                  📋 Copy to Clipboard
                </button>
                <span style={{ color:C.muted, fontSize:11, alignSelf:"center" }}>Then paste into Notepad/TextEdit and save as <b>pmt_snapshot.json</b></span>
              </div>
              <textarea id="snapshot-ta" readOnly value={snapshotText}
                style={{ flex:1, minHeight:300, fontFamily:"monospace", fontSize:10, padding:10, border:`1px solid ${C.border}`, borderRadius:6, resize:"none", background:"#fff" }}
                onClick={e => e.target.select()} />
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── EXECUTIVE SUMMARY TAB ───────────────────────────────────────────────────
const PRIORITY_COLORS = { "1 - Critical": "#7b0d0d", "1-Critical": "#7b0d0d", "Critical": "#7b0d0d", "2 - High": "#c0392b", "2-High": "#c0392b", "High": "#c0392b", "3 - Medium": "#f5a623", "3-Medium": "#f5a623", "Medium": "#f5a623", "4 - Low": "#1a73e8", "4-Low": "#1a73e8", "Low": "#1a73e8" };
const getPriorityColor = (p) => { const k = Object.keys(PRIORITY_COLORS).find(k => String(p||"").toLowerCase().includes(k.toLowerCase().replace(/[^a-z0-9]/g,""))); return k ? PRIORITY_COLORS[k] : "#888"; };

function ExecutiveSummaryTab({ wp, raid, req, cap, openModal }) {
  const [wpModal, setWpModal] = useState(null);
  const [raidModal, setRaidModal] = useState(null);
  const [storyModal, setStoryModal] = useState(null);

  const anyData = wp || raid || req || cap;
  if (!anyData) return <Empty label="Upload files above to populate the Executive Summary." />;

  // ── 1. KPIs ────────────────────────────────────────────────────────────────
  // Use leaf rows only (Children == 0) so counts match what the drill-down shows
  const isWpLeaf = r => { const c = r["Children"]; if (c === null || c === undefined || c === "" || c === "0") return true; const n = Number(c); return isNaN(n) || n === 0; };
  const wpLeaves   = wp ? wp.allRows.filter(isWpLeaf) : [];
  const getWpS     = r => String(r["Default Status"] || r["Status"] || "").toLowerCase();
  const wpTotal    = wpLeaves.length;
  const wpOffTrack = wpLeaves.filter(r => getWpS(r).includes("off track")).length;
  const wpOnTrack  = wpLeaves.filter(r => getWpS(r).includes("on track")).length;
  const wpComplete = wpLeaves.filter(r => getWpS(r).includes("complete")).length;
  const wpHealth   = wpTotal > 0 ? Math.round(((wpOnTrack + wpComplete) / wpTotal) * 100) : null;

  const raidOpen    = raid ? raid.open.length : null;
  const raidDelayed = raid ? raid.delayed.length : null;

  const reqDone  = req ? req.done.length : null;
  const reqTotal = req ? req.total : null;
  const reqPct   = reqTotal > 0 ? Math.round((reqDone / reqTotal) * 100) : null;

  const latestSprint = cap ? cap.sprintChart[cap.sprintChart.length - 1] : null;
  const capSurplus   = latestSprint ? latestSprint.diff : null;

  // ── 2. Workstream status ───────────────────────────────────────────────────
  const isWpLeafRow = r => { const c = r["Children"]; if (c === null || c === undefined || c === "" || c === "0") return true; const n = Number(c); return isNaN(n) || n === 0; };

  const workstreams = wp ? Object.entries(wp.componentMap).map(([name, d]) => {
    const total = d.total || 1;
    const pctVal = Math.round(((d.complete + d.onTrack * 0.5) / total) * 100);
    const health = d.offTrack > 0 ? "At Risk" : d.onTrack > 0 ? "On Track" : d.complete === d.total ? "Complete" : "Not Started";
    const healthColor = health === "At Risk" ? C.delayed : health === "On Track" ? C.onTrack : health === "Complete" ? C.complete : C.muted;
    // Count only leaf rows for the badge — matches what drill-down shows
    const leafOffTrack = d.rows.filter(r => isWpLeafRow(r) && getWpS(r).includes("off track")).length;
    return { name, d: { ...d, offTrack: leafOffTrack }, pctVal, health, healthColor };
  }) : [];

  // ── 3. Component RAG (Lvl 3 from scorecard workplan scope) ────────────────
  const compRows = wp ? wp.allRows.filter(r =>
    String(r["Activity Grp - Lvl 1"] || "").trim() === "Technology - SAP Configuration & Build" &&
    String(r["Activity Grp - Lvl 2"] || "").trim() === "Component Build"
  ) : [];
  const compNames = Array.from(new Set(compRows.map(r => String(r["Activity Grp - Lvl 3"] || "").trim()).filter(Boolean))).sort();

  const getCompStatus = (compName) => {
    const rows = compRows.filter(r => String(r["Activity Grp - Lvl 3"] || "").trim() === compName);
    const leaves = rows.filter(r => { const c = r["Children"]; return !c || Number(c) === 0; });
    const getS = r => String(r["Default Status"] || r["Status"] || "").toLowerCase();
    // Handle Excel decimal (0.92), whole number (92), string "0.92", string "92%", string "92"
    const normPct = v => {
      if (v === "" || v == null) return null;
      const s = String(v).replace("%", "").trim();
      if (s === "" || isNaN(Number(s))) return null;
      const n = Number(s);
      return n <= 1 ? Math.round(n * 100) : Math.round(n);
    };
    const isDone = r => getS(r).includes("complete") || normPct(r["% Complete"] ?? r["% complete"]) === 100;
    const hasOffTrack = leaves.some(r => getS(r).includes("off track"));
    const hasOnTrack  = leaves.some(r => getS(r).includes("on track"));
    const allComplete = leaves.length > 0 && leaves.every(r => isDone(r));
    const delayedCount = leaves.filter(r => getS(r).includes("off track")).length;
    const lvl3Header = rows.find(r => Number(r["Lvl"] ?? 0) === 3);
    const headerPct = lvl3Header ? normPct(lvl3Header["% Complete"] ?? lvl3Header["% complete"]) : null;
    const status = hasOffTrack ? "Off Track" : (allComplete || headerPct === 100) ? "Complete" : hasOnTrack ? "On Track" : "Not Started";
    const pctValues = leaves.map(r => normPct(r["% Complete"] ?? r["% complete"])).filter(v => v != null);
    const pct2 = headerPct != null ? headerPct : (pctValues.length ? Math.round(pctValues.reduce((a, b) => a + b, 0) / pctValues.length) : null);
    return { status, pct: pct2, rows, delayedCount };
  };

  // ── 4. Top RAIDs ──────────────────────────────────────────────────────────
  const PRIORITY_RANK = { "1": 4, "critical": 4, "2": 3, "high": 3, "3": 2, "medium": 2, "4": 1, "low": 1 };
  const getPRank = r => {
    const p = String(r[raid?.keys?.priority] || "").toLowerCase();
    for (const [k, v] of Object.entries(PRIORITY_RANK)) { if (p.includes(k)) return v; }
    return 0;
  };
  const topRaids = raid ? [...raid.open]
    .filter(r => { const t = String(r[raid.keys.type] || "").toLowerCase(); return t.includes("risk") || t.includes("issue"); })
    .sort((a, b) => {
      const pr = getPRank(b) - getPRank(a);
      if (pr !== 0) return pr;
      const da = daysUntil(a[raid.keys.date]), db = daysUntil(b[raid.keys.date]);
      return (da ?? 999) - (db ?? 999);
    })
    .slice(0, 5) : [];

  // ── 5. Key milestones ─────────────────────────────────────────────────────
  const milestones = wp ? wp.allRows
    .filter(r => {
      const c = r["Children"]; const isLeaf = !c || Number(c) === 0;
      const d = daysUntil(r["Finish"] || r["End Date"]);
      const s = String(r["Default Status"] || r["Status"] || "").toLowerCase();
      return isLeaf && d != null && d <= 30 && !s.includes("complete");
    })
    .sort((a, b) => daysUntil(a["Finish"] || a["End Date"]) - daysUntil(b["Finish"] || b["End Date"]))
    .slice(0, 6) : [];

  // ── Helpers ────────────────────────────────────────────────────────────────
  const StatusPill = ({ status }) => {
    const sl = String(status || "").toLowerCase();
    const bg    = sl.includes("off track") || sl.includes("at risk") ? "#fee2e2" : sl.includes("on track") ? "#fef3c7" : sl.includes("complete") ? "#dbeafe" : "#f1f5f9";
    const color = sl.includes("off track") || sl.includes("at risk") ? "#b91c1c" : sl.includes("on track") ? "#b45309" : sl.includes("complete") ? "#1d4ed8" : "#64748b";
    const border= sl.includes("off track") || sl.includes("at risk") ? "#fca5a5" : sl.includes("on track") ? "#fcd34d" : sl.includes("complete") ? "#93c5fd" : "#cbd5e1";
    return <span style={{ background: bg, color, border: `1px solid ${border}`, borderRadius: 4, padding: "2px 8px", fontSize: 10, fontWeight: 700, whiteSpace: "nowrap" }}>{status}</span>;
  };

  const PctBar = ({ pct, width = 80 }) => {
    if (pct == null) return <span style={{ color: C.muted, fontSize: 11 }}>—</span>;
    const bg = pct >= 75 ? C.green : pct >= 40 ? C.gold : C.delayed;
    return (
      <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
        <div style={{ width, background: "#e2e8f0", borderRadius: 4, height: 6, overflow: "hidden" }}>
          <div style={{ width: `${pct}%`, height: "100%", background: bg, borderRadius: 4 }} />
        </div>
        <span style={{ fontSize: 11, fontWeight: 700, color: C.text, minWidth: 30 }}>{pct}%</span>
      </div>
    );
  };

  const Rag = ({ status }) => {
    const sl = String(status || "").toLowerCase();
    const bg = sl.includes("off track") || sl.includes("at risk") ? "#ef4444" : sl.includes("on track") ? "#f59e0b" : sl.includes("complete") ? "#22c55e" : "#94a3b8";
    return <span style={{ width: 9, height: 9, borderRadius: "50%", background: bg, display: "inline-block", flexShrink: 0 }} />;
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>

      {/* ── 1. KPIs + CR Summary ──────────────────────────────────────────── */}
      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(5, minmax(0,1fr)) 4px repeat(4, minmax(0,1fr))", gap: 8, alignItems: "stretch" }}>

          {/* ── 5 KPI tiles ── */}
          {[
            ["Workplan Health",  wpHealth != null ? `${wpHealth}%` : "—",
              wpHealth == null ? C.muted : wpHealth >= 75 ? C.green : wpHealth >= 50 ? C.gold : C.delayed,
              `${wpOnTrack} on track · ${wpOffTrack} at risk`,
              wp ? () => setWpModal({ title: "Workplan Health", rows: wp.allRows, initialFilter: "Off Track" }) : null],

            ["Open RAIDs",       raidOpen ?? "—",
              raidDelayed > 0 ? C.delayed : C.navyLight,
              `${raidDelayed ?? 0} delayed`,
              raid ? () => setRaidModal({ title: "Open RAIDs", rows: raid.open, initialStatusFilter: raidDelayed > 0 ? "Delayed" : "All" }) : null],

            ["Risks & Issues",   raid ? (raid.openRisks.length + raid.openIssues.length) : "—",
              raid && (raid.openRisks.length + raid.openIssues.length) > 0 ? C.delayed : C.green,
              raid ? `${raid.openRisks.length} risks · ${raid.openIssues.length} issues` : "—",
              raid ? () => setRaidModal({ title: "Open Risks & Issues", rows: [...raid.openRisks, ...raid.openIssues], initialTypeFilter: "All" }) : null],

            ["Stories Complete", reqPct != null ? `${reqPct}%` : "—",
              reqPct == null ? C.muted : reqPct >= 75 ? C.green : reqPct >= 40 ? C.gold : C.delayed,
              `${reqDone ?? 0} of ${reqTotal ?? 0} stories`,
              req?.done?.length ? () => setStoryModal({ title: "Stories Complete", rows: req.done }) : null],

            ["Stories Blocked",  req ? req.blocked.length : "—",
              req && req.blocked.length > 0 ? C.delayed : C.green,
              req && req.blocked.length > 0 ? "need attention" : "none blocked",
              req?.blocked?.length ? () => setStoryModal({ title: "Stories Blocked", rows: req.blocked }) : null],
          ].map(([label, value, color, sub, onClick]) => (
            <div key={label} onClick={onClick}
              onMouseEnter={e => { if (onClick) e.currentTarget.style.boxShadow = "0 3px 10px rgba(0,0,0,0.12)"; }}
              onMouseLeave={e => { e.currentTarget.style.boxShadow = "0 1px 3px rgba(0,0,0,0.06)"; }}
              style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 7,
                padding: "10px 12px", borderTop: `3px solid ${color}`,
                boxShadow: "0 1px 3px rgba(0,0,0,0.06)", cursor: onClick ? "pointer" : "default" }}>
              <div style={{ color: C.muted, fontSize: 9, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 2 }}>{label}</div>
              <div style={{ color, fontSize: 22, fontWeight: 800, lineHeight: 1 }}>{value}</div>
              {sub && <div style={{ color: C.muted, fontSize: 9, marginTop: 2 }}>{sub}</div>}
              {onClick && <div style={{ color: C.accent, fontSize: 9, marginTop: 2 }}>Details →</div>}
            </div>
          ))}

          {/* ── Divider ── */}
          <div style={{ background: C.border, borderRadius: 2, alignSelf: "stretch" }} />

          {/* ── 4 CR tiles ── */}
          {[
            { label: "CR Approved",  rows: raid?.cr?.approved ?? [], hours: raid?.cr?.approvedHours ?? 0, color: C.green },
            { label: "CR Rejected",  rows: raid?.cr?.rejected ?? [], hours: raid?.cr?.rejectedHours ?? 0, color: C.delayed },
            { label: "CR Deferred",  rows: raid?.cr?.deferred ?? [], hours: raid?.cr?.deferredHours ?? 0, color: C.gold },
            { label: "CR Pending",   rows: raid?.cr?.pending  ?? [], hours: raid?.cr?.pendingHours  ?? 0, color: C.complete },
          ].map(({ label, rows, hours, color }) => (
            <div key={label} onClick={() => rows.length && setRaidModal({ title: label, rows })}
              onMouseEnter={e => { if (rows.length) e.currentTarget.style.boxShadow = "0 3px 10px rgba(0,0,0,0.12)"; }}
              onMouseLeave={e => { e.currentTarget.style.boxShadow = "0 1px 3px rgba(0,0,0,0.06)"; }}
              style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 7,
                padding: "10px 12px", borderTop: `3px solid ${color}`,
                boxShadow: "0 1px 3px rgba(0,0,0,0.06)", cursor: rows.length ? "pointer" : "default" }}>
              <div style={{ color: C.muted, fontSize: 9, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 2 }}>{label}</div>
              <div style={{ color, fontSize: 22, fontWeight: 800, lineHeight: 1 }}>{rows.length > 0 ? rows.length : "—"}</div>
              <div style={{ color: C.muted, fontSize: 9, marginTop: 2 }}>{hours > 0 ? `${hours.toLocaleString()} hrs` : "—"}</div>
              {rows.length > 0 && <div style={{ color: C.accent, fontSize: 9, marginTop: 2 }}>Details →</div>}
            </div>
          ))}
        </div>

        {/* Colour legend */}
        <div style={{ display: "flex", gap: 14, alignItems: "center" }}>
          <span style={{ fontSize: 10, color: C.muted, fontWeight: 600 }}>KPI colours:</span>
          {[[C.green, "Good"], [C.gold, "Watch"], [C.delayed, "At Risk"], [C.navyLight, "Informational"]].map(([col, lbl]) => (
            <span key={lbl} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 10, color: C.muted }}>
              <span style={{ width: 10, height: 10, borderRadius: 2, background: col, display: "inline-block" }} />{lbl}
            </span>
          ))}
        </div>
      </div>




      {/* Modals */}
      {wpModal    && <WorkplanDrillModal title={wpModal.title} rows={wpModal.rows} initialFilter={wpModal.initialFilter} onClose={() => setWpModal(null)} />}
      {raidModal  && <RaidDrillModal title={raidModal.title} rows={raidModal.rows} raidKeys={raid?.keys} initialStatusFilter={raidModal.initialStatusFilter} initialTypeFilter={raidModal.initialTypeFilter} onClose={() => setRaidModal(null)} />}
      {storyModal && <StoryDrillModal title={storyModal.title} rows={storyModal.rows} reqKeys={req?.keys} onClose={() => setStoryModal(null)} />}
    </div>
  );
}

// ─── RAID KPI DRILL-DOWN MODAL ───────────────────────────────────────────────
function RaidKpiModal({ title, rows, K, teamKey, allTeams, allTypes, allComps, statusCol, hideType, hideStatus, colConfig, setColConfig, onClose }) {
  const [teamFilter,   setTeamFilter]   = useState("All");
  const [statusFilter, setStatusFilter] = useState("All");
  const [typeFilter,   setTypeFilter]   = useState("All");
  const [compFilter,   setCompFilter]   = useState("All");
  const [showCols,     setShowCols]     = useState(false);

  // Independent cross-filtering — each filter's counts reflect all OTHER active filters
  const matchStatus = r => statusFilter === "All" || (statusFilter === "Delayed" ? String(r[K.status]||"").toLowerCase().includes("delay") : !String(r[K.status]||"").toLowerCase().includes("delay") && !String(r[K.status]||"").toLowerCase().includes("complete"));
  const matchType   = r => typeFilter   === "All" || String(r[K.type]||"").trim()      === typeFilter;
  const matchComp   = r => compFilter   === "All" || String(r[K.component]||"").trim() === compFilter;
  const matchTeam   = r => teamFilter   === "All" || String(r[teamKey]||"").trim()     === teamFilter;

  const filtered = rows.filter(r => matchTeam(r) && matchStatus(r) && matchType(r) && matchComp(r));

  // Each filter shows counts based on all OTHER active filters (true independent cross-filtering)
  const teamCounts   = allTeams.map(t => ({ val:t, count: rows.filter(r => String(r[teamKey]||"").trim()===t && matchStatus(r) && matchType(r) && matchComp(r)).length }));
  const statusCounts = {
    all:     rows.filter(r => matchTeam(r) && matchType(r) && matchComp(r)).length,
    delayed: rows.filter(r => matchTeam(r) && matchType(r) && matchComp(r) && String(r[K.status]||"").toLowerCase().includes("delay")).length,
    onTrack: rows.filter(r => matchTeam(r) && matchType(r) && matchComp(r) && !String(r[K.status]||"").toLowerCase().includes("delay") && !String(r[K.status]||"").toLowerCase().includes("complete")).length,
  };
  const typesWithCount = allTypes.map(t => ({ val:t, count: rows.filter(r => matchTeam(r) && matchStatus(r) && matchComp(r) && String(r[K.type]||"").trim()===t).length })).filter(t=>t.count>0);
  const compsWithCount = allComps.map(c => ({ val:c, count: rows.filter(r => matchTeam(r) && matchStatus(r) && matchType(r) && String(r[K.component]||"").trim()===c).length })).filter(c=>c.count>0);

  // pill(val, isActive, count, onClick, color) — highlights if count>0, dims if 0
  const filterBtn = (val, isActive, onClick, count, color) => {
    const hasItems = count > 0;
    const borderCol = isActive ? (color||C.navyLight) : hasItems ? (color ? color+"80" : "#b0bbc8") : C.border;
    const bg = isActive ? (color||C.navyLight) : C.white;
    const textCol = isActive ? "#fff" : hasItems ? C.text : C.muted;
    return (
      <button key={val} onClick={onClick} disabled={!hasItems && val!=="All"}
        style={{ display:"flex", alignItems:"center", gap:4, padding:"3px 10px", borderRadius:20,
          border:`2px solid ${borderCol}`, background:bg, color:textCol,
          cursor: hasItems||val==="All" ? "pointer" : "default",
          fontSize:10, fontWeight:700, transition:"all .12s",
          opacity: !hasItems && val!=="All" ? 0.4 : 1 }}>
        {val}
        <span style={{ background: isActive?"rgba(255,255,255,0.25)":"#f1f5f9", color: isActive?"#fff":C.text,
          borderRadius:10, padding:"1px 6px", fontSize:10, fontWeight:800, minWidth:16, textAlign:"center" }}>{count}</span>
      </button>
    );
  };

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:1000, display:"flex", alignItems:"center", justifyContent:"center" }} onClick={onClose}>
      <div style={{ background:C.white, borderRadius:10, width:"98%", maxWidth:1300, maxHeight:"92vh", display:"flex", flexDirection:"column", boxShadow:"0 24px 60px rgba(0,0,0,0.3)" }} onClick={e=>e.stopPropagation()}>

        {/* Header */}
        <div style={{ background:C.headerBg, padding:"12px 20px", borderRadius:"10px 10px 0 0", display:"flex", justifyContent:"space-between", alignItems:"center", flexShrink:0 }}>
          <div style={{ color:"#fff", fontWeight:700, fontSize:13 }}>{title} <span style={{ opacity:.6, fontWeight:400 }}>({filtered.length} of {rows.length})</span></div>
          <button onClick={onClose} style={{ background:"rgba(255,255,255,0.15)", border:"none", color:"#fff", borderRadius:5, padding:"5px 14px", cursor:"pointer", fontSize:13, fontWeight:600 }}>✕</button>
        </div>

        {/* Filters */}
        <div style={{ padding:"10px 16px", borderBottom:`1px solid ${C.border}`, background:"#f8fafc", display:"flex", flexDirection:"column", gap:8, flexShrink:0 }}>

          {/* Team filter — always show */}
          {allTeams.length > 0 && (
            <div style={{ display:"flex", gap:4, alignItems:"center", flexWrap:"wrap" }}>
              <span style={{ fontSize:10, color:C.muted, fontWeight:600, marginRight:2 }}>Team</span>
              {filterBtn("All", teamFilter==="All", () => setTeamFilter("All"), rows.filter(r => matchStatus(r) && matchType(r) && matchComp(r)).length)}
              {teamCounts.map(({val,count}) => filterBtn(val, teamFilter===val, () => setTeamFilter(teamFilter===val?"All":val), count))}
            </div>
          )}
          {/* Fallback if teamKey not resolving — show raw team values */}
          {allTeams.length === 0 && (
            <div style={{ fontSize:10, color:C.muted, fontStyle:"italic" }}>
              Team filter unavailable — team column not detected (key: {teamKey||"none"})
            </div>
          )}

          {/* Status + Type row — hidden contextually */}
          {(!hideStatus || !hideType) && (
            <div style={{ display:"flex", gap:12, alignItems:"center", flexWrap:"wrap" }}>
              {!hideStatus && (
                <div style={{ display:"flex", gap:4, alignItems:"center" }}>
                  <span style={{ fontSize:10, color:C.muted, fontWeight:600, marginRight:2 }}>Status</span>
                  {filterBtn("All",      statusFilter==="All",      () => setStatusFilter("All"),      statusCounts.all,     null)}
                  {filterBtn("Delayed",  statusFilter==="Delayed",  () => setStatusFilter(statusFilter==="Delayed"?"All":"Delayed"),   statusCounts.delayed, C.delayed)}
                  {filterBtn("On Track", statusFilter==="On Track", () => setStatusFilter(statusFilter==="On Track"?"All":"On Track"), statusCounts.onTrack, C.onTrack)}
                </div>
              )}
              {!hideStatus && !hideType && <div style={{ width:1, height:18, background:C.border }} />}
              {!hideType && (
                <div style={{ display:"flex", gap:4, alignItems:"center", flexWrap:"wrap" }}>
                  <span style={{ fontSize:10, color:C.muted, fontWeight:600, marginRight:2 }}>Type</span>
                  {filterBtn("All", typeFilter==="All", () => setTypeFilter("All"), rows.filter(r => matchTeam(r) && matchStatus(r) && matchComp(r)).length)}
                  {typesWithCount.map(({val,count}) => filterBtn(val, typeFilter===val, () => setTypeFilter(typeFilter===val?"All":val), count))}
                </div>
              )}
              {/* Columns button */}
              <button onClick={() => setShowCols(p=>!p)}
                style={{ marginLeft:"auto", padding:"4px 12px", borderRadius:5, border:`1px solid ${showCols?C.navyLight:C.border}`,
                  background: showCols?C.navyLight:C.white, color: showCols?"#fff":C.muted,
                  cursor:"pointer", fontSize:10, fontWeight:600 }}>⚙ Columns</button>
            </div>
          )}
          {/* Columns button standalone when both filters hidden */}
          {hideStatus && hideType && (
            <div style={{ display:"flex", justifyContent:"flex-end" }}>
              <button onClick={() => setShowCols(p=>!p)}
                style={{ padding:"4px 12px", borderRadius:5, border:`1px solid ${showCols?C.navyLight:C.border}`,
                  background: showCols?C.navyLight:C.white, color: showCols?"#fff":C.muted,
                  cursor:"pointer", fontSize:10, fontWeight:600 }}>⚙ Columns</button>
            </div>
          )}

          {/* Component filter */}
          {compsWithCount.length > 0 && (
            <div style={{ display:"flex", gap:4, alignItems:"center", flexWrap:"wrap" }}>
              <span style={{ fontSize:10, color:C.muted, fontWeight:600, marginRight:2 }}>Component</span>
              {filterBtn("All", compFilter==="All", () => setCompFilter("All"), rows.filter(r => matchTeam(r) && matchStatus(r) && matchType(r)).length)}
              {compsWithCount.map(({val,count}) => filterBtn(val, compFilter===val, () => setCompFilter(compFilter===val?"All":val), count))}
            </div>
          )}

          {/* Column show/hide */}
          {showCols && (
            <div style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:8, padding:"10px 12px" }}>
              <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                {Object.entries(colConfig).map(([key, col]) => (
                  <label key={key} style={{ display:"flex", alignItems:"center", gap:4, background:"#f8fafc",
                    border:`1px solid ${col.visible?C.navyLight:C.border}`, borderRadius:5, padding:"4px 9px", cursor:"pointer" }}>
                    <input type="checkbox" checked={col.visible}
                      onChange={e => setColConfig(p => ({...p, [key]:{...p[key], visible:e.target.checked}}))}
                      style={{ cursor:"pointer", width:12, height:12 }} />
                    <span style={{ fontSize:10, color:col.visible?C.navyLight:C.muted, fontWeight:col.visible?700:400 }}>{col.label}</span>
                  </label>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* Table */}
        <div style={{ overflowY:"auto", overflowX:"auto", flex:1 }}>
          <table style={{ borderCollapse:"collapse", fontSize:11, tableLayout:"fixed",
            width: Object.values(colConfig).filter(c=>c.visible).reduce((s,c)=>s+c.width,0)+"px", minWidth:"100%" }}>
            <thead style={{ position:"sticky", top:0, zIndex:2 }}>
              <tr style={{ background:"#162f50" }}>
                {[["raidId","RAID ID"],["status","Status"],["type","Type"],["component","Component"],
                  ["experience","Experience"],["topic","Topic"],["desc","Description"],
                  ["comment","Comments / Resolution"],["owner","Owner"],["team","Primary Team (Owner)"],
                  ["critPath","Critical Path"],["dueDate","Due Date"]
                ].filter(([key]) => colConfig[key]?.visible).map(([key,label],idx,arr) => (
                  <th key={key} style={{ padding:"8px 10px", textAlign:"left", color:"#fff", fontWeight:700, fontSize:10,
                    width:colConfig[key].width, position:"relative",
                    borderRight: idx<arr.length-1?"1px solid rgba(255,255,255,0.1)":"none" }}>
                    {label}
                    <div onMouseDown={e => {
                      e.preventDefault();
                      const startX=e.clientX, startW=colConfig[key].width;
                      const onMove=mv=>setColConfig(p=>({...p,[key]:{...p[key],width:Math.max(50,startW+mv.clientX-startX)}}));
                      const onUp=()=>{ window.removeEventListener("mousemove",onMove); window.removeEventListener("mouseup",onUp); };
                      window.addEventListener("mousemove",onMove); window.addEventListener("mouseup",onUp);
                    }}
                    style={{ position:"absolute",right:0,top:0,bottom:0,width:6,cursor:"col-resize",background:"transparent",zIndex:10 }}
                    onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.3)"}
                    onMouseLeave={e=>e.currentTarget.style.background="transparent"} />
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={Object.values(colConfig).filter(c=>c.visible).length} style={{ padding:"24px", textAlign:"center", color:C.muted }}>No items match filters</td></tr>
              ) : filtered.map((r,i) => {
                const status = String(r[K.status]||"").trim();
                const sCol = statusCol(status);
                const due = daysUntil(r[K.date]);
                const dueStr = fmtDate(r[K.date]);
                const dueCol = due!=null&&due<=7?C.delayed:due!=null&&due<=14?C.gold:C.muted;
                return (
                  <tr key={i} style={{ background:i%2===0?C.white:"#f7f9fc", borderBottom:`1px solid ${C.border}`, verticalAlign:"top" }}>
                    {colConfig.raidId.visible    && <td style={{ padding:"8px 10px", fontWeight:700, color:C.navyLight, wordBreak:"break-word", width:colConfig.raidId.width }}>{String(r[K.id]||"—")}</td>}
                    {colConfig.status.visible    && <td style={{ padding:"8px 10px", width:colConfig.status.width }}><span style={{ background:sCol+"20", color:sCol, border:`1px solid ${sCol}40`, borderRadius:4, padding:"2px 6px", fontSize:10, fontWeight:700, whiteSpace:"nowrap" }}>{status||"—"}</span></td>}
                    {colConfig.type.visible      && <td style={{ padding:"8px 10px", color:C.text, wordBreak:"break-word", width:colConfig.type.width }}>{String(r[K.type]||"—")}</td>}
                    {colConfig.component.visible && <td style={{ padding:"8px 10px", color:C.text, wordBreak:"break-word", width:colConfig.component.width }}>{String(r[K.component]||"—")}</td>}
                    {colConfig.experience.visible&& <td style={{ padding:"8px 10px", color:C.muted, wordBreak:"break-word", width:colConfig.experience.width }}>{String(r[K.experience]||"—")}</td>}
                    {colConfig.topic.visible     && <td style={{ padding:"8px 10px", color:C.muted, wordBreak:"break-word", width:colConfig.topic.width }}>{String(r[K.topic]||"—")}</td>}
                    {colConfig.desc.visible      && <td style={{ padding:"8px 10px", color:C.text, wordBreak:"break-word", lineHeight:1.5, width:colConfig.desc.width }}>{String(r[K.desc]||"—")}</td>}
                    {colConfig.comment.visible   && <td style={{ padding:"8px 10px", color:C.muted, wordBreak:"break-word", lineHeight:1.5, width:colConfig.comment.width }}>{String(r[K.comment]||"—")}</td>}
                    {colConfig.owner?.visible     && <td style={{ padding:"8px 10px", color:C.text, wordBreak:"break-word", width:colConfig.owner.width }}>{String(r[K.owner]||"—")}</td>}
                    {colConfig.team?.visible      && <td style={{ padding:"8px 10px", color:C.text, wordBreak:"break-word", width:colConfig.team?.width||140 }}>{String(r[teamKey]||"—")}</td>}
                    {colConfig.critPath?.visible  && <td style={{ padding:"8px 10px", width:colConfig.critPath.width }}>{(() => { const v=String(r[K.critPath]||"").trim(); if(!v||v==="—") return <span style={{color:C.muted}}>—</span>; const hi=v.toLowerCase()!=="no"&&v.toLowerCase()!=="n/a"; return <span style={{background:hi?"#fee2e2":"#f1f5f9",color:hi?C.delayed:C.muted,borderRadius:3,padding:"2px 6px",fontSize:10,fontWeight:600}}>{v}</span>; })()}</td>}
                    {colConfig.dueDate?.visible   && <td style={{ padding:"8px 10px", color:dueCol, fontWeight:600, whiteSpace:"nowrap", width:colConfig.dueDate.width }}>{dueStr}</td>}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── CHANGE REQUEST TAB ──────────────────────────────────────────────────────
function ChangeRequestTab({ raid }) {
  const [modal, setModal] = useState(null);

  if (!raid) return <Empty label="Upload RAID Log file above to view this tab." />;
  if (!raid.cr || raid.cr.all.length === 0) return (
    <Card>
      <div style={{ textAlign:"center", padding:"32px 0", color:C.muted, fontSize:13 }}>
        No Change Requests detected in RAID log.<br/>
        <span style={{ fontSize:11 }}>Ensure "Change Request Analysis" column is populated.</span>
      </div>
    </Card>
  );

  const K = raid.keys;
  const cr = raid.cr;
  const teamKey = K.team || "Primary Team (Owner)";

  const BUCKETS = [
    { lbl:"Approved",       rows:cr.approved, hours:cr.approvedHours, textCol:"#166534", bg:"#dcfce7", borderC:"#86efac" },
    { lbl:"Pending Review", rows:cr.pending,  hours:cr.pendingHours,  textCol:"#1d4ed8", bg:"#dbeafe", borderC:"#93c5fd" },
    { lbl:"Rejected",       rows:cr.rejected, hours:cr.rejectedHours, textCol:"#b91c1c", bg:"#fee2e2", borderC:"#fca5a5" },
    { lbl:"Deferred",       rows:cr.deferred, hours:cr.deferredHours, textCol:"#92400e", bg:"#fef3c7", borderC:"#fcd34d" },
  ];

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>

      {/* KPI tiles */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:10 }}>
        {BUCKETS.map(({ lbl, rows, hours, textCol, bg, borderC }) => (
          <div key={lbl}
            onClick={() => rows.length && setModal({ title:`CR — ${lbl}`, rows })}
            style={{ background:bg, border:`1.5px solid ${borderC}`, borderRadius:8, padding:"14px 16px",
              cursor:rows.length?"pointer":"default", transition:"box-shadow .15s, transform .15s" }}
            onMouseEnter={e => { if(rows.length){e.currentTarget.style.boxShadow="0 4px 12px rgba(0,0,0,0.12)";e.currentTarget.style.transform="translateY(-1px)";}}}
            onMouseLeave={e => { e.currentTarget.style.boxShadow="none";e.currentTarget.style.transform="none"; }}>
            <div style={{ fontSize:10, fontWeight:700, color:textCol, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:4 }}>{lbl}</div>
            <div style={{ fontSize:28, fontWeight:800, color:textCol, lineHeight:1 }}>{rows.length}</div>
            <div style={{ fontSize:11, color:textCol, opacity:0.8, marginTop:4, fontWeight:600 }}>
              {hours > 0 ? `${hours.toLocaleString()} hrs estimated` : "—"}
            </div>
            {rows.length > 0 && <div style={{ fontSize:10, color:textCol, opacity:0.7, marginTop:4 }}>Click to drill down →</div>}
          </div>
        ))}
      </div>

      {/* Summary bar */}
      <div style={{ background:"#f8fafc", border:`1px solid ${C.border}`, borderRadius:8, padding:"10px 14px", display:"flex", gap:20, flexWrap:"wrap", alignItems:"center" }}>
        <span style={{ fontSize:12, color:C.muted }}><b style={{ color:C.text }}>Total CRs:</b> {cr.all.length}</span>
        <span style={{ fontSize:12, color:C.muted }}><b style={{ color:C.text }}>Total Hours:</b> {cr.totalHours.toLocaleString()}</span>
        <span style={{ fontSize:11, color:C.muted, fontStyle:"italic" }}>Approved includes Inform-Accepted (Reviewed)</span>
      </div>

      {/* Full CR table */}
      <Card style={{ padding:0 }}>
        <div style={{ padding:"12px 16px", background:"#d0d5de", borderRadius:"10px 10px 0 0", borderBottom:`1px solid ${C.border}` }}>
          <div style={{ fontSize:10, fontWeight:700, color:C.text, textTransform:"uppercase", letterSpacing:"0.06em" }}>
            All Change Requests
            <span style={{ fontSize:9, color:C.muted, fontWeight:400, textTransform:"none", marginLeft:6 }}>· {cr.all.length} total</span>
          </div>
        </div>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
            <thead>
              <tr style={{ background:"#162f50" }}>
                {["RAID ID","Decision Status","Hours Est.","Component","Experience","Description","Owner","Primary Team"].map((h,i,arr) => (
                  <th key={h} style={{ padding:"8px 10px", textAlign:"left", color:"#fff", fontWeight:700, fontSize:10,
                    whiteSpace:"nowrap", borderRight:i<arr.length-1?"1px solid rgba(255,255,255,0.1)":"none" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {cr.all.map((r,i) => {
                const bucket = BUCKETS.find(b => b.rows.includes(r));
                const decStatus = bucket ? bucket.lbl : String(r[K.crStatus]||"—").trim();
                const textCol = bucket ? bucket.textCol : C.muted;
                const bg = bucket ? bucket.bg : "#f8fafc";
                const hours = (() => { const v=String(r[K.crHours]||"").replace(/[^0-9.]/g,""); const n=parseFloat(v); return isNaN(n)?null:Math.round(n); })();
                return (
                  <tr key={i} style={{ background:i%2===0?C.white:"#f7f9fc", borderBottom:`1px solid ${C.border}`, verticalAlign:"top" }}>
                    <td style={{ padding:"8px 10px", fontWeight:700, color:C.navyLight, whiteSpace:"nowrap" }}>{String(r[K.id]||"—")}</td>
                    <td style={{ padding:"8px 10px" }}>
                      <span style={{ background:bg, color:textCol, border:`1px solid ${bucket?bucket.borderC:C.border}`, borderRadius:4, padding:"2px 8px", fontSize:10, fontWeight:700, whiteSpace:"nowrap" }}>{decStatus}</span>
                    </td>
                    <td style={{ padding:"8px 10px", color:C.text, textAlign:"right", fontWeight:600 }}>{hours ? `${hours}h` : "—"}</td>
                    <td style={{ padding:"8px 10px", color:C.text, wordBreak:"break-word", maxWidth:130 }}>{String(r[K.component]||"—")}</td>
                    <td style={{ padding:"8px 10px", color:C.muted, wordBreak:"break-word", maxWidth:90 }}>{String(r[K.experience]||"—")}</td>
                    <td style={{ padding:"8px 10px", color:C.text, wordBreak:"break-word", maxWidth:300, lineHeight:1.5 }}>{String(r[K.desc]||"—")}</td>
                    <td style={{ padding:"8px 10px", color:C.muted, whiteSpace:"nowrap" }}>{String(r[K.owner]||"—")}</td>
                    <td style={{ padding:"8px 10px", color:C.muted, whiteSpace:"nowrap" }}>{String(r[teamKey]||"—")}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      {/* Drill-down modal */}
      {modal && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:1000, display:"flex", alignItems:"center", justifyContent:"center" }}
          onClick={() => setModal(null)}>
          <div style={{ background:C.white, borderRadius:10, width:"95%", maxWidth:1000, maxHeight:"88vh", display:"flex", flexDirection:"column", boxShadow:"0 24px 60px rgba(0,0,0,0.3)" }}
            onClick={e => e.stopPropagation()}>
            <div style={{ background:C.headerBg, padding:"12px 20px", borderRadius:"10px 10px 0 0", display:"flex", justifyContent:"space-between", alignItems:"center", flexShrink:0 }}>
              <div style={{ color:"#fff", fontWeight:700, fontSize:13 }}>{modal.title} <span style={{ opacity:.6, fontWeight:400 }}>({modal.rows.length})</span></div>
              <button onClick={() => setModal(null)} style={{ background:"rgba(255,255,255,0.15)", border:"none", color:"#fff", borderRadius:5, padding:"5px 14px", cursor:"pointer", fontSize:13, fontWeight:600 }}>✕</button>
            </div>
            <div style={{ overflowY:"auto", flex:1 }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
                <thead style={{ position:"sticky", top:0, background:"#f0f4f8", zIndex:2 }}>
                  <tr>
                    {["RAID ID","Hours Est.","Component","Experience","Description","Owner","Primary Team"].map(h => (
                      <th key={h} style={{ padding:"8px 10px", textAlign:"left", color:C.muted, fontWeight:700, whiteSpace:"nowrap", borderBottom:`2px solid ${C.border}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {modal.rows.map((r,i) => {
                    const hours = (() => { const v=String(r[K.crHours]||"").replace(/[^0-9.]/g,""); const n=parseFloat(v); return isNaN(n)?null:Math.round(n); })();
                    return (
                      <tr key={i} style={{ background:i%2===0?C.white:"#f9fafb", borderBottom:`1px solid ${C.border}` }}>
                        <td style={{ padding:"7px 10px", fontWeight:700, color:C.navyLight }}>{String(r[K.id]||"—")}</td>
                        <td style={{ padding:"7px 10px", color:C.text, fontWeight:600, textAlign:"right" }}>{hours?`${hours}h`:"—"}</td>
                        <td style={{ padding:"7px 10px", color:C.text, wordBreak:"break-word", maxWidth:130 }}>{String(r[K.component]||"—")}</td>
                        <td style={{ padding:"7px 10px", color:C.muted }}>{String(r[K.experience]||"—")}</td>
                        <td style={{ padding:"7px 10px", color:C.text, wordBreak:"break-word", maxWidth:300, lineHeight:1.5 }}>{String(r[K.desc]||"—")}</td>
                        <td style={{ padding:"7px 10px", color:C.muted }}>{String(r[K.owner]||"—")}</td>
                        <td style={{ padding:"7px 10px", color:C.muted }}>{String(r[teamKey]||"—")}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── BACKLOG CHART DRILL-DOWN MODAL ──────────────────────────────────────────
function BacklogChartDrillModal({ title, rows, K, teamKey, colConfig, COL_KEYS, tableWidth, priColorMap, sColor, renderCritPath, onClose }) {
  const [expF,  setExpF]  = useState("All");
  const [compF, setCompF] = useState("All");
  const [typeF, setTypeF] = useState("All");
  const [teamF, setTeamF] = useState("All");

  const allExps  = Array.from(new Set(rows.map(r => String(r[K.experience]||"").trim()).filter(Boolean))).sort();
  const allComps = Array.from(new Set(rows.map(r => String(r[K.component]||"").trim()).filter(Boolean))).sort();
  const allTypes = Array.from(new Set(rows.map(r => String(r[K.type]||"").trim()).filter(Boolean))).sort();
  const allTeams = Array.from(new Set(rows.map(r => String(r[teamKey]||"").trim()).filter(Boolean))).sort();

  const mE  = r => expF  === "All" || String(r[K.experience]||"").trim() === expF;
  const mC  = r => compF === "All" || String(r[K.component]||"").trim()  === compF;
  const mT  = r => typeF === "All" || String(r[K.type]||"").trim()       === typeF;
  const mTm = r => teamF === "All" || String(r[teamKey]||"").trim()      === teamF;

  const filtered = rows.filter(r => mE(r) && mC(r) && mT(r) && mTm(r));

  const pill = (val, isActive, count, onClick, col) => {
    const has = count > 0;
    return (
      <button key={val} onClick={onClick} disabled={!has && val !== "All"}
        style={{ display:"flex", alignItems:"center", gap:4, padding:"4px 10px", borderRadius:20,
          border:`2px solid ${isActive?(col||C.navyLight):has?"#b0bbc8":C.border}`,
          background:isActive?(col||C.navyLight):C.white, color:isActive?"#fff":has?C.text:C.muted,
          cursor:has||val==="All"?"pointer":"default", fontSize:10, fontWeight:700,
          transition:"all .12s", opacity:!has&&val!=="All"?0.4:1 }}>
        {val}
        <span style={{background:isActive?"rgba(255,255,255,0.25)":"#f1f5f9",color:isActive?"#fff":C.text,
          borderRadius:10,padding:"1px 6px",fontSize:10,fontWeight:800,minWidth:18,textAlign:"center"}}>{count}</span>
      </button>
    );
  };

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}}
      onClick={onClose}>
      <div style={{background:C.white,borderRadius:10,width:"98%",maxWidth:1300,maxHeight:"92vh",display:"flex",flexDirection:"column",boxShadow:"0 24px 60px rgba(0,0,0,0.3)"}}
        onClick={e=>e.stopPropagation()}>

        {/* Header */}
        <div style={{background:C.headerBg,padding:"12px 20px",borderRadius:"10px 10px 0 0",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
          <div style={{color:"#fff",fontWeight:700,fontSize:13}}>Backlog — {title} <span style={{opacity:.6,fontWeight:400}}>({filtered.length} of {rows.length})</span></div>
          <button onClick={onClose} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",borderRadius:5,padding:"5px 14px",cursor:"pointer",fontSize:13,fontWeight:600}}>✕</button>
        </div>

        {/* Filters */}
        <div style={{padding:"10px 16px",borderBottom:`1px solid ${C.border}`,background:"#f8fafc",display:"flex",flexDirection:"column",gap:8,flexShrink:0}}>
          {/* Experience */}
          {allExps.length > 1 && (
            <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:10,color:C.muted,fontWeight:600,marginRight:2}}>Experience</span>
              {pill("All",expF==="All",rows.filter(r=>mC(r)&&mT(r)&&mTm(r)).length,()=>setExpF("All"))}
              {allExps.map(e=>pill(e,expF===e,rows.filter(r=>String(r[K.experience]||"").trim()===e&&mC(r)&&mT(r)&&mTm(r)).length,()=>setExpF(expF===e?"All":e)))}
            </div>
          )}
          {/* Type + Team */}
          <div style={{display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
            {allTypes.length > 1 && (
              <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontSize:10,color:C.muted,fontWeight:600,marginRight:2}}>Type</span>
                {pill("All",typeF==="All",rows.filter(r=>mE(r)&&mC(r)&&mTm(r)).length,()=>setTypeF("All"))}
                {allTypes.map(t=>pill(t,typeF===t,rows.filter(r=>String(r[K.type]||"").trim()===t&&mE(r)&&mC(r)&&mTm(r)).length,()=>setTypeF(typeF===t?"All":t)))}
              </div>
            )}
            {allTeams.length > 1 && (
              <>
                {allTypes.length > 1 && <div style={{width:1,height:18,background:C.border}}/>}
                <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
                  <span style={{fontSize:10,color:C.muted,fontWeight:600,marginRight:2}}>Team</span>
                  {pill("All",teamF==="All",rows.filter(r=>mE(r)&&mC(r)&&mT(r)).length,()=>setTeamF("All"))}
                  {allTeams.map(t=>pill(t,teamF===t,rows.filter(r=>String(r[teamKey]||"").trim()===t&&mE(r)&&mC(r)&&mT(r)).length,()=>setTeamF(teamF===t?"All":t)))}
                </div>
              </>
            )}
          </div>
          {/* Component */}
          {allComps.length > 1 && (
            <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:10,color:C.muted,fontWeight:600,marginRight:2}}>Component</span>
              {pill("All",compF==="All",rows.filter(r=>mE(r)&&mT(r)&&mTm(r)).length,()=>setCompF("All"))}
              {allComps.map(c=>pill(c,compF===c,rows.filter(r=>String(r[K.component]||"").trim()===c&&mE(r)&&mT(r)&&mTm(r)).length,()=>setCompF(compF===c?"All":c)))}
            </div>
          )}
          <div style={{fontSize:10,color:C.muted}}>
            Showing <b style={{color:C.text}}>{filtered.length}</b> of <b style={{color:C.text}}>{rows.length}</b> items
          </div>
        </div>

        {/* Table */}
        <div style={{overflowX:"auto",overflowY:"auto",flex:1}}>
          <table style={{borderCollapse:"collapse",fontSize:11,tableLayout:"fixed",width:tableWidth+"px",minWidth:"100%"}}>
            <thead style={{position:"sticky",top:0,zIndex:2}}>
              <tr style={{background:"#162f50"}}>
                {COL_KEYS.filter(([key])=>colConfig[key]?.visible).map(([key,label],idx,arr)=>(
                  <th key={key} style={{padding:"8px 10px",textAlign:"left",color:"#fff",fontWeight:700,fontSize:10,
                    width:colConfig[key].width,whiteSpace:"nowrap",
                    borderRight:idx<arr.length-1?"1px solid rgba(255,255,255,0.1)":"none"}}>{label}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length===0?(
                <tr><td colSpan={COL_KEYS.filter(([k])=>colConfig[k]?.visible).length}
                  style={{padding:"24px",textAlign:"center",color:C.muted}}>No items match filters</td></tr>
              ):filtered.map((r,i)=>{
                const status=String(r[K.status]||"").trim();
                const sCol=sColor(status);
                const pri=String(r[K.priority]||"").trim();
                const priCol=priColorMap[pri]||C.muted;
                const due=daysUntil(r[K.date]);
                const dueStr=fmtDate(r[K.date]);
                const dueCol=due!=null&&due<=7?C.delayed:due!=null&&due<=14?C.gold:C.muted;
                return (
                  <tr key={i} style={{background:i%2===0?C.white:"#f7f9fc",borderBottom:`1px solid ${C.border}`,verticalAlign:"top"}}>
                    {colConfig.raidId?.visible    &&<td style={{padding:"8px 10px",fontWeight:700,color:C.navyLight,wordBreak:"break-word",width:colConfig.raidId.width}}>{String(r[K.id]||"—")}</td>}
                    {colConfig.status?.visible    &&<td style={{padding:"8px 10px",width:colConfig.status.width}}><span style={{background:sCol+"20",color:sCol,border:`1px solid ${sCol}40`,borderRadius:4,padding:"2px 6px",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>{status||"—"}</span></td>}
                    {colConfig.type?.visible      &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.type.width}}>{String(r[K.type]||"—")}</td>}
                    {colConfig.priority?.visible  &&<td style={{padding:"8px 10px",width:colConfig.priority.width}}>{pri?<span style={{background:priCol+"20",color:priCol,border:`1px solid ${priCol}40`,borderRadius:4,padding:"2px 6px",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>{pri}</span>:<span style={{color:C.muted}}>—</span>}</td>}
                    {colConfig.component?.visible &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.component.width}}>{String(r[K.component]||"—")}</td>}
                    {colConfig.experience?.visible&&<td style={{padding:"8px 10px",color:C.muted,wordBreak:"break-word",width:colConfig.experience.width}}>{String(r[K.experience]||"—")}</td>}
                    {colConfig.topic?.visible     &&<td style={{padding:"8px 10px",color:C.muted,wordBreak:"break-word",width:colConfig.topic.width}}>{String(r[K.topic]||"—")}</td>}
                    {colConfig.desc?.visible      &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",lineHeight:1.5,width:colConfig.desc.width}}>{String(r[K.desc]||"—")}</td>}
                    {colConfig.comment?.visible   &&<td style={{padding:"8px 10px",color:C.muted,wordBreak:"break-word",lineHeight:1.5,width:colConfig.comment.width}}>{String(r[K.comment]||"—")}</td>}
                    {colConfig.owner?.visible     &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.owner.width}}>{String(r[K.owner]||"—")}</td>}
                    {colConfig.team?.visible      &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.team.width}}>{String(r[teamKey]||"—")}</td>}
                    {colConfig.critPath?.visible  &&<td style={{padding:"8px 10px",width:colConfig.critPath.width}}>{renderCritPath(r)}</td>}
                    {colConfig.dueDate?.visible   &&<td style={{padding:"8px 10px",color:dueCol,fontWeight:600,whiteSpace:"nowrap",width:colConfig.dueDate.width}}>{dueStr}</td>}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── BACKLOG TAB ─────────────────────────────────────────────────────────────
function BacklogTab({ raid }) {
  const [chartPriFilter, setChartPriFilter] = useState("All"); // chart only
  const [priorityFilter, setPriorityFilter]   = useState("All"); // table only
  const [typeFilter,     setTypeFilter]     = useState("All");
  const [compFilter,     setCompFilter]     = useState("All");
  const [teamFilter,     setTeamFilter]     = useState("All");
  const [expFilter,      setExpFilter]      = useState("All");
  const [showColPanel,   setShowColPanel]   = useState(false);
  const [drillModal,     setDrillModal]     = useState(null); // table filter drill-down
  const [chartDrill,     setChartDrill]     = useState(null); // { title, rows } chart drill-down
  const [colConfig, setColConfig] = useState({
    raidId:    { label:"RAID ID",               visible:true,  width:90  },
    status:    { label:"Status",                visible:true,  width:90  },
    type:      { label:"Type",                  visible:true,  width:90  },
    priority:  { label:"Priority",              visible:true,  width:100 },
    component: { label:"Component",             visible:true,  width:130 },
    experience:{ label:"Experience",            visible:true,  width:90  },
    topic:     { label:"Topic",                 visible:true,  width:90  },
    desc:      { label:"Description",           visible:true,  width:260 },
    comment:   { label:"Comments / Resolution", visible:true,  width:220 },
    owner:     { label:"Owner",                 visible:true,  width:110 },
    team:      { label:"Primary Team (Owner)",  visible:true,  width:140 },
    critPath:  { label:"Critical Path",         visible:true,  width:100 },
    dueDate:   { label:"Due Date",              visible:true,  width:85  },
  });

  if (!raid) return <Empty label="Upload RAID Log file above to view this tab." />;

  const K       = raid.keys;
  const teamKey = K.team || "Primary Team (Owner)";
  const rows    = raid.deferred || [];

  if (rows.length === 0) return (
    <Card>
      <div style={{ textAlign:"center", padding:"32px 0", color:C.muted, fontSize:13 }}>
        No deferred RAID items found.<br/>
        <span style={{ fontSize:11 }}>Items with Status = "Deferred" will appear here.</span>
      </div>
    </Card>
  );

  const allTypes      = Array.from(new Set(rows.map(r => String(r[K.type]||"").trim()).filter(Boolean))).sort();
  const allComps      = Array.from(new Set(rows.map(r => String(r[K.component]||"").trim()).filter(Boolean))).sort();
  const allTeams      = Array.from(new Set(rows.map(r => String(r[teamKey]||"").trim()).filter(Boolean))).sort();
  const allPriorities = Array.from(new Set(rows.map(r => String(r[K.priority]||"").trim()).filter(Boolean))).sort();
  const allExps       = Array.from(new Set(rows.map(r => String(r[K.experience]||"").trim()).filter(Boolean))).sort();

  // Cross-filter helpers
  const mP  = r => priorityFilter === "All" || String(r[K.priority]||"").trim()   === priorityFilter;
  const mT  = r => typeFilter     === "All" || String(r[K.type]||"").trim()        === typeFilter;
  const mC  = r => compFilter     === "All" || String(r[K.component]||"").trim()   === compFilter;
  const mTm = r => teamFilter     === "All" || String(r[teamKey]||"").trim()       === teamFilter;
  const mE  = r => expFilter      === "All" || String(r[K.experience]||"").trim()  === expFilter;

  const filtered = rows.filter(r => mP(r) && mT(r) && mC(r) && mTm(r) && mE(r));

  const priColors = ["#b91c1c","#c2410c","#d97706","#1d4ed8","#475569","#6366f1","#0891b2","#059669"];
  const priColorMap = {};
  allPriorities.forEach((p,i) => { priColorMap[p] = priColors[i % priColors.length]; });

  const sColor = s => { const sl=String(s||"").toLowerCase(); return sl.includes("delay")?C.delayed:sl.includes("complete")?C.complete:"#6366f1"; };
  const renderCritPath = (r) => {
    const v = String(r[K.critPath]||"").trim();
    if (!v || v==="—") return <span style={{color:C.muted}}>—</span>;
    const hi = v.toLowerCase()!=="no" && v.toLowerCase()!=="n/a";
    return <span style={{background:hi?"#fee2e2":"#f1f5f9",color:hi?C.delayed:C.muted,borderRadius:3,padding:"2px 6px",fontSize:10,fontWeight:600}}>{v}</span>;
  };

  // pill(val, isActive, count, onFilter, col, drillRows)
  // Single click = filter; clicking count badge = open drill-down
  const pill = (val, isActive, count, onFilter, col, drillRows) => {
    const has = count > 0;
    return (
      <button key={val} onClick={onFilter} disabled={!has && val!=="All"}
        style={{ display:"flex", alignItems:"center", gap:0, padding:"0", borderRadius:20,
          border:`2px solid ${isActive?(col||C.navyLight):has?(col?col+"80":"#b0bbc8"):C.border}`,
          background:isActive?(col||C.navyLight):C.white, color:isActive?"#fff":has?C.text:C.muted,
          cursor:has||val==="All"?"pointer":"default", fontSize:10, fontWeight:700,
          transition:"all .12s", opacity:!has&&val!=="All"?0.4:1, overflow:"hidden" }}>
        <span style={{padding:"4px 8px 4px 10px"}}>{val}</span>
        <span
          onClick={e => { e.stopPropagation(); if(has && drillRows) setChartDrill({title:val, rows:drillRows}); }}
          style={{background:isActive?"rgba(255,255,255,0.25)":"rgba(0,0,0,0.07)",color:isActive?"#fff":C.text,
            padding:"4px 8px",fontSize:10,fontWeight:800,minWidth:24,textAlign:"center",
            borderLeft:`1px solid ${isActive?"rgba(255,255,255,0.2)":"rgba(0,0,0,0.1)"}`,
            cursor:has&&drillRows?"pointer":"default",
            display:"flex",alignItems:"center",justifyContent:"center"}}>
          {count}
        </span>
      </button>
    );
  };

  const visibleCols = Object.values(colConfig).filter(c=>c.visible).length;
  const tableWidth  = Object.values(colConfig).filter(c=>c.visible).reduce((s,c)=>s+c.width,0);
  const COL_KEYS = [
    ["raidId","RAID ID"],["status","Status"],["type","Type"],["priority","Priority"],
    ["component","Component"],["experience","Experience"],["topic","Topic"],
    ["desc","Description"],["comment","Comments / Resolution"],["owner","Owner"],
    ["team","Primary Team (Owner)"],["critPath","Critical Path"],["dueDate","Due Date"],
  ];

  return (
    <div style={{display:"flex",flexDirection:"column",gap:14}}>

      {/* Experience × Priority chart — bars clickable as filters */}
      <Card>
        <div style={{fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:12}}>
          By Experience — Count by Priority
          {chartPriFilter!=="All" && (
            <span onClick={()=>setChartPriFilter("All")}
              style={{marginLeft:10,fontSize:10,color:C.accent,cursor:"pointer",fontWeight:600,textTransform:"none"}}>
              ✕ Clear priority filter
            </span>
          )}
        </div>
        {(() => {
          const experiences = Array.from(new Set(rows.filter(r=>mT(r)&&mC(r)&&mTm(r)&&mE(r)).map(r=>String(r[K.experience]||"").trim()).filter(Boolean))).sort();
          if (experiences.length===0) return <div style={{color:C.muted,fontSize:12,textAlign:"center",padding:"16px 0"}}>No Experience data found.</div>;
          const data = experiences.map(exp => {
            const expRows = rows.filter(r=>String(r[K.experience]||"").trim()===exp && mT(r)&&mC(r)&&mTm(r)&&mE(r));
            const priCounts = {};
            allPriorities.forEach(p => { priCounts[p] = expRows.filter(r=>String(r[K.priority]||"").trim()===p).length; });
            return {exp, priCounts, total:expRows.length};
          }).sort((a,b)=>b.total-a.total);
          const maxTotal = Math.max(...data.map(d=>d.total),1);
          return (
            <div style={{display:"flex",flexDirection:"column",gap:7}}>
              {data.map(({exp,priCounts,total})=>(
                <div key={exp} style={{display:"flex",alignItems:"center",gap:8}}>
                  <div style={{minWidth:160,fontSize:11,fontWeight:600,color:C.text,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}} title={exp}>{exp}</div>
                  <div style={{flex:1,display:"flex",height:20,borderRadius:4,overflow:"hidden",background:"#f0f2f5"}}>
                    {allPriorities.map(p=>priCounts[p]>0&&(
                      <div key={p} onClick={()=>setChartPriFilter(chartPriFilter===p?"All":p)}
                        style={{width:`${(priCounts[p]/maxTotal)*100}%`,background:priColorMap[p],cursor:"pointer",
                          display:"flex",alignItems:"center",justifyContent:"center",minWidth:4,
                          opacity:chartPriFilter!=="All"&&chartPriFilter!==p?0.3:1,transition:"opacity .15s",
                          outline:chartPriFilter===p?"2px solid rgba(0,0,0,0.3) inset":""}} >
                        {priCounts[p]>=2&&<span style={{color:"#fff",fontSize:10,fontWeight:700}}>{priCounts[p]}</span>}
                      </div>
                    ))}
                  </div>
                  <div style={{display:"flex",gap:4,flexWrap:"wrap",minWidth:130}}>
                    {allPriorities.map(p=>priCounts[p]>0&&(
                      <span key={p} onClick={()=>{ setChartPriFilter(chartPriFilter===p?"All":p); setChartDrill({title:p, rows:rows.filter(r=>String(r[K.experience]||"").trim()===exp&&String(r[K.priority]||"").trim()===p&&mT(r)&&mC(r)&&mTm(r)&&mE(r))}); }}
                        style={{background:priColorMap[p]+(chartPriFilter===p?"":"20"),color:chartPriFilter===p?"#fff":priColorMap[p],
                          border:`1px solid ${priColorMap[p]}60`,borderRadius:3,padding:"2px 6px",fontSize:10,fontWeight:700,
                          cursor:"pointer",opacity:chartPriFilter!=="All"&&chartPriFilter!==p?0.4:1,transition:"all .15s"}}>
                        {p.replace(/^\d+\s*-\s*/,"")}: {priCounts[p]}
                      </span>
                    ))}
                  </div>
                </div>
              ))}
              {/* Priority legend — acts as filter */}
              <div style={{display:"flex",gap:8,marginTop:8,flexWrap:"wrap",borderTop:`1px solid ${C.border}`,paddingTop:10}}>
                <span style={{fontSize:10,color:C.muted,fontWeight:600,marginRight:4}}>Filter by priority:</span>
                {pill("All", chartPriFilter==="All", rows.filter(r=>mT(r)&&mC(r)&&mTm(r)&&mE(r)).length, ()=>setChartPriFilter("All"), null, rows.filter(r=>mT(r)&&mC(r)&&mTm(r)&&mE(r)))}
                {allPriorities.map(p=>pill(p, chartPriFilter===p,
                  rows.filter(r=>String(r[K.priority]||"").trim()===p&&mT(r)&&mC(r)&&mTm(r)&&mE(r)).length,
                  ()=>setChartPriFilter(chartPriFilter===p?"All":p), priColorMap[p],
                  rows.filter(r=>String(r[K.priority]||"").trim()===p&&mT(r)&&mC(r)&&mTm(r)&&mE(r))))}
              </div>
            </div>
          );
        })()}
      </Card>

      {/* Table card */}
      <Card style={{padding:0}}>
        <div style={{padding:"12px 16px",background:"#d0d5de",borderRadius:"10px 10px 0 0",borderBottom:`1px solid ${C.border}`}}>
          <div style={{fontSize:10,fontWeight:700,color:C.text,textTransform:"uppercase",letterSpacing:"0.06em"}}>
            Deferred RAID Items
            <span style={{fontSize:9,color:C.muted,fontWeight:400,textTransform:"none",marginLeft:6}}>· Status = Deferred</span>
          </div>
        </div>

        {/* Filters */}
        <div style={{padding:"10px 16px",borderBottom:`1px solid ${C.border}`,background:C.white,display:"flex",flexDirection:"column",gap:10}}>

          {/* Row 1: Team */}
          {allTeams.length>0 && (
            <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:10,color:C.muted,fontWeight:600,marginRight:2}}>Team</span>
              {pill("All",teamFilter==="All",rows.filter(r=>mP(r)&&mT(r)&&mC(r)&&mE(r)).length,()=>setTeamFilter("All"),null,rows.filter(r=>mP(r)&&mT(r)&&mC(r)&&mE(r)))}
              {allTeams.map(t=>pill(t,teamFilter===t,rows.filter(r=>String(r[teamKey]||"").trim()===t&&mP(r)&&mT(r)&&mC(r)&&mE(r)).length,()=>setTeamFilter(teamFilter===t?"All":t),null,rows.filter(r=>String(r[teamKey]||"").trim()===t&&mP(r)&&mT(r)&&mC(r)&&mE(r))))}
              <button onClick={()=>setShowColPanel(p=>!p)}
                style={{marginLeft:"auto",padding:"4px 12px",borderRadius:5,border:`1px solid ${showColPanel?C.navyLight:C.border}`,
                  background:showColPanel?C.navyLight:C.white,color:showColPanel?"#fff":C.muted,cursor:"pointer",fontSize:10,fontWeight:600}}>⚙ Columns</button>
            </div>
          )}

          {/* Row 2: Type + Experience */}
          <div style={{display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
            <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:10,color:C.muted,fontWeight:600,marginRight:2}}>Type</span>
              {pill("All",typeFilter==="All",rows.filter(r=>mP(r)&&mC(r)&&mTm(r)&&mE(r)).length,()=>setTypeFilter("All"),null,rows.filter(r=>mP(r)&&mC(r)&&mTm(r)&&mE(r)))}
              {allTypes.map(t=>pill(t,typeFilter===t,rows.filter(r=>String(r[K.type]||"").trim()===t&&mP(r)&&mC(r)&&mTm(r)&&mE(r)).length,()=>setTypeFilter(typeFilter===t?"All":t),null,rows.filter(r=>String(r[K.type]||"").trim()===t&&mP(r)&&mC(r)&&mTm(r)&&mE(r))))}
            </div>
            {allExps.length>0&&(
              <>
                <div style={{width:1,height:18,background:C.border}}/>
                <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
                  <span style={{fontSize:10,color:C.muted,fontWeight:600,marginRight:2}}>Experience</span>
                  {pill("All",expFilter==="All",rows.filter(r=>mP(r)&&mT(r)&&mC(r)&&mTm(r)).length,()=>setExpFilter("All"),null,rows.filter(r=>mP(r)&&mT(r)&&mC(r)&&mTm(r)))}
                  {allExps.map(e=>pill(e,expFilter===e,rows.filter(r=>String(r[K.experience]||"").trim()===e&&mP(r)&&mT(r)&&mC(r)&&mTm(r)).length,()=>setExpFilter(expFilter===e?"All":e),null,rows.filter(r=>String(r[K.experience]||"").trim()===e&&mP(r)&&mT(r)&&mC(r)&&mTm(r))))}
                </div>
              </>
            )}
            {allTeams.length===0&&(
              <button onClick={()=>setShowColPanel(p=>!p)}
                style={{marginLeft:"auto",padding:"4px 12px",borderRadius:5,border:`1px solid ${showColPanel?C.navyLight:C.border}`,
                  background:showColPanel?C.navyLight:C.white,color:showColPanel?"#fff":C.muted,cursor:"pointer",fontSize:10,fontWeight:600}}>⚙ Columns</button>
            )}
          </div>

          {/* Row 3: Component */}
          {allComps.length>0&&(
            <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:10,color:C.muted,fontWeight:600,marginRight:2}}>Component</span>
              {pill("All",compFilter==="All",rows.filter(r=>mP(r)&&mT(r)&&mTm(r)&&mE(r)).length,()=>setCompFilter("All"),null,rows.filter(r=>mP(r)&&mT(r)&&mTm(r)&&mE(r)))}
              {allComps.map(c=>pill(c,compFilter===c,rows.filter(r=>String(r[K.component]||"").trim()===c&&mP(r)&&mT(r)&&mTm(r)&&mE(r)).length,()=>setCompFilter(compFilter===c?"All":c),null,rows.filter(r=>String(r[K.component]||"").trim()===c&&mP(r)&&mT(r)&&mTm(r)&&mE(r))))}
            </div>
          )}

          {/* Col config */}
          {showColPanel&&(
            <div style={{background:"#f8fafc",border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 12px"}}>
              <div style={{fontSize:11,fontWeight:700,color:C.text,marginBottom:8}}>
                Show / Hide Columns <span style={{fontSize:10,color:C.muted,fontWeight:400}}>— drag column edges to resize</span>
              </div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                {Object.entries(colConfig).map(([key,col])=>(
                  <label key={key} style={{display:"flex",alignItems:"center",gap:4,background:C.white,
                    border:`1px solid ${col.visible?C.navyLight:C.border}`,borderRadius:5,padding:"4px 9px",cursor:"pointer"}}>
                    <input type="checkbox" checked={col.visible}
                      onChange={e=>setColConfig(p=>({...p,[key]:{...p[key],visible:e.target.checked}}))}
                      style={{cursor:"pointer",width:12,height:12}}/>
                    <span style={{fontSize:10,color:col.visible?C.navyLight:C.muted,fontWeight:col.visible?700:400}}>{col.label}</span>
                  </label>
                ))}
              </div>
            </div>
          )}

          <div style={{fontSize:10,color:C.muted}}>
            Showing <b style={{color:C.text}}>{filtered.length}</b> of <b style={{color:C.text}}>{rows.length}</b> deferred items
            {priorityFilter!=="All"&&<span style={{marginLeft:6,color:priColorMap[priorityFilter],fontWeight:700}}>· Priority: {priorityFilter}</span>}
          </div>
        </div>

        {/* Table */}
        <div style={{overflowX:"auto"}}>
          <table style={{borderCollapse:"collapse",fontSize:11,tableLayout:"fixed",width:tableWidth+"px",minWidth:"100%"}}>
            <thead>
              <tr style={{background:"#162f50"}}>
                {COL_KEYS.filter(([key])=>colConfig[key]?.visible).map(([key,label],idx,arr)=>(
                  <th key={key} style={{padding:"8px 10px",textAlign:"left",color:"#fff",fontWeight:700,fontSize:10,
                    width:colConfig[key].width,position:"relative",
                    borderRight:idx<arr.length-1?"1px solid rgba(255,255,255,0.1)":"none"}}>
                    {label}
                    <div onMouseDown={e=>{
                      e.preventDefault();
                      const sx=e.clientX,sw=colConfig[key].width;
                      const mv=m=>setColConfig(p=>({...p,[key]:{...p[key],width:Math.max(50,sw+m.clientX-sx)}}));
                      const up=()=>{window.removeEventListener("mousemove",mv);window.removeEventListener("mouseup",up);};
                      window.addEventListener("mousemove",mv);window.addEventListener("mouseup",up);
                    }}
                    style={{position:"absolute",right:0,top:0,bottom:0,width:6,cursor:"col-resize",zIndex:10}}
                    onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.3)"}
                    onMouseLeave={e=>e.currentTarget.style.background="transparent"}/>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length===0?(
                <tr><td colSpan={visibleCols} style={{padding:"24px",textAlign:"center",color:C.muted}}>No items match current filters</td></tr>
              ):filtered.map((r,i)=>{
                const status=String(r[K.status]||"").trim();
                const sCol=sColor(status);
                const pri=String(r[K.priority]||"").trim();
                const priCol=priColorMap[pri]||C.muted;
                const due=daysUntil(r[K.date]);
                const dueStr=fmtDate(r[K.date]);
                const dueCol=due!=null&&due<=7?C.delayed:due!=null&&due<=14?C.gold:C.muted;
                return (
                  <tr key={i} style={{background:i%2===0?C.white:"#f7f9fc",borderBottom:`1px solid ${C.border}`,verticalAlign:"top"}}>
                    {colConfig.raidId?.visible    &&<td style={{padding:"8px 10px",fontWeight:700,color:C.navyLight,wordBreak:"break-word",width:colConfig.raidId.width}}>{String(r[K.id]||"—")}</td>}
                    {colConfig.status?.visible    &&<td style={{padding:"8px 10px",width:colConfig.status.width}}>
                      <span style={{background:sCol+"20",color:sCol,border:`1px solid ${sCol}40`,borderRadius:4,padding:"2px 6px",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>{status||"—"}</span>
                    </td>}
                    {colConfig.type?.visible      &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.type.width}}>{String(r[K.type]||"—")}</td>}
                    {colConfig.priority?.visible  &&<td style={{padding:"8px 10px",width:colConfig.priority.width}}>
                      {pri?<span style={{background:priCol+"20",color:priCol,border:`1px solid ${priCol}40`,borderRadius:4,padding:"2px 6px",fontSize:10,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap"}}
                        onClick={()=>setPriorityFilter(priorityFilter===pri?"All":pri)}>{pri}</span>:<span style={{color:C.muted}}>—</span>}
                    </td>}
                    {colConfig.component?.visible &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.component.width}}>{String(r[K.component]||"—")}</td>}
                    {colConfig.experience?.visible&&<td style={{padding:"8px 10px",color:C.muted,wordBreak:"break-word",width:colConfig.experience.width}}>{String(r[K.experience]||"—")}</td>}
                    {colConfig.topic?.visible     &&<td style={{padding:"8px 10px",color:C.muted,wordBreak:"break-word",width:colConfig.topic.width}}>{String(r[K.topic]||"—")}</td>}
                    {colConfig.desc?.visible      &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",lineHeight:1.5,width:colConfig.desc.width}}>{String(r[K.desc]||"—")}</td>}
                    {colConfig.comment?.visible   &&<td style={{padding:"8px 10px",color:C.muted,wordBreak:"break-word",lineHeight:1.5,width:colConfig.comment.width}}>{String(r[K.comment]||"—")}</td>}
                    {colConfig.owner?.visible     &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.owner.width}}>{String(r[K.owner]||"—")}</td>}
                    {colConfig.team?.visible      &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.team.width}}>{String(r[teamKey]||"—")}</td>}
                    {colConfig.critPath?.visible  &&<td style={{padding:"8px 10px",width:colConfig.critPath.width}}>{renderCritPath(r)}</td>}
                    {colConfig.dueDate?.visible   &&<td style={{padding:"8px 10px",color:dueCol,fontWeight:600,whiteSpace:"nowrap",width:colConfig.dueDate.width}}>{dueStr}</td>}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
      {/* Table filter drill-down modal (simple, no extra filters) */}
      {drillModal && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}}
          onClick={()=>setDrillModal(null)}>
          <div style={{background:C.white,borderRadius:10,width:"98%",maxWidth:1300,maxHeight:"92vh",display:"flex",flexDirection:"column",boxShadow:"0 24px 60px rgba(0,0,0,0.3)"}}
            onClick={e=>e.stopPropagation()}>
            <div style={{background:C.headerBg,padding:"12px 20px",borderRadius:"10px 10px 0 0",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
              <div style={{color:"#fff",fontWeight:700,fontSize:13}}>Backlog — {drillModal.title} <span style={{opacity:.6,fontWeight:400}}>({drillModal.rows.length} items)</span></div>
              <button onClick={()=>setDrillModal(null)} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",borderRadius:5,padding:"5px 14px",cursor:"pointer",fontSize:13,fontWeight:600}}>✕</button>
            </div>
            <div style={{overflowX:"auto",overflowY:"auto",flex:1}}>
              <table style={{borderCollapse:"collapse",fontSize:11,tableLayout:"fixed",width:tableWidth+"px",minWidth:"100%"}}>
                <thead style={{position:"sticky",top:0,zIndex:2}}>
                  <tr style={{background:"#162f50"}}>
                    {COL_KEYS.filter(([key])=>colConfig[key]?.visible).map(([key,label],idx,arr)=>(
                      <th key={key} style={{padding:"8px 10px",textAlign:"left",color:"#fff",fontWeight:700,fontSize:10,
                        width:colConfig[key].width,whiteSpace:"nowrap",
                        borderRight:idx<arr.length-1?"1px solid rgba(255,255,255,0.1)":"none"}}>{label}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {drillModal.rows.map((r,i)=>{
                    const status=String(r[K.status]||"").trim();
                    const sCol=sColor(status);
                    const pri=String(r[K.priority]||"").trim();
                    const priCol=priColorMap[pri]||C.muted;
                    const due=daysUntil(r[K.date]);
                    const dueStr=fmtDate(r[K.date]);
                    const dueCol=due!=null&&due<=7?C.delayed:due!=null&&due<=14?C.gold:C.muted;
                    return (
                      <tr key={i} style={{background:i%2===0?C.white:"#f7f9fc",borderBottom:`1px solid ${C.border}`,verticalAlign:"top"}}>
                        {colConfig.raidId?.visible    &&<td style={{padding:"8px 10px",fontWeight:700,color:C.navyLight,wordBreak:"break-word",width:colConfig.raidId.width}}>{String(r[K.id]||"—")}</td>}
                        {colConfig.status?.visible    &&<td style={{padding:"8px 10px",width:colConfig.status.width}}><span style={{background:sCol+"20",color:sCol,border:`1px solid ${sCol}40`,borderRadius:4,padding:"2px 6px",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>{status||"—"}</span></td>}
                        {colConfig.type?.visible      &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.type.width}}>{String(r[K.type]||"—")}</td>}
                        {colConfig.priority?.visible  &&<td style={{padding:"8px 10px",width:colConfig.priority.width}}>{pri?<span style={{background:priCol+"20",color:priCol,border:`1px solid ${priCol}40`,borderRadius:4,padding:"2px 6px",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>{pri}</span>:<span style={{color:C.muted}}>—</span>}</td>}
                        {colConfig.component?.visible &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.component.width}}>{String(r[K.component]||"—")}</td>}
                        {colConfig.experience?.visible&&<td style={{padding:"8px 10px",color:C.muted,wordBreak:"break-word",width:colConfig.experience.width}}>{String(r[K.experience]||"—")}</td>}
                        {colConfig.topic?.visible     &&<td style={{padding:"8px 10px",color:C.muted,wordBreak:"break-word",width:colConfig.topic.width}}>{String(r[K.topic]||"—")}</td>}
                        {colConfig.desc?.visible      &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",lineHeight:1.5,width:colConfig.desc.width}}>{String(r[K.desc]||"—")}</td>}
                        {colConfig.comment?.visible   &&<td style={{padding:"8px 10px",color:C.muted,wordBreak:"break-word",lineHeight:1.5,width:colConfig.comment.width}}>{String(r[K.comment]||"—")}</td>}
                        {colConfig.owner?.visible     &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.owner.width}}>{String(r[K.owner]||"—")}</td>}
                        {colConfig.team?.visible      &&<td style={{padding:"8px 10px",color:C.text,wordBreak:"break-word",width:colConfig.team.width}}>{String(r[teamKey]||"—")}</td>}
                        {colConfig.critPath?.visible  &&<td style={{padding:"8px 10px",width:colConfig.critPath.width}}>{renderCritPath(r)}</td>}
                        {colConfig.dueDate?.visible   &&<td style={{padding:"8px 10px",color:dueCol,fontWeight:600,whiteSpace:"nowrap",width:colConfig.dueDate.width}}>{dueStr}</td>}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Chart drill-down modal — with Experience, Component, Type, Team filters */}
      {chartDrill && <BacklogChartDrillModal
        title={chartDrill.title}
        rows={chartDrill.rows}
        K={K} teamKey={teamKey}
        colConfig={colConfig}
        COL_KEYS={COL_KEYS}
        tableWidth={tableWidth}
        priColorMap={priColorMap}
        sColor={sColor}
        renderCritPath={renderCritPath}
        onClose={()=>setChartDrill(null)}
      />}
    </div>
  );
}


// ─── RAID ANALYSIS TAB ───────────────────────────────────────────────────────
function RaidAnalysisTab({ raid }) {
  const [raidModal, setRaidModal] = useState(null);
  const [selectedTeam, setSelectedTeam] = useState(null);
  const [statusFilter, setStatusFilter] = useState("All");
  const [typeFilter, setTypeFilter] = useState("All");
  const [compFilter, setCompFilter] = useState("All");
  // Persistent column config — survives modal close/reopen
  const [modalColConfig, setModalColConfig] = useState({
    raidId:    { label:"RAID ID",                  visible:true,  width:90  },
    status:    { label:"Status",                   visible:true,  width:90  },
    type:      { label:"Type",                     visible:true,  width:90  },
    component: { label:"Component",                visible:true,  width:130 },
    experience:{ label:"Experience",               visible:true,  width:90  },
    topic:     { label:"Topic",                    visible:true,  width:90  },
    desc:      { label:"Description",              visible:true,  width:260 },
    comment:   { label:"Comments / Resolution",    visible:true,  width:220 },
    owner:     { label:"Owner",                    visible:true,  width:110 },
    team:      { label:"Primary Team (Owner)",     visible:true,  width:140 },
    critPath:  { label:"Critical Path",            visible:true,  width:100 },
    dueDate:   { label:"Due Date",                 visible:true,  width:85  },
  });
  const [showColPanel, setShowColPanel] = useState(false);
  const [colConfig, setColConfig] = useState({
    raidId:    { label:"RAID ID",              visible:true,  width:90  },
    status:    { label:"Status",               visible:true,  width:90  },
    type:      { label:"Type",                 visible:true,  width:90  },
    component: { label:"Component",            visible:true,  width:130 },
    experience:{ label:"Experience",           visible:true,  width:90  },
    topic:     { label:"Topic",                visible:true,  width:90  },
    desc:      { label:"Description",          visible:true,  width:260 },
    comment:   { label:"Comments / Resolution",visible:true,  width:220 },
    owner:     { label:"Owner",                visible:true,  width:110 },
    critPath:  { label:"Critical Path",        visible:true,  width:100 },
    dueDate:   { label:"Due Date",             visible:true,  width:85  },
  });

  if (!raid) return <Empty label="Upload RAID Log file above to view this tab." />;

  const K = raid.keys;

  // Defensive fallback — if K.team not detected, try known column names
  const teamKey = K.team 
    || Object.keys(raid.items[0] || {}).find(k => k === "Primary Team (Owner)")
    || Object.keys(raid.items[0] || {}).find(k => k === "Primary Team")
    || Object.keys(raid.items[0] || {}).find(k => /primary.?team/i.test(k))
    || "Primary Team (Owner)";
  console.log("[RAID] K.team:", K.team, "teamKey:", teamKey, "sample cols:", Object.keys(raid.items[0]||{}).filter(k => /team/i.test(k)));

  // ── Teams ─────────────────────────────────────────────────────────────────
  const allTeams = Array.from(new Set(
    raid.items.map(r => String(r[teamKey] || "").trim()).filter(Boolean)
  )).sort();

  // ── Priority chart data ───────────────────────────────────────────────────
  const raidByPriority = (() => {
    const map = {};
    raid.items
      .filter(r => String(r[K.status]||"").toLowerCase() !== "complete")
      .forEach(r => {
        const p = String(r[K.priority]||"Unknown");
        if (!map[p]) map[p] = { total:0, open:0, delayed:0, rows:[] };
        map[p].total++; map[p].rows.push(r);
        const s = String(r[K.status]||"").toLowerCase();
        if (s.includes("delay")||s.includes("off")) map[p].delayed++;
        else map[p].open++;
      });
    return map;
  })();

  // ── Team RAID table ───────────────────────────────────────────────────────
  const isCR = r => {
    const v = String(r[K.crAnalysis] || "").toLowerCase().trim();
    return ["tech reviewed - change request needed", "sd reviewed - change request needed", "ocm reviewed - change request needed"]
      .some(t => v.includes(t.slice(0, 20)));
  };

  const teamRows = selectedTeam
    ? raid.items.filter(r => String(r[teamKey]||"").trim() === selectedTeam &&
        String(r[K.status]||"").toLowerCase() !== "complete" &&
        String(r[K.status]||"").toLowerCase() !== "deferred" &&
        !isCR(r))
    : [];

  const allTypes = Array.from(new Set(teamRows.map(r => String(r[K.type]||"").trim()).filter(Boolean))).sort();
  const allComponents = Array.from(new Set(teamRows.map(r => String(r[K.component]||"").trim()).filter(Boolean))).sort();

  // Independent cross-filter helpers for team table
  const tMatchS = r => statusFilter === "All" || (statusFilter === "Delayed" ? String(r[K.status]||"").toLowerCase().includes("delay") : !String(r[K.status]||"").toLowerCase().includes("delay") && !String(r[K.status]||"").toLowerCase().includes("complete"));
  const tMatchT = r => typeFilter   === "All" || String(r[K.type]||"").trim()      === typeFilter;
  const tMatchC = r => compFilter   === "All" || String(r[K.component]||"").trim() === compFilter;

  const filteredRows = teamRows.filter(r => tMatchS(r) && tMatchT(r) && tMatchC(r));

  const statusCol = s => {
    const sl = String(s||"").toLowerCase();
    return sl.includes("delay") ? C.delayed : sl.includes("complete") ? C.complete : C.onTrack;
  };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>

      {/* ── KPI tiles — same as RAID Summary on Overview ─────────────────── */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(160px,1fr))", gap:10 }}>
        {[
          { lbl:"Open Issues",   val:raid.openIssues.length, col:C.delayed,   rows:raid.openIssues, hideType:true,  hideStatus:false,
            delayed: raid.openIssues.filter(r=>String(r[K.status]||"").toLowerCase().includes("delay")).length },
          { lbl:"Open Risks",    val:raid.openRisks.length,  col:C.gold,      rows:raid.openRisks,  hideType:true,  hideStatus:false,
            delayed: raid.openRisks.filter(r=>String(r[K.status]||"").toLowerCase().includes("delay")).length },
          { lbl:"Delayed RAIDs", val:raid.delayed.length,    col:"#7b0d0d",   rows:raid.delayed,    hideType:false, hideStatus:true,  delayed:0 },
          { lbl:"Total Open",    val:raid.open.length,       col:C.navyLight, rows:raid.open,       hideType:false, hideStatus:false, delayed:0 },
          { lbl:"Due in 8 Days", val:raid.open.filter(r=>{ const d=daysUntil(r[K.date]); return d!=null&&d>=0&&d<=8; }).length,
            col:"#b45309", rows:raid.open.filter(r=>{ const d=daysUntil(r[K.date]); return d!=null&&d>=0&&d<=8; }), hideType:false, hideStatus:false, delayed:0 },
          { lbl:"Due in 14 Days",val:raid.open.filter(r=>{ const d=daysUntil(r[K.date]); return d!=null&&d>=0&&d<=14; }).length,
            col:"#0891b2", rows:raid.open.filter(r=>{ const d=daysUntil(r[K.date]); return d!=null&&d>=0&&d<=14; }), hideType:false, hideStatus:false, delayed:0 },
        ].map(({ lbl, val, col, rows, hideType, hideStatus, delayed }) => (
          <div key={lbl} onClick={() => setRaidModal({ title:lbl, rows, hideType, hideStatus })}
            style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:7, padding:"12px 14px",
              borderTop:`3px solid ${col}`, cursor:"pointer", boxShadow:"0 1px 3px rgba(0,0,0,0.06)",
              transition:"box-shadow .15s", position:"relative" }}
            onMouseEnter={e=>e.currentTarget.style.boxShadow="0 4px 12px rgba(0,0,0,0.12)"}
            onMouseLeave={e=>e.currentTarget.style.boxShadow="0 1px 3px rgba(0,0,0,0.06)"}>
            {/* Delayed indicator badge for Issues/Risks */}
            {delayed > 0 && (
              <div style={{ position:"absolute", top:8, right:8, background:C.delayed, color:"#fff",
                borderRadius:10, padding:"2px 7px", fontSize:10, fontWeight:800, display:"flex", alignItems:"center", gap:3 }}>
                ⚠ {delayed} delayed
              </div>
            )}
            <div style={{ color:C.muted, fontSize:10, fontWeight:600, textTransform:"uppercase", letterSpacing:"0.07em", paddingRight: delayed>0?80:0 }}>{lbl}</div>
            <div style={{ color:col, fontSize:26, fontWeight:800, lineHeight:1.2 }}>{val}</div>
            <div style={{ color:C.accent, fontSize:10, marginTop:2 }}>Click to drill down →</div>
          </div>
        ))}
      </div>

      {/* ── Priority chart ───────────────────────────────────────────────── */}
      <Card>
        <div style={{ fontSize:11, fontWeight:700, color:C.muted, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:10 }}>
          By Priority — Open vs Delayed
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:7 }}>
          {Object.entries(raidByPriority)
            .sort((a,b) => String(a[0]).localeCompare(String(b[0])))
            .map(([pri, d]) => {
              const maxTotal = Math.max(...Object.values(raidByPriority).map(x=>x.total), 1);
              const openRows    = d.rows.filter(r => !String(r[K.status]||"").toLowerCase().includes("delay"));
              const delayedRows = d.rows.filter(r => String(r[K.status]||"").toLowerCase().includes("delay"));
              return (
                <div key={pri} style={{ display:"flex", alignItems:"center", gap:8 }}>
                  <div style={{ minWidth:100, fontSize:11, fontWeight:700, color:C.text, whiteSpace:"nowrap" }}>{pri}</div>
                  <div style={{ flex:1, display:"flex", height:20, borderRadius:4, overflow:"hidden", background:"#f0f2f5" }}>
                    {d.open > 0 && <div style={{ width:`${(d.open/maxTotal)*100}%`, background:C.onTrack, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", minWidth:4 }} onClick={()=>setRaidModal({ title:`${pri}`, rows:d.rows, hideStatus:false })}>{d.open >= 2 && <span style={{ color:"#fff", fontSize:10, fontWeight:700 }}>{d.open}</span>}</div>}
                    {d.delayed > 0 && <div style={{ width:`${(d.delayed/maxTotal)*100}%`, background:C.delayed, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", minWidth:4 }} onClick={()=>setRaidModal({ title:`${pri} — Delayed`, rows:d.rows, hideStatus:false })}>{d.delayed >= 2 && <span style={{ color:"#fff", fontSize:10, fontWeight:700 }}>{d.delayed}</span>}</div>}
                  </div>
                  <div style={{ display:"flex", gap:5, minWidth:110 }}>
                    <span style={{ background:C.onTrack+"20", color:"#856a00", border:`1px solid ${C.onTrack}50`, borderRadius:3, padding:"2px 7px", fontSize:10, fontWeight:700, cursor:"pointer" }} onClick={()=>openRows.length&&setRaidModal({ title:`${pri}`, rows:d.rows, hideStatus:false })}>Open: {d.open}</span>
                    <span style={{ background:C.delayed+"20", color:C.delayed, border:`1px solid ${C.delayed}40`, borderRadius:3, padding:"2px 7px", fontSize:10, fontWeight:700, cursor:"pointer" }} onClick={()=>delayedRows.length&&setRaidModal({ title:`${pri} — Delayed`, rows:d.rows, hideStatus:false })}>Del: {d.delayed}</span>
                  </div>
                </div>
              );
            })}
        </div>
        <div style={{ display:"flex", gap:12, marginTop:8 }}>
          <span style={{ display:"flex", alignItems:"center", gap:4, fontSize:10, color:C.muted }}><span style={{ width:10,height:10,borderRadius:2,background:C.onTrack,display:"inline-block" }} />Open</span>
          <span style={{ display:"flex", alignItems:"center", gap:4, fontSize:10, color:C.muted }}><span style={{ width:10,height:10,borderRadius:2,background:C.delayed,display:"inline-block" }} />Delayed</span>
        </div>
      </Card>

      {/* ── Team selector + RAID table ───────────────────────────────────── */}
      <Card style={{ padding:0 }}>
        {/* Team buttons */}
        <div style={{ padding:"12px 16px", borderBottom:`1px solid ${C.border}`, background:"#d0d5de", borderRadius:"10px 10px 0 0" }}>
          <div style={{ fontSize:10, fontWeight:700, color:C.text, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:8 }}>
            Select Team to view open RAIDs
            <span style={{ fontSize:9, color:C.muted, fontWeight:400, textTransform:"none", marginLeft:6 }}>· CRs excluded</span>
          </div>
          <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
            {allTeams.map(team => {
              const count = raid.items.filter(r =>
                String(r[teamKey]||"").trim() === team &&
                String(r[K.status]||"").toLowerCase() !== "complete"
              ).length;
              const delayed = raid.items.filter(r =>
                String(r[teamKey]||"").trim() === team &&
                String(r[K.status]||"").toLowerCase().includes("delay")
              ).length;
              const active = selectedTeam === team;
              return (
                <button key={team} onClick={() => { setSelectedTeam(active ? null : team); setStatusFilter("All"); setTypeFilter("All"); setCompFilter("All"); }}
                  style={{ display:"flex", alignItems:"center", gap:6, padding:"6px 12px",
                    borderRadius:6, border:`2px solid ${active ? C.navyLight : C.border}`,
                    background: active ? C.navyLight : C.white,
                    color: active ? "#fff" : C.text,
                    cursor:"pointer", fontSize:11, fontWeight:600, transition:"all .15s" }}>
                  {team}
                  <span style={{ background: active ? "rgba(255,255,255,0.25)" : "#f1f5f9",
                    color: active ? "#fff" : C.text,
                    borderRadius:10, padding:"1px 7px", fontSize:10, fontWeight:800 }}>{count}</span>
                  {delayed > 0 && (
                    <span style={{ background:"#fee2e2", color:C.delayed, borderRadius:10, padding:"1px 6px", fontSize:10, fontWeight:800 }}>⚠{delayed}</span>
                  )}
                </button>
              );
            })}
          </div>
        </div>

        {/* Filters + table */}
        {selectedTeam && (
          <>
            {/* Filter bar — independent cross-filters with highlight */}
            <div style={{ padding:"10px 16px", borderBottom:`1px solid ${C.border}`, background:C.white, display:"flex", flexDirection:"column", gap:10 }}>
              {(() => {
                // Reusable pill — active=selected, highlighted=has results given other filters
                const pill = (val, isActive, count, onClick, col) => {
                  const hasItems = count > 0;
                  const borderCol = isActive ? (col||C.navyLight) : hasItems ? (col ? col+"80" : C.border) : C.border;
                  const bg = isActive ? (col||C.navyLight) : C.white;
                  const textCol = isActive ? "#fff" : hasItems ? C.text : C.muted;
                  return (
                    <button key={val} onClick={onClick} disabled={!hasItems && val!=="All"}
                      style={{ display:"flex", alignItems:"center", gap:4, padding:"4px 10px", borderRadius:20,
                        border:`2px solid ${borderCol}`, background:bg, color:textCol,
                        cursor: hasItems||val==="All" ? "pointer" : "default",
                        fontSize:10, fontWeight:700, transition:"all .12s",
                        opacity: !hasItems && val!=="All" ? 0.4 : 1 }}>
                      {val}
                      <span style={{ background: isActive?"rgba(255,255,255,0.25)":"#f1f5f9",
                        color: isActive?"#fff":C.text, borderRadius:10, padding:"1px 6px", fontSize:10, fontWeight:800, minWidth:18, textAlign:"center" }}>
                        {count}
                      </span>
                    </button>
                  );
                };

                // Status counts — based on type+comp filters only
                const sCounts = {
                  all:     teamRows.filter(r => tMatchT(r) && tMatchC(r)).length,
                  delayed: teamRows.filter(r => tMatchT(r) && tMatchC(r) && String(r[K.status]||"").toLowerCase().includes("delay")).length,
                  onTrack: teamRows.filter(r => tMatchT(r) && tMatchC(r) && !String(r[K.status]||"").toLowerCase().includes("delay") && !String(r[K.status]||"").toLowerCase().includes("complete")).length,
                };
                // Type counts — based on status+comp filters only
                const tCounts = allTypes.map(t => ({ val:t, count: teamRows.filter(r => tMatchS(r) && tMatchC(r) && String(r[K.type]||"").trim()===t).length }));
                // Comp counts — based on status+type filters only
                const cCounts = allComponents.map(c => ({ val:c, count: teamRows.filter(r => tMatchS(r) && tMatchT(r) && String(r[K.component]||"").trim()===c).length }));

                return (
                  <>
                    {/* Row 1: Status + Type */}
                    <div style={{ display:"flex", alignItems:"center", gap:12, flexWrap:"wrap" }}>
                      <div style={{ display:"flex", gap:4, alignItems:"center" }}>
                        <span style={{ fontSize:10, color:C.muted, fontWeight:600, marginRight:2 }}>Status</span>
                        {pill("All",      statusFilter==="All",      sCounts.all,     () => setStatusFilter("All"),                                      C.navyLight)}
                        {pill("Delayed",  statusFilter==="Delayed",  sCounts.delayed, () => setStatusFilter(statusFilter==="Delayed"?"All":"Delayed"),    C.delayed)}
                        {pill("On Track", statusFilter==="On Track", sCounts.onTrack, () => setStatusFilter(statusFilter==="On Track"?"All":"On Track"),  C.onTrack)}
                      </div>
                      <div style={{ width:1, height:20, background:C.border }} />
                      <div style={{ display:"flex", gap:4, alignItems:"center", flexWrap:"wrap" }}>
                        <span style={{ fontSize:10, color:C.muted, fontWeight:600, marginRight:2 }}>Type</span>
                        {pill("All", typeFilter==="All", teamRows.filter(r => tMatchS(r) && tMatchC(r)).length, () => setTypeFilter("All"), null)}
                        {tCounts.map(({val,count}) => pill(val, typeFilter===val, count, () => setTypeFilter(typeFilter===val?"All":val), null))}
                      </div>
                      <button onClick={() => setShowColPanel(p => !p)}
                        style={{ marginLeft:"auto", padding:"4px 12px", borderRadius:5, border:`1px solid ${showColPanel?C.navyLight:C.border}`,
                          background:showColPanel?C.navyLight:C.white, color:showColPanel?"#fff":C.muted,
                          cursor:"pointer", fontSize:10, fontWeight:600 }}>⚙ Columns</button>
                    </div>

                    {/* Row 2: Component */}
                    {allComponents.length > 0 && (
                      <div style={{ display:"flex", gap:4, alignItems:"center", flexWrap:"wrap" }}>
                        <span style={{ fontSize:10, color:C.muted, fontWeight:600, marginRight:2 }}>Component</span>
                        {pill("All", compFilter==="All", teamRows.filter(r => tMatchS(r) && tMatchT(r)).length, () => setCompFilter("All"), null)}
                        {cCounts.map(({val,count}) => pill(val, compFilter===val, count, () => setCompFilter(compFilter===val?"All":val), null))}
                      </div>
                    )}
                  </>
                );
              })()}
              {/* Column config panel */}
              {showColPanel && (
                <div style={{ background:"#f8fafc", border:`1px solid ${C.border}`, borderRadius:8, padding:"12px 14px" }}>
                  <div style={{ fontSize:11, fontWeight:700, color:C.text, marginBottom:8 }}>
                    Show / Hide Columns <span style={{ fontSize:10, color:C.muted, fontWeight:400 }}>— drag column edges in the table to resize</span>
                  </div>
                  <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                    {Object.entries(colConfig).map(([key, col]) => (
                      <label key={key} style={{ display:"flex", alignItems:"center", gap:5, background:C.white,
                        border:`1px solid ${col.visible ? C.navyLight : C.border}`,
                        borderRadius:6, padding:"5px 10px", cursor:"pointer", userSelect:"none" }}>
                        <input type="checkbox" checked={col.visible}
                          onChange={e => setColConfig(p => ({...p, [key]: {...p[key], visible: e.target.checked}}))}
                          style={{ cursor:"pointer", width:13, height:13 }} />
                        <span style={{ fontSize:11, color: col.visible ? C.navyLight : C.muted, fontWeight: col.visible ? 700 : 400 }}>{col.label}</span>
                      </label>
                    ))}
                  </div>
                  <button onClick={() => setColConfig(p => Object.fromEntries(Object.entries(p).map(([k,v]) => [k, {...v, visible:true}])))}
                    style={{ marginTop:8, padding:"4px 12px", borderRadius:4, border:`1px solid ${C.border}`, background:C.white, cursor:"pointer", fontSize:10, color:C.muted }}>
                    Show all columns
                  </button>
                </div>
              )}
              <div style={{ fontSize:10, color:C.muted }}>
                Showing <b style={{ color:C.text }}>{filteredRows.length}</b> of <b style={{ color:C.text }}>{teamRows.length}</b> open items for <b style={{ color:C.text }}>{selectedTeam}</b>
              </div>
            </div>

            {/* Table */}
            <div style={{ overflowX:"auto" }}>
              <table style={{ borderCollapse:"collapse", fontSize:11, tableLayout:"fixed",
                width: Object.values(colConfig).filter(c=>c.visible).reduce((s,c)=>s+c.width,0) + "px", minWidth:"100%" }}>
                <thead>
                  <tr style={{ background:"#162f50" }}>
                    {[
                      ["raidId","RAID ID"], ["status","Status"], ["type","Type"], ["component","Component"],
                      ["experience","Experience"], ["topic","Topic"], ["desc","Description"],
                      ["comment","Comments / Resolution"], ["owner","Owner"], ["critPath","Critical Path"], ["dueDate","Due Date"]
                    ].filter(([key]) => colConfig[key].visible).map(([key, label], idx, arr) => (
                      <th key={key} style={{ padding:"8px 10px", textAlign:"left", color:"#fff", fontWeight:700, fontSize:10,
                        width:colConfig[key].width, position:"relative",
                        borderRight: idx < arr.length-1 ? "1px solid rgba(255,255,255,0.1)" : "none" }}>
                        {label}
                        <div
                          onMouseDown={e => {
                            e.preventDefault();
                            const startX = e.clientX;
                            const startW = colConfig[key].width;
                            const onMove = mv => {
                              const newW = Math.max(50, startW + mv.clientX - startX);
                              setColConfig(p => ({...p, [key]: {...p[key], width: newW}}));
                            };
                            const onUp = () => { window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp); };
                            window.addEventListener("mousemove", onMove);
                            window.addEventListener("mouseup", onUp);
                          }}
                          style={{ position:"absolute", right:0, top:0, bottom:0, width:6, cursor:"col-resize",
                            background:"transparent", zIndex:10 }}
                          onMouseEnter={e => e.currentTarget.style.background="rgba(255,255,255,0.3)"}
                          onMouseLeave={e => e.currentTarget.style.background="transparent"}
                        />
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.length === 0 ? (
                    <tr><td colSpan={Object.values(colConfig).filter(c=>c.visible).length} style={{ padding:"20px", textAlign:"center", color:C.muted }}>No items match current filters</td></tr>
                  ) : filteredRows.map((r, i) => {
                    const status = String(r[K.status]||"").trim();
                    const sCol = statusCol(status);
                    const due = daysUntil(r[K.date]);
                    const dueStr = fmtDate(r[K.date]);
                    const dueCol = due != null && due <= 7 ? C.delayed : due != null && due <= 14 ? C.gold : C.muted;
                    return (
                      <tr key={i} style={{ background:i%2===0?C.white:"#f7f9fc", borderBottom:`1px solid ${C.border}`, verticalAlign:"top" }}>
                        {colConfig.raidId.visible    && <td style={{ padding:"8px 10px", fontWeight:700, color:C.navyLight, wordBreak:"break-word", width:colConfig.raidId.width }}>{String(r[K.id]||"—")}</td>}
                        {colConfig.status.visible    && <td style={{ padding:"8px 10px", width:colConfig.status.width }}>
                          <span style={{ background:sCol+"20", color:sCol, border:`1px solid ${sCol}40`, borderRadius:4, padding:"2px 6px", fontSize:10, fontWeight:700, whiteSpace:"nowrap" }}>{status||"—"}</span>
                        </td>}
                        {colConfig.type.visible      && <td style={{ padding:"8px 10px", color:C.text, wordBreak:"break-word", width:colConfig.type.width }}>{String(r[K.type]||"—")}</td>}
                        {colConfig.component.visible && <td style={{ padding:"8px 10px", color:C.text, wordBreak:"break-word", width:colConfig.component.width }}>{String(r[K.component]||"—")}</td>}
                        {colConfig.experience.visible&& <td style={{ padding:"8px 10px", color:C.muted, wordBreak:"break-word", width:colConfig.experience.width }}>{String(r[K.experience]||"—")}</td>}
                        {colConfig.topic.visible     && <td style={{ padding:"8px 10px", color:C.muted, wordBreak:"break-word", width:colConfig.topic.width }}>{String(r[K.topic]||"—")}</td>}
                        {colConfig.desc.visible      && <td style={{ padding:"8px 10px", color:C.text, wordBreak:"break-word", lineHeight:1.5, width:colConfig.desc.width }}>{String(r[K.desc]||"—")}</td>}
                        {colConfig.comment.visible   && <td style={{ padding:"8px 10px", color:C.muted, wordBreak:"break-word", lineHeight:1.5, width:colConfig.comment.width }}>{String(r[K.comment]||"—")}</td>}
                        {colConfig.owner.visible     && <td style={{ padding:"8px 10px", color:C.text, wordBreak:"break-word", width:colConfig.owner.width }}>{String(r[K.owner]||"—")}</td>}
                        {colConfig.critPath.visible  && <td style={{ padding:"8px 10px", width:colConfig.critPath.width }}>
                          {(() => {
                            const v = String(r[K.critPath]||"").trim();
                            if (!v || v === "—") return <span style={{ color:C.muted }}>—</span>;
                            const isHighlight = v.toLowerCase() !== "no" && v.toLowerCase() !== "n/a";
                            return <span style={{ background: isHighlight ? "#fee2e2" : "#f1f5f9", color: isHighlight ? C.delayed : C.muted, borderRadius:3, padding:"2px 6px", fontSize:10, fontWeight:600 }}>{v}</span>;
                          })()}
                        </td>}
                        {colConfig.dueDate.visible   && <td style={{ padding:"8px 10px", color:dueCol, fontWeight:600, whiteSpace:"nowrap", width:colConfig.dueDate.width }}>{dueStr}</td>}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </>
        )}

        {!selectedTeam && (
          <div style={{ padding:"24px", textAlign:"center", color:C.muted, fontSize:12 }}>
            Select a team above to view their open RAID items
          </div>
        )}
      </Card>

      {raidModal && (() => {
        const resolvedTeamKey = teamKey || K.team || "Primary Team (Owner)";
        const allModalTeams = Array.from(new Set(raidModal.rows.map(r => String(r[resolvedTeamKey]||"").trim()).filter(Boolean))).sort();
        const allModalTypes = Array.from(new Set(raidModal.rows.map(r => String(r[K.type]||"").trim()).filter(Boolean))).sort();
        const allModalComps = Array.from(new Set(raidModal.rows.map(r => String(r[K.component]||"").trim()).filter(Boolean))).sort();
        return (
          <RaidKpiModal
            title={raidModal.title}
            rows={raidModal.rows}
            K={K} teamKey={resolvedTeamKey}
            allTeams={allModalTeams} allTypes={allModalTypes} allComps={allModalComps}
            statusCol={statusCol}
            hideType={raidModal.hideType || false}
            hideStatus={raidModal.hideStatus || false}
            colConfig={modalColConfig}
            setColConfig={setModalColConfig}
            onClose={() => setRaidModal(null)}
          />
        );
      })()}
    </div>
  );
}

// ─── OVERVIEW ────────────────────────────────────────────────────────────────

// ── RAID priority bar (inline horizontal) ────────────────────────────────────

// RAID drill-down modal with team filter + status + type filters
function RaidDrillModal({ title, rows, raidKeys, onClose, initialStatusFilter, initialTypeFilter }) {
  const [expanded, setExpanded] = useState({});
  const [statusFilter, setStatusFilter] = useState(initialStatusFilter || "All");
  const [typeFilter, setTypeFilter] = useState(initialTypeFilter || "All");
  if (!rows?.length) return null;
  const K = raidKeys || {};

  const teamCol    = K.team      || "Primary Team";
  const ownerCol   = K.owner     || "Primary Owner";
  const statusCol  = K.status    || "Status";
  const priorityCol= K.priority  || "Priority";
  const typeCol    = K.type      || "Type";
  const descCol    = K.desc      || "Description";
  const commentCol = K.comment   || "Comments/ Resolution History";
  const idCol      = K.id        || "RAID ID";
  const compCol    = K.component || "Component";
  const expCol     = K.experience|| "Experience";
  const topicCol   = K.topic     || "Topic";
  const critCol    = K.critPath  || "Critical Path";
  const dateCol    = K.date      || "Due Date";

  // Derive available types from data
  const allTypes = Array.from(new Set(rows.map(r => String(r[typeCol] || "").trim()).filter(Boolean))).sort();
  // Canonical type filter labels — map to keywords
  const TYPE_FILTERS = [
    { label: "All",       match: null },
    { label: "Risk",      match: "risk" },
    { label: "Issue",     match: "issue" },
    { label: "Action",    match: "action" },
    { label: "Decision",  match: "decision" },
  ];
  // Status filter labels
  const STATUS_FILTERS = [
    { label: "All",     match: null },
    { label: "Delayed", match: "delayed" },
    { label: "On Track",match: "on track" },
  ];

  // Apply filters
  const filteredRows = rows.filter(r => {
    const s = String(r[statusCol] || "").toLowerCase();
    const t = String(r[typeCol]   || "").toLowerCase();
    const statusOk = statusFilter === "All" || s.includes(STATUS_FILTERS.find(f => f.label === statusFilter)?.match || "");
    const typeOk   = typeFilter   === "All" || t.includes(TYPE_FILTERS.find(f => f.label === typeFilter)?.match || "");
    return statusOk && typeOk;
  });

  // Count badges for each filter
  const statusCounts = STATUS_FILTERS.reduce((acc, f) => {
    acc[f.label] = f.match ? rows.filter(r => String(r[statusCol]||"").toLowerCase().includes(f.match)).length : rows.length;
    return acc;
  }, {});
  const typeCounts = TYPE_FILTERS.reduce((acc, f) => {
    acc[f.label] = f.match ? rows.filter(r => String(r[typeCol]||"").toLowerCase().includes(f.match)).length : rows.length;
    return acc;
  }, {});

  // Group filtered rows by Primary Owner Team, sort each group by Due Date asc
  const groups = {};
  filteredRows.forEach(r => {
    const c = String(r[teamCol] || "Unknown");
    if (!groups[c]) groups[c] = [];
    groups[c].push(r);
  });
  const parseDue = r => {
    const v = r[dateCol];
    if (!v || v === "" || v === "—") return Infinity;
    try { return typeof v === "number" ? (v - 25569) * 86400000 : new Date(v).getTime(); }
    catch { return Infinity; }
  };
  Object.values(groups).forEach(arr => arr.sort((a, b) => parseDue(a) - parseDue(b)));
  const sortedGroups = Object.entries(groups).sort((a, b) => a[0].localeCompare(b[0]));

  const isOpen     = key => expanded[key] === true;
  const toggleGroup = c => setExpanded(prev => ({ ...prev, [c]: !isOpen(c) }));
  const expandAll   = () => { const e = {}; sortedGroups.forEach(([c]) => e[c] = true);  setExpanded(e); };
  const collapseAll = () => { const e = {}; sortedGroups.forEach(([c]) => e[c] = false); setExpanded(e); };

  const cols = [idCol, priorityCol, statusCol, expCol, compCol, topicCol, descCol, commentCol, ownerCol, critCol, dateCol];
  const wideCols = new Set([descCol, commentCol]);

  const FilterPills = ({ filters, counts, active, onSelect, delayedHighlight }) => (
    <div style={{ display:"flex", gap:5, flexWrap:"wrap" }}>
      {filters.map(({ label }) => {
        const isActive = active === label;
        const isDelayed = delayedHighlight && label === "Delayed";
        const activeBg = isDelayed ? C.delayed + "cc" : label === "All" ? "rgba(255,255,255,0.2)" : "rgba(255,255,255,0.25)";
        return (
          <button key={label} onClick={() => { onSelect(label); setExpanded({}); }}
            style={{ display:"flex", alignItems:"center", gap:5, padding:"3px 10px",
              borderRadius:20, border:`1px solid ${isActive ? "transparent" : "rgba(255,255,255,0.3)"}`,
              background: isActive ? activeBg : "transparent",
              color: isActive ? "#fff" : "rgba(255,255,255,0.65)",
              cursor:"pointer", fontSize:11, fontWeight: isActive ? 700 : 500, transition:"all .12s" }}>
            {label}
            <span style={{ background:"rgba(0,0,0,0.2)", borderRadius:8, padding:"1px 5px", fontSize:10 }}>
              {counts[label] ?? 0}
            </span>
          </button>
        );
      })}
    </div>
  );

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:1000, display:"flex", alignItems:"center", justifyContent:"center" }}
      onClick={onClose}>
      <div style={{ background:C.white, borderRadius:10, width:"98%", maxWidth:1350, maxHeight:"90vh",
        display:"flex", flexDirection:"column", boxShadow:"0 24px 60px rgba(0,0,0,0.35)" }}
        onClick={e => e.stopPropagation()}>

        {/* Header */}
        <div style={{ background:C.headerBg, padding:"12px 20px", borderRadius:"10px 10px 0 0", flexShrink:0, display:"flex", flexDirection:"column", gap:10 }}>
          {/* Title row */}
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:12, flexWrap:"wrap" }}>
            <div style={{ color:"#fff", fontWeight:700, fontSize:13 }}>
              {title}
              <span style={{ opacity:.6, fontWeight:400, marginLeft:8 }}>
                ({filteredRows.length} items · {sortedGroups.length} teams)
              </span>
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <button onClick={expandAll}
                style={{ padding:"4px 12px", borderRadius:4, border:"1px solid rgba(255,255,255,0.4)", background:"rgba(255,255,255,0.15)", color:"#fff", fontSize:11, cursor:"pointer", fontWeight:600 }}>
                ⊞ Expand All
              </button>
              <button onClick={collapseAll}
                style={{ padding:"4px 12px", borderRadius:4, border:"1px solid rgba(255,255,255,0.4)", background:"rgba(255,255,255,0.15)", color:"#fff", fontSize:11, cursor:"pointer", fontWeight:600 }}>
                ⊟ Collapse All
              </button>
              <button onClick={onClose} style={{ background:"rgba(255,255,255,0.15)", border:"none", color:"#fff", borderRadius:5, padding:"5px 14px", cursor:"pointer", fontSize:13, fontWeight:600 }}>✕</button>
            </div>
          </div>
          {/* Filter rows */}
          <div style={{ display:"flex", gap:16, flexWrap:"wrap", alignItems:"center" }}>
            <div style={{ display:"flex", alignItems:"center", gap:7 }}>
              <span style={{ color:"rgba(255,255,255,0.5)", fontSize:10, fontWeight:600, textTransform:"uppercase", letterSpacing:"0.06em" }}>Status</span>
              <FilterPills filters={STATUS_FILTERS} counts={statusCounts} active={statusFilter} onSelect={setStatusFilter} delayedHighlight />
            </div>
            <div style={{ width:"1px", height:20, background:"rgba(255,255,255,0.2)" }} />
            <div style={{ display:"flex", alignItems:"center", gap:7 }}>
              <span style={{ color:"rgba(255,255,255,0.5)", fontSize:10, fontWeight:600, textTransform:"uppercase", letterSpacing:"0.06em" }}>Type</span>
              <FilterPills filters={TYPE_FILTERS} counts={typeCounts} active={typeFilter} onSelect={setTypeFilter} />
            </div>
          </div>
        </div>

        {/* Body */}
        <div style={{ overflowY:"auto", overflowX:"auto", flex:1 }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
            <thead style={{ position:"sticky", top:0, background:"#f0f4f8", zIndex:2 }}>
              <tr>
                <th style={{ textAlign:"left", padding:"8px 10px", color:C.muted, fontWeight:700, borderBottom:`2px solid ${C.border}`, whiteSpace:"nowrap", minWidth:160 }}>Primary Owner Team</th>
                {cols.map(c => (
                  <th key={c} style={{ textAlign:"left", padding:"8px 10px", color:C.muted, fontWeight:700,
                    borderBottom:`2px solid ${C.border}`, whiteSpace:"nowrap",
                    minWidth: wideCols.has(c) ? 260 : c === dateCol ? 95 : 100 }}>{c}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sortedGroups.map(([team, teamRows]) => {
                const open = isOpen(team);
                const hasDelayed = teamRows.some(r => String(r[statusCol]||"").toLowerCase() === "delayed");
                return <React.Fragment key={team}>
                  <tr onClick={() => toggleGroup(team)}
                    style={{ background:"#e2eaf6", cursor:"pointer", borderBottom:`1px solid ${C.border}` }}>
                    <td colSpan={cols.length + 1} style={{ padding:"9px 14px" }}>
                      <span style={{ display:"flex", alignItems:"center", gap:10 }}>
                        <span style={{ fontSize:14, userSelect:"none", lineHeight:1 }}>{open ? "▼" : "▶"}</span>
                        <span style={{ fontWeight:700, color:C.navy, fontSize:12 }}>{team}</span>
                        <span style={{ background:C.navyLight+"25", color:C.navyLight, borderRadius:10, padding:"1px 9px", fontSize:10, fontWeight:700 }}>
                          {teamRows.length} items
                        </span>
                        {hasDelayed && (
                          <span style={{ background:C.delayed+"20", color:C.delayed, borderRadius:10, padding:"1px 9px", fontSize:10, fontWeight:700 }}>
                            ⚠ Has Delayed
                          </span>
                        )}
                        <span style={{ color:C.muted, fontSize:10, marginLeft:"auto" }}>
                          {open ? "Click to collapse" : "Click to expand"}
                        </span>
                      </span>
                    </td>
                  </tr>,
                  {open && teamRows.map((r, i) => {
                    const daysLeft = daysUntil(r[dateCol]);
                    const dueColor = daysLeft != null && daysLeft < 0 ? C.delayed : daysLeft != null && daysLeft <= 7 ? C.yellow : C.text;
                    const statusVal = String(r[statusCol] || "—");
                    const statusColor = statusVal.toLowerCase() === "delayed" ? C.delayed : statusVal.toLowerCase() === "complete" ? C.complete : C.onTrack;
                    return (
                      <tr key={`row-${team}-${i}`}
                        style={{ background: i % 2 === 0 ? "#f9fbff" : C.white, borderBottom:`1px solid ${C.border}`, verticalAlign:"top" }}>
                        <td style={{ padding:"7px 10px 7px 30px", color:C.muted, fontSize:10, whiteSpace:"nowrap" }}>↳</td>
                        {cols.map(c => {
                          const v = String(r[c] || "—");
                          const isPri  = c === priorityCol;
                          const isStat = c === statusCol;
                          const isWide = wideCols.has(c);
                          const isDate = c === dateCol;
                          const priColor = isPri ? getPriorityColor(v) : null;
                          return (
                            <td key={c} style={{
                              padding:"7px 10px",
                              color: isDate ? dueColor : C.text,
                              whiteSpace: isWide ? "pre-wrap" : "nowrap",
                              maxWidth: isWide ? 300 : isDate ? 100 : 180,
                              wordBreak: isWide ? "break-word" : "normal",
                              overflow: isWide ? "visible" : "hidden",
                              textOverflow: isWide ? "unset" : "ellipsis",
                              fontSize:11, lineHeight: isWide ? 1.5 : "normal",
                              fontWeight: isDate && daysLeft != null && daysLeft < 0 ? 700 : "normal"
                            }} title={isWide ? undefined : v}>
                              {isPri && priColor
                                ? <span style={{ background:priColor+"20", color:priColor, border:`1px solid ${priColor}40`, borderRadius:3, padding:"2px 7px", fontSize:10, fontWeight:700 }}>{v}</span>
                                : isStat
                                  ? <span style={{ background:statusColor+"20", color:statusColor, border:`1px solid ${statusColor}40`, borderRadius:3, padding:"2px 7px", fontSize:10, fontWeight:700 }}>{v}</span>
                                  : isDate && daysLeft != null
                                    ? <>{fmtDate(r[c])}<div style={{ fontSize:9, color:dueColor, marginTop:1 }}>{daysLeft < 0 ? `${Math.abs(daysLeft)}d overdue` : daysLeft === 0 ? "today" : `${daysLeft}d left`}</div></>
                                    : isWide ? v : v.slice(0, 60)}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </React.Fragment>;
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}


function CommentCell({ value }) {
  const raw = String(value ?? "").trim();
  // Treat pandas NaN exports and empty values as blank
  const EMPTY = ["", "nan", "NaN", "null", "undefined", "0", "-", "—"];
  if (EMPTY.includes(raw)) return <span style={{ color:"#ccc" }}>—</span>;
  const lines = raw.split(/\n/);
  return (
    <>{lines.map((line, i) => (
      <div key={i} style={{ marginBottom: line.trim() ? 2 : 4 }}>{line || " "}</div>
    ))}</>
  );
}
function WorkplanDrillModal({ title, rows, onClose, initialFilter }) {
  const getS   = r => String(r["Default Status"] || r["Status"] || "");
  const getLvl = r => Number(r["Lvl"] ?? 99);
  const isLeafR = r => { const c = r["Children"]; if (c === null || c === undefined || c === "" || c === "0") return true; const n = Number(c); return isNaN(n) || n === 0; };

  // Pre-compute which group keys have off-track descendants — expand those by default
  const initExpanded = () => {
    const exp = {};
    rows.forEach((r, i) => {
      if (isLeafR(r)) return;
      const grpLvl = getLvl(r);
      for (let j = i + 1; j < rows.length; j++) {
        if (getLvl(rows[j]) <= grpLvl) break;
        if (getS(rows[j]).toLowerCase().includes("off track")) { exp[`node_${i}`] = true; break; }
      }
    });
    return exp;
  };

  const [expanded, setExpanded] = useState(() => initExpanded());
  const [statusFilter, setStatusFilter] = useState(initialFilter || "All");
  if (!rows?.length) return null;

  // When "All" → show everything. When "Complete" → show only complete. Otherwise exclude Complete.
  const nonComplete = rows.filter(r => getS(r).toLowerCase() !== "complete");
  const activeRows = statusFilter === "All"
    ? rows
    : statusFilter === "Complete"
      ? rows   // Complete filter uses full rows then leafMatches filters
      : (nonComplete.length > 0 ? nonComplete : rows);

  // Apply status filter — only keep leaf rows matching the filter,
  // plus ancestor group rows that have at least one matching leaf descendant.
  const isLeafCheck = r => { const c = r["Children"]; if (c === null || c === undefined || c === "" || c === "0") return true; const n = Number(c); return isNaN(n) || n === 0; };

  const leafMatches = statusFilter === "All"
    ? new Set(activeRows.map((_, i) => i))
    : new Set(
        activeRows.reduce((acc, r, i) => {
          if (!isLeafCheck(r)) return acc;
          const s = getS(r).toLowerCase();
          if (statusFilter === "Off Track"   && s.includes("off track"))  acc.push(i);
          if (statusFilter === "On Track"    && s.includes("on track"))   acc.push(i);
          if (statusFilter === "Not Started" && s.includes("not start"))  acc.push(i);
          if (statusFilter === "Complete"    && s.includes("complete"))   acc.push(i);
          return acc;
        }, [])
      );

  // For each group row at index i, check if any row at index j>i with lvl > group.lvl is in leafMatches
  const groupHasMatch = (groupIdx) => {
    const groupLvl = getLvl(activeRows[groupIdx]);
    for (let j = groupIdx + 1; j < activeRows.length; j++) {
      if (getLvl(activeRows[j]) <= groupLvl) break; // left the subtree
      if (leafMatches.has(j)) return true;
    }
    return false;
  };

  const filteredRows = statusFilter === "All" ? activeRows : activeRows.filter((r, i) => {
    if (isLeafCheck(r)) return leafMatches.has(i);
    return groupHasMatch(i);
  });

  // Status counts for filter badges — always count from full rows for accuracy
  const allLeaves = rows.filter(r => { const c = r["Children"]; if (c === null || c === undefined || c === "" || c === "0") return true; const n = Number(c); return isNaN(n) || n === 0; });
  const counts = {
    "All":         allLeaves.length,
    "Off Track":   allLeaves.filter(r => getS(r).toLowerCase().includes("off track")).length,
    "On Track":    allLeaves.filter(r => getS(r).toLowerCase().includes("on track")).length,
    "Not Started": allLeaves.filter(r => getS(r).toLowerCase().includes("not start")).length,
    "Complete":    allLeaves.filter(r => getS(r).toLowerCase().includes("complete")).length,
  };

  // isLeaf: handle string "0", "", null, undefined from Excel exports
  const isLeaf = r => {
    const c = r["Children"];
    if (c === null || c === undefined || c === "" || c === "0") return true;
    const n = Number(c);
    return isNaN(n) || n === 0;
  };

  // Track the minimum Lvl as the root level — guard against empty
  const lvlValues = filteredRows.map(getLvl).filter(l => !isNaN(l) && l !== 99);
  const minLvl = lvlValues.length > 0 ? Math.min(...lvlValues) : 1;

  const isDelayedStatus = s => s.toLowerCase() === "off track" || s.toLowerCase() === "delayed";

  const nodes = filteredRows.map((r, i) => ({
    r, i,
    key: `node_${i}`,
    lvl: getLvl(r),
    depth: Math.max(0, getLvl(r) - minLvl),
    isGroup: !isLeaf(r),
    children: Number(r["Children"] || 0),
    delayedCount: 0,
    totalCount: 0,
  }));

  // Determine parent key for each node so we can check if parent is collapsed
  // A node's parent is the nearest preceding node with (lvl == this.lvl - 1)
  const parentKey = {};
  nodes.forEach((node, i) => {
    for (let j = i - 1; j >= 0; j--) {
      if (nodes[j].lvl === node.lvl - 1) {
        parentKey[node.key] = nodes[j].key;
        break;
      }
    }
  });

  // For each group node, compute how many of its direct+indirect descendants are delayed
  // A descendant of node[i] = any node[j] where j>i and nodes[j].lvl > nodes[i].lvl
  // until we hit a node with lvl <= nodes[i].lvl
  nodes.forEach((node, i) => {
    if (!node.isGroup) return;
    let delayed = 0, total = 0;
    for (let j = i + 1; j < nodes.length; j++) {
      if (nodes[j].lvl <= node.lvl) break; // out of subtree
      if (!nodes[j].isGroup) {             // only count leaf tasks
        total++;
        if (isDelayedStatus(getS(nodes[j].r))) delayed++;
      }
    }
    node.delayedCount = delayed;
    node.totalCount   = total;
  });

  // isVisible: node is visible if all ancestors are expanded
  const isOpen    = key => expanded[key] === true;
  const isVisible = node => {
    let pk = parentKey[node.key];
    while (pk) {
      if (!isOpen(pk)) return false;
      pk = parentKey[pk];
    }
    return true;
  };

  const toggle = key => setExpanded(p => ({ ...p, [key]: !isOpen(key) }));

  const expandAll = () => {
    const e = {};
    nodes.filter(n => n.isGroup).forEach(n => e[n.key] = true);
    setExpanded(e);
  };
  const collapseAll = () => setExpanded({});

  const visibleNodes = nodes.filter(isVisible);

  const INDENT = 18; // px per depth level

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:1000, display:"flex", alignItems:"center", justifyContent:"center" }}
      onClick={onClose}>
      <div style={{ background:C.white, borderRadius:10, width:"98%", maxWidth:1300, maxHeight:"90vh",
        display:"flex", flexDirection:"column", boxShadow:"0 24px 60px rgba(0,0,0,0.35)" }}
        onClick={e => e.stopPropagation()}>

        {/* Header */}
        <div style={{ background:C.headerBg, padding:"12px 20px", borderRadius:"10px 10px 0 0",
          display:"flex", flexDirection:"column", gap:10, flexShrink:0 }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:12, flexWrap:"wrap" }}>
            <div style={{ color:"#fff", fontWeight:700, fontSize:13 }}>
              {title.replace("Technology - ","")}
              <span style={{ opacity:.6, fontWeight:400, marginLeft:8 }}>
                ({filteredRows.filter(r => isLeaf(r)).length} tasks{statusFilter !== "All" ? ` — ${statusFilter}` : ""})
              </span>
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <button onClick={expandAll}
                style={{ padding:"4px 12px", borderRadius:4, border:"1px solid rgba(255,255,255,0.4)", background:"rgba(255,255,255,0.15)", color:"#fff", fontSize:11, cursor:"pointer", fontWeight:600 }}>
                ⊞ Expand All
              </button>
              <button onClick={collapseAll}
                style={{ padding:"4px 12px", borderRadius:4, border:"1px solid rgba(255,255,255,0.4)", background:"rgba(255,255,255,0.15)", color:"#fff", fontSize:11, cursor:"pointer", fontWeight:600 }}>
                ⊟ Collapse All
              </button>
              <button onClick={onClose} style={{ background:"rgba(255,255,255,0.15)", border:"none", color:"#fff", borderRadius:5, padding:"5px 14px", cursor:"pointer", fontSize:13, fontWeight:600 }}>✕</button>
            </div>
          </div>
          {/* Status filter pills */}
          <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
            {[
              ["All",         "rgba(255,255,255,0.2)", "#fff"],
              ["Off Track",   C.delayed + "cc",        "#fff"],
              ["On Track",    C.onTrack + "cc",        "#fff"],
              ["Not Started", "rgba(255,255,255,0.12)","rgba(255,255,255,0.7)"],
              ["Complete",    C.complete + "cc",       "#fff"],
            ].map(([label, activeBg, activeText]) => {
              const active = statusFilter === label;
              const count  = counts[label] ?? 0;
              return (
                <button key={label} onClick={() => { setStatusFilter(label); setExpanded({}); }}
                  style={{ display:"flex", alignItems:"center", gap:5, padding:"4px 10px",
                    borderRadius:20, border:`1px solid ${active ? "transparent" : "rgba(255,255,255,0.3)"}`,
                    background: active ? activeBg : "transparent",
                    color: active ? activeText : "rgba(255,255,255,0.65)",
                    cursor:"pointer", fontSize:11, fontWeight: active ? 700 : 500, transition:"all .12s" }}>
                  {label}
                  <span style={{ background:"rgba(0,0,0,0.2)", borderRadius:8, padding:"1px 5px", fontSize:10 }}>
                    {count}
                  </span>
                </button>
              );
            })}
          </div>
        </div>

        {/* Table */}
        <div style={{ overflowY:"auto", overflowX:"auto", flex:1 }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
            <thead style={{ position:"sticky", top:0, background:"#f0f4f8", zIndex:2 }}>
              <tr style={{ borderBottom:`2px solid ${C.border}` }}>
                <th style={{ textAlign:"left", padding:"8px 10px", color:C.muted, fontWeight:700, minWidth:320 }}>Task / Group</th>
                <th style={{ textAlign:"center", padding:"8px 10px", color:C.muted, fontWeight:700, whiteSpace:"nowrap" }}>Status</th>
                <th style={{ textAlign:"center", padding:"8px 10px", color:C.muted, fontWeight:700, whiteSpace:"nowrap" }}>% Done</th>
                <th style={{ textAlign:"center", padding:"8px 10px", color:C.muted, fontWeight:700, whiteSpace:"nowrap" }}>Start</th>
                <th style={{ textAlign:"center", padding:"8px 10px", color:C.muted, fontWeight:700, whiteSpace:"nowrap" }}>Finish</th>
                <th style={{ textAlign:"left",   padding:"8px 10px", color:C.muted, fontWeight:700, whiteSpace:"nowrap" }}>Workstream</th>
                <th style={{ textAlign:"left",   padding:"8px 10px", color:C.muted, fontWeight:700, whiteSpace:"nowrap" }}>Support</th>
                <th style={{ textAlign:"left",   padding:"8px 10px", color:C.muted, fontWeight:700, whiteSpace:"nowrap" }}>Primary Owner</th>
                <th style={{ textAlign:"left",   padding:"8px 10px", color:C.muted, fontWeight:700, whiteSpace:"nowrap" }}>Secondary Owner</th>
                <th style={{ textAlign:"left",   padding:"8px 10px", color:C.muted, fontWeight:700, minWidth:200 }}>Comments</th>
              </tr>
            </thead>
            <tbody>
              {visibleNodes.map(node => {
                const r = node.r;
                const s = getS(r);
                const sc = SC[s] || C.muted;
                const p2 = pct(r["% Complete"] ?? r["% complete"]);
                const daysLeft = daysUntil(r["Finish"] || r["End Date"]);
                const isComplete = s.toLowerCase() === "complete";
                const dueColor = isComplete ? C.muted : daysLeft != null && daysLeft < 0 ? C.delayed : daysLeft != null && daysLeft <= 7 ? C.yellow : C.muted;
                const indent = node.depth * INDENT;

                // Group header row (Children > 0)
                if (node.isGroup) {
                  const open = isOpen(node.key);
                  const bgColors = ["#d4e4f7","#dcedf9","#e4f2fb","#ecf5fd","#f2f8fe"];
                  const bg = node.delayedCount > 0 ? 
                    ["#f7dede","#f9e4e4","#faeaea","#fbeeee","#fdf5f5"][Math.min(node.depth, 4)] :
                    bgColors[Math.min(node.depth, bgColors.length - 1)];
                  const hasDelayedChildren = node.delayedCount > 0;
                  // Derived status: if own status is On Track but has delayed children → show both
                  const ownIsOnTrack = s.toLowerCase() === "on track" || s === "";
                  return (
                    <tr key={node.key} onClick={() => toggle(node.key)}
                      style={{ background:bg, cursor:"pointer", borderBottom:`1px solid ${C.border}` }}>
                      <td colSpan={10} style={{ padding:`8px 10px 8px ${10 + indent}px` }}>
                        <span style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
                          <span style={{ fontSize:11, userSelect:"none", width:14, flexShrink:0 }}>{open ? "▼" : "▶"}</span>
                          <span style={{ fontWeight:700, color: hasDelayedChildren ? C.navy : C.navy, fontSize: Math.max(10, 13 - node.depth) }}>
                            {String(r["Task Name"] || "—")}
                          </span>
                          <span style={{ color:C.muted, fontSize:10 }}>{node.totalCount} tasks</span>
                          {/* Own status badge */}
                          {s && (
                            <span style={{ background:sc+"20", color:sc, border:`1px solid ${sc}40`, borderRadius:3, padding:"1px 6px", fontSize:10, fontWeight:700 }}>
                              {s}
                            </span>
                          )}
                          {/* Derived delayed indicator — shown even if own status is On Track */}
                          {hasDelayedChildren && (
                            <span style={{ background:"#c0392b15", color:C.delayed, border:`1px solid ${C.delayed}50`,
                              borderRadius:3, padding:"1px 7px", fontSize:10, fontWeight:700,
                              display:"flex", alignItems:"center", gap:4 }}>
                              ⚠ {node.delayedCount} delayed task{node.delayedCount > 1 ? "s" : ""}
                            </span>
                          )}
                        </span>
                      </td>
                    </tr>
                  );
                }

                // Leaf task row
                return (
                  <tr key={node.key}
                    style={{ background: node.i % 2 === 0 ? C.white : "#f9fbff", borderBottom:`1px solid ${C.border}30`, verticalAlign:"top" }}>
                    <td style={{ padding:`6px 10px 6px ${10 + indent}px`, maxWidth:340 }}>
                      <div style={{ display:"flex", alignItems:"flex-start", gap:5 }}>
                        <span style={{ color:C.muted, fontSize:10, flexShrink:0, marginTop:1 }}>↳</span>
                        <span style={{ color:C.text, fontSize:11, lineHeight:1.4 }} title={r["Task Name"]}>
                          {String(r["Task Name"] || "—").slice(0, 90)}
                        </span>
                      </div>
                    </td>
                    <td style={{ padding:"6px 8px", textAlign:"center", whiteSpace:"nowrap" }}>
                      <span style={{ background:sc+"20", color:sc, border:`1px solid ${sc}40`, borderRadius:3, padding:"1px 6px", fontSize:10, fontWeight:700 }}>{s || "—"}</span>
                    </td>
                    <td style={{ padding:"6px 8px", textAlign:"center", whiteSpace:"nowrap" }}>
                      <div style={{ display:"flex", alignItems:"center", gap:4, justifyContent:"center" }}>
                        <div style={{ width:36, background:"#e2e8f0", borderRadius:3, height:5, overflow:"hidden" }}>
                          <div style={{ width:`${p2||0}%`, height:"100%", background:p2>=75?C.green:p2>=40?C.gold:C.delayed }} />
                        </div>
                        <span style={{ fontSize:10, fontWeight:600 }}>{p2 != null ? `${p2}%` : "—"}</span>
                      </div>
                    </td>
                    <td style={{ padding:"6px 8px", textAlign:"center", whiteSpace:"nowrap", color:C.muted, fontSize:11 }}>
                      {fmtDate(r["Start"])}
                    </td>
                    <td style={{ padding:"6px 8px", textAlign:"center", whiteSpace:"nowrap", color:dueColor,
                      fontWeight: !isComplete && daysLeft != null && daysLeft < 0 ? 700 : "normal", fontSize:11 }}>
                      {fmtDate(r["Finish"] || r["End Date"])}
                      {!isComplete && daysLeft != null && (
                        <div style={{ fontSize:9, marginTop:1 }}>
                          {daysLeft < 0 ? `${Math.abs(daysLeft)}d overdue` : daysLeft === 0 ? "today" : `${daysLeft}d left`}
                        </div>
                      )}
                    </td>
                    <td style={{ padding:"6px 8px", color:C.muted, fontSize:10, whiteSpace:"nowrap" }}>{String(r["Workstream"]||"—").slice(0,18)}</td>
                    <td style={{ padding:"6px 8px", color:C.muted, fontSize:10, whiteSpace:"nowrap" }}>{String(r["Support"]||"—").slice(0,18)}</td>
                    <td style={{ padding:"6px 8px", color:C.text,  fontSize:10, whiteSpace:"nowrap" }}>{String(r["Primary Owner"]||"—").slice(0,18)}</td>
                    <td style={{ padding:"6px 8px", color:C.muted, fontSize:10, whiteSpace:"nowrap" }}>{String(r["Secondary Owner"]||"—").slice(0,18)}</td>
                    <td style={{ padding:"6px 8px", color:C.muted, fontSize:10, maxWidth:260, lineHeight:1.5, verticalAlign:"top" }}>
                      <CommentCell value={r["Comments"] || r["Comment"] || r["comments"] || ""} />
                    </td>                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function OverviewTab({ wp, raid, req, cap, openModal }) {
  const [raidModal, setRaidModal] = useState(null);       // { title, rows }
  const [wpModal, setWpModal] = useState(null);           // { title, rows }
  const [wpStatusFilter, setWpStatusFilter] = useState("All");

  // ── compute per-component scorecard ──────────────────────────────────────
  const components = wp ? Object.keys(wp.componentMap) : [];

  const getCompRaid = (compName) => {
    if (!raid) return { open:0, delayed:0, issues:[], risks:[], openItems:[] };
    const comp = compName.replace("Technology - ","").replace("Technology","").toLowerCase().trim();
    const items = raid.items.filter(r => {
      const c = String(r[raid.keys.component]||"").toLowerCase().trim();
      if (!c) return false;
      // Match if component contains comp name or vice versa (first meaningful word)
      const compWords = comp.split(/[\s\-\/]+/).filter(w => w.length > 2);
      return compWords.some(w => c.includes(w)) || c.split(/[\s\-\/]+/).filter(w=>w.length>2).some(w => comp.includes(w));
    });
    const open    = items.filter(r => { const s = String(r[raid.keys.status]||"").toLowerCase(); return s !== "complete" && s !== "deferred"; });
    const delayed = items.filter(r => String(r[raid.keys.status]||"").toLowerCase() === "delayed");
    const issues  = open.filter(r => String(r[raid.keys.type]||"").toLowerCase().includes("issue"));
    const risks   = open.filter(r => String(r[raid.keys.type]||"").toLowerCase().includes("risk"));
    return { open: open.length, delayed: delayed.length, issues, risks, openItems: open };
  };

  const getCompReq = (compName) => {
    if (!req) return null;
    // Exact match first on Sub Process column, then fuzzy
    const exactMatch = req.compBySprint && req.compBySprint[compName];
    const compData   = req.byComponent && (req.byComponent[compName] || null);

    // If no exact match, try partial
    let resolvedKey = null;
    if (exactMatch) {
      resolvedKey = compName;
    } else if (req.byComponent) {
      const cn = compName.toLowerCase().trim();
      resolvedKey = Object.keys(req.byComponent).find(k => {
        const kl = k.toLowerCase().trim();
        return kl === cn || kl.includes(cn) || cn.includes(kl);
      }) || null;
    }
    if (!resolvedKey) return null;

    const sprintData = req.compBySprint[resolvedKey] || {};
    const sprintEntries = Object.entries(sprintData)
      .sort((a,b) => String(a[0]).localeCompare(String(b[0])));

    const cd = req.byComponent[resolvedKey];
    const bs = req.compBuildStatus && req.compBuildStatus[resolvedKey];
    const funcDist = bs ? bs.func : {};
    const techDist = bs ? bs.tech : {};

    return {
      total: cd ? cd.total : 0,
      complete: cd ? cd.complete : 0,
      partial:  cd ? cd.partial  : 0,
      inProgress: cd ? cd.inProgress : 0,
      notStarted: cd ? cd.notStarted : 0,
      blocked:  cd ? cd.blocked : 0,
      sprintEntries,
      funcDist,
      techDist,
      rows: cd ? cd.rows : []
    };
  };

  const getCompCap = (compName) => {
    if (!cap) return null;
    const comp = compName.toLowerCase();
    const items = cap.items.filter(r => {
      const ws = String(r[cap.keys.workstream]||"").toLowerCase();
      return ws.includes(comp.split(" ")[0]) || comp.includes(ws.split(" ")[0]);
    });
    if (!items.length) return null;
    const planned = items.reduce((s,r)=>s+Number(r[cap.keys.planned]||0),0);
    return { planned: Math.round(planned) };
  };

  // ── RAID priority distribution (exclude Complete items) ──────────────────
  const raidByPriority = raid ? (() => {
    const map = {};
    raid.items
      .filter(r => String(r[raid.keys.status]||"").toLowerCase() !== "complete")
      .forEach(r => {
        const p = String(r[raid.keys.priority]||"Unknown");
        const t = String(r[raid.keys.team]||"Unknown");
        if (!map[p]) map[p] = { total:0, open:0, delayed:0, rows:[], byTeam:{} };
        map[p].total++; map[p].rows.push(r);
        const s = String(r[raid.keys.status]||"").toLowerCase();
        if (s.includes("delay")||s.includes("off")) map[p].delayed++;
        else map[p].open++;
        if (!map[p].byTeam[t]) map[p].byTeam[t] = 0;
        map[p].byTeam[t]++;
      });
    return map;
  })() : {};

  // Collect all unique teams for stacked chart colors
  const allTeams = raid ? Array.from(new Set(raid.items.map(r=>String(r[raid.keys.team]||"Unknown")))).sort() : [];
  const teamColors = ["#1a73e8","#f5a623","#c0392b","#27ae60","#8e44ad","#16a085","#d35400","#2980b9","#8e44ad","#c0392b","#7f8c8d","#2c3e50","#e74c3c","#3498db","#1abc9c"];
  const teamColorMap = {};
  allTeams.forEach((t,i) => { teamColorMap[t] = teamColors[i % teamColors.length]; });

  const anyData = wp || raid || req || cap;

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      {!anyData && <Empty label="Upload files above to populate the dashboard." />}

      {/* ── 1. CR SUMMARY ────────────────────────────────────────────────── */}
      {raid && (
        <div>


        </div>
      )}

      {/* ── 2. WORKPLAN MATRIX ────────────────────────────────────────────── */}
      {wp && (() => {
        const WP_STATUS_FILTERS = [
          { label: "All",         color: C.navyLight },
          { label: "At Risk",     color: C.delayed   },
          { label: "On Track",    color: C.onTrack   },
          { label: "Complete",    color: C.complete  },
          { label: "Not Started", color: C.muted     },
        ];

        // Build rows with health pre-computed so we can filter + count
        const wpRows = Object.entries(wp.componentMap).map(([comp, d]) => {
          const pctVal = d.total > 0 ? Math.round(((d.complete + d.onTrack * 0.5) / d.total) * 100) : 0;
          const due = d.rows.filter(r => { const dd = daysUntil(r["Finish"] || r["End Date"]); return dd != null && dd >= 0 && dd <= 14 && !(r["Default Status"] || "").toLowerCase().includes("complete"); }).length;
          const health = d.offTrack > 0 ? "At Risk" : d.onTrack > 0 ? "On Track" : d.complete === d.total ? "Complete" : "Not Started";
          const healthColor = health === "At Risk" ? C.delayed : health === "On Track" ? C.onTrack : health === "Complete" ? C.complete : C.muted;
          return { comp, d, pctVal, due, health, healthColor };
        });

        // Count per status for badges
        const statusCounts = {};
        wpRows.forEach(r => { statusCounts[r.health] = (statusCounts[r.health] || 0) + 1; });

        const filtered = wpStatusFilter === "All" ? wpRows : wpRows.filter(r => r.health === wpStatusFilter);

        return (
          <Card>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12, flexWrap:"wrap", gap:8 }}>
              <SecTitle title="Workplan Summary" color={C.navyLight} />
              {/* Status filter pills */}
              <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                {WP_STATUS_FILTERS.map(({ label, color }) => {
                  const count = label === "All" ? wpRows.length : (statusCounts[label] || 0);
                  const active = wpStatusFilter === label;
                  return (
                    <button key={label} onClick={() => setWpStatusFilter(label)}
                      style={{ display:"flex", alignItems:"center", gap:5, padding:"4px 11px",
                        borderRadius:20, border: `2px solid ${active ? color : C.border}`,
                        background: active ? color : C.white,
                        color: active ? "#fff" : C.muted,
                        cursor:"pointer", fontSize:11, fontWeight:700,
                        transition:"all .15s", boxShadow: active ? `0 2px 8px ${color}40` : "none" }}>
                      {label}
                      <span style={{ background: active ? "rgba(255,255,255,0.3)" : C.border,
                        color: active ? "#fff" : C.text,
                        borderRadius:10, padding:"1px 6px", fontSize:10, fontWeight:800 }}>
                        {count}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead>
                  <tr style={{ background:"#162f50" }}>
                    {["Workstream","Total","On Track","Delayed","% Complete","Due ≤14 Days","Status"].map((h,i)=>(
                      <th key={h} style={{ textAlign:i===0?"left":"center", padding:"9px 12px", color:"#fff", fontSize:11, fontWeight:700, whiteSpace:"nowrap", borderRight:`1px solid rgba(255,255,255,0.1)` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.length === 0 ? (
                    <tr><td colSpan={7} style={{ padding:"24px", textAlign:"center", color:C.muted, fontSize:12 }}>No workstreams match "{wpStatusFilter}"</td></tr>
                  ) : filtered.map(({ comp, d, pctVal, due, health, healthColor }, i) => (
                    <tr key={comp} style={{ background:i%2===0?C.white:"#f7f9fc", borderBottom:`1px solid ${C.border}`,
                      cursor:"pointer", transition:"background .1s" }}
                      onClick={()=>setWpModal({ title:comp, rows:d.rows, initialFilter: d.offTrack > 0 ? "Off Track" : "All" })}
                      onMouseEnter={e=>e.currentTarget.style.background="#eef4ff"}
                      onMouseLeave={e=>e.currentTarget.style.background=i%2===0?C.white:"#f7f9fc"}>
                      <td style={{ padding:"10px 12px", color:C.text, fontWeight:600, fontSize:12 }}>
                        <span style={{ display:"flex", alignItems:"center", gap:6 }}>
                          <span style={{ width:3, height:14, background:healthColor, borderRadius:2, display:"inline-block", flexShrink:0 }} />
                          {comp.replace("Technology - ","")}
                        </span>
                      </td>
                      <td style={{ padding:"10px 12px", textAlign:"center", color:C.text, fontWeight:700 }}>{d.total}</td>
                      <td style={{ padding:"10px 12px", textAlign:"center", color:C.onTrack, fontWeight:700 }}>{d.onTrack}</td>
                      <td style={{ padding:"10px 12px", textAlign:"center", color:d.offTrack>0?C.delayed:C.muted, fontWeight:700 }}>{d.offTrack}</td>
                      <td style={{ padding:"10px 12px", textAlign:"center" }}>
                        <div style={{ display:"flex", alignItems:"center", gap:6, justifyContent:"center" }}>
                          <div style={{ width:60, background:"#e2e8f0", borderRadius:4, height:7, overflow:"hidden" }}>
                            <div style={{ width:`${pctVal}%`, height:"100%", background:pctVal>=75?C.green:pctVal>=40?C.gold:C.delayed, borderRadius:4 }} />
                          </div>
                          <span style={{ color:C.text, fontWeight:700, fontSize:11 }}>{pctVal}%</span>
                        </div>
                      </td>
                      <td style={{ padding:"10px 12px", textAlign:"center", color:due>0?C.yellow:C.muted, fontWeight:700 }}>{due}</td>
                      <td style={{ padding:"10px 12px", textAlign:"center" }}>
                        <span style={{ background:healthColor+"20", color:healthColor, border:`1px solid ${healthColor}40`, borderRadius:4, padding:"2px 8px", fontSize:10, fontWeight:700 }}>{health}</span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ color:C.muted, fontSize:10, marginTop:8 }}>💡 Click any row to drill down with Lvl 2 / 3 / 4 filters</div>
          </Card>
        );
      })()}


            {/* Modals */}
      {raidModal && <RaidDrillModal title={raidModal.title} rows={raidModal.rows} raidKeys={raid?.keys} onClose={()=>setRaidModal(null)} />}
      {wpModal && <WorkplanDrillModal title={wpModal.title} rows={wpModal.rows} initialFilter={wpModal.initialFilter} onClose={()=>setWpModal(null)} />}
    </div>
  );
}


// ─── COMPONENT SCORECARD TAB ─────────────────────────────────────────────────
function StoryDrillModal({ title, rows, reqKeys, onClose }) {
  const [statusFilter, setStatusFilter] = useState("All");
  const [expanded, setExpanded] = useState({});
  if (!rows?.length) return null;
  const K = reqKeys || {};

  // Status rank for Overall Status derivation
  const STATUS_RANK = { "blocked":6, "in progress":5, "partial build complete":4, "not started":3, "complete":2, "n/a":1 };
  const getRank = v => { const s=String(v||"").toLowerCase(); for(const[k,r] of Object.entries(STATUS_RANK)){if(s.includes(k))return r;} return 0; };

  // Derive overall status per row: worst of Func + Tech build status
  const getOverallStatus = r => {
    const fb = String(r[K.funcBuildStatus]||"").trim();
    const tb = String(r[K.techBuildStatus]||"").trim();
    if (!fb && !tb) return "—";
    if (!fb) return tb;
    if (!tb) return fb;
    return getRank(fb) >= getRank(tb) ? fb : tb;
  };

  // Column definitions in the required order
  const COLS = [
    { key: K.pmExperience    || "PM Experience",                         label: "PM Experience" },
    { key: K.component       || "Sub Process",                           label: "Sub Process" },
    { key: K.reqId           || "Req Id",                                label: "Req Id" },
    { key: K.bizReq          || "Business Requirements",                 label: "Business Requirements" },
    { key: K.story           || "User Story",                            label: "User Story" },
    { key: K.acceptance      || "Acceptance Criteria",                   label: "Acceptance Criteria" },
    { key: K.derivedStatus   || "User Story Review Status (D&A)",        label: "User Story Review Status (D&A)" },
    { key: K.sprint          || "Build Cycle",                           label: "Build Cycle (Playback)" },
    { key: K.closureSprint   || "Targeted Closure Sprint",               label: "Targeted Closure Sprint" },
    { key: K.funcBuildStatus || "Functional Status Master List",         label: "Functional Build Status" },
    { key: K.techBuildStatus || "Technical Status Master List",          label: "Tech Build Status" },
    { key: "__overall__",                                                 label: "Overall Story Status" },
    ...(K.buildComment ? [{ key: K.buildComment, label: "Build Management Comments", isBuildComment: true }] : []),
  ].filter(c => c.key);

  const WIDE_COLS = new Set([
    K.bizReq, K.story, K.acceptance, K.derivedStatus, K.buildComment,
    "Business Requirements","User Story","Acceptance Criteria","Derived Status","Build Management Comments"
  ]);

  // Build Overall Status buckets for filter buttons — based on getOverallStatus (same as last column)
  const OVERALL_FILTERS = [
    { label: "All",                  match: null,                bg: "rgba(255,255,255,0.2)",  color: "#fff",    border: "rgba(255,255,255,0.4)" },
    { label: "Blocked",              match: "block",             bg: "#b91c1c",                color: "#fff",    border: "#b91c1c" },
    { label: "In Progress",          match: "progress",          bg: C.inProgress,             color: "#fff",    border: C.inProgress },
    { label: "Partial",              match: "partial",           bg: "#0369a1",                color: "#fff",    border: "#0369a1" },
    { label: "Not Started",          match: "not start",         bg: "#475569",                color: "#fff",    border: "#475569" },
    { label: "Complete",             match: "complete",          bg: C.complete,               color: "#fff",    border: C.complete },
    { label: "N/A",                  match: "n/a",               bg: "#7e22ce",                color: "#fff",    border: "#7e22ce" },
  ];

  const getOverallBucket = r => {
    const ov = getOverallStatus(r).toLowerCase();
    if (ov.includes("block"))                                    return "Blocked";
    if (ov.includes("progress"))                                 return "In Progress";
    if (ov.includes("partial") && !ov.includes("complete"))      return "Partial";
    if (ov.includes("complete") && !ov.includes("partial"))      return "Complete";
    if (ov.includes("n/a") || ov === "na")                       return "N/A";
    if (ov.includes("not start") || ov === "—")                  return "Not Started";
    return "Not Started";
  };

  // Counts per bucket
  const bucketCounts = OVERALL_FILTERS.reduce((acc, f) => {
    acc[f.label] = f.label === "All" ? rows.length : rows.filter(r => getOverallBucket(r) === f.label).length;
    return acc;
  }, {});

  // Filter rows by selected bucket
  const filtered = statusFilter === "All"
    ? rows
    : rows.filter(r => getOverallBucket(r) === statusFilter);

  const statusColor = s => {
    const v = String(s||"").toLowerCase();
    if (v.includes("complete") && !v.includes("partial")) return C.complete;
    if (v.includes("partial"))   return "#60a5fa";
    if (v.includes("progress"))  return C.inProgress;
    if (v.includes("block"))     return C.delayed;
    if (v.includes("n/a")||v==="na") return C.muted;
    return C.notStarted;
  };

  // Group filtered rows by Sub Process
  const groups = {};
  filtered.forEach(r => {
    const g = String(r[K.component] || "Unknown").trim() || "Unknown";
    if (!groups[g]) groups[g] = [];
    groups[g].push(r);
  });
  const sortedGroups = Object.entries(groups).sort((a, b) => a[0].localeCompare(b[0]));

  const isOpen     = key => expanded[key] !== false; // default open
  const toggle     = key => setExpanded(p => ({ ...p, [key]: !isOpen(key) }));
  const expandAll  = () => { const e = {}; sortedGroups.forEach(([g]) => e[g] = true);  setExpanded(e); };
  const collapseAll= () => { const e = {}; sortedGroups.forEach(([g]) => e[g] = false); setExpanded(e); };

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:1000,
      display:"flex", alignItems:"center", justifyContent:"center" }} onClick={onClose}>
      <div style={{ background:C.white, borderRadius:10, width:"98%", maxWidth:1400,
        maxHeight:"90vh", display:"flex", flexDirection:"column",
        boxShadow:"0 24px 60px rgba(0,0,0,0.35)" }} onClick={e => e.stopPropagation()}>

        {/* Header */}
        <div style={{ background:C.headerBg, padding:"12px 20px", borderRadius:"10px 10px 0 0",
          display:"flex", flexDirection:"column", gap:10, flexShrink:0 }}>
          {/* Title + expand/collapse + close */}
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:12, flexWrap:"wrap" }}>
            <div style={{ color:"#fff", fontWeight:700, fontSize:13 }}>
              {title}
              <span style={{ opacity:.6, fontWeight:400, marginLeft:8 }}>
                ({filtered.length} of {rows.length} stories · {sortedGroups.length} sub processes)
              </span>
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <button onClick={expandAll}
                style={{ padding:"4px 12px", borderRadius:4, border:"1px solid rgba(255,255,255,0.4)", background:"rgba(255,255,255,0.15)", color:"#fff", fontSize:11, cursor:"pointer", fontWeight:600 }}>
                ⊞ Expand All
              </button>
              <button onClick={collapseAll}
                style={{ padding:"4px 12px", borderRadius:4, border:"1px solid rgba(255,255,255,0.4)", background:"rgba(255,255,255,0.15)", color:"#fff", fontSize:11, cursor:"pointer", fontWeight:600 }}>
                ⊟ Collapse All
              </button>
              <button onClick={onClose} style={{ background:"rgba(255,255,255,0.15)", border:"none",
                color:"#fff", borderRadius:5, padding:"5px 14px", cursor:"pointer", fontSize:13, fontWeight:600 }}>
                ✕ Close
              </button>
            </div>
          </div>
          {/* Overall Status filter bubbles */}
          <div style={{ display:"flex", gap:6, flexWrap:"wrap", alignItems:"center" }}>
            <span style={{ color:"rgba(255,255,255,0.5)", fontSize:10, fontWeight:600, textTransform:"uppercase", letterSpacing:"0.06em", marginRight:4 }}>Overall Status</span>
            {OVERALL_FILTERS.filter(f => f.label === "All" || (bucketCounts[f.label] ?? 0) > 0).map(f => {
              const active = statusFilter === f.label;
              return (
                <button key={f.label} onClick={() => setStatusFilter(f.label)}
                  style={{ display:"flex", alignItems:"center", gap:5, padding:"4px 12px",
                    borderRadius:20,
                    border: `2px solid ${active ? f.border : "rgba(255,255,255,0.25)"}`,
                    background: active ? f.bg : "rgba(255,255,255,0.08)",
                    color: active ? f.color : "rgba(255,255,255,0.7)",
                    cursor:"pointer", fontSize:11, fontWeight: active ? 700 : 500,
                    transition:"all .12s",
                    boxShadow: active ? `0 2px 8px rgba(0,0,0,0.3)` : "none" }}>
                  {f.label}
                  <span style={{ background: active ? "rgba(0,0,0,0.25)" : "rgba(255,255,255,0.15)",
                    borderRadius:10, padding:"1px 6px", fontSize:10, fontWeight:700,
                    color: active ? f.color : "rgba(255,255,255,0.8)" }}>
                    {bucketCounts[f.label] ?? 0}
                  </span>
                </button>
              );
            })}
          </div>
        </div>

        {/* Table */}
        <div style={{ overflowY:"auto", overflowX:"auto", flex:1 }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
            <thead style={{ position:"sticky", top:0, background:"#f0f4f8", zIndex:2 }}>
              <tr>
                {COLS.map(c => (
                  <th key={c.key} style={{ textAlign:"left", padding:"8px 10px", color:C.muted,
                    fontWeight:700, borderBottom:`2px solid ${C.border}`, whiteSpace:"nowrap",
                    minWidth: c.key === "__overall__" ? 160 : c.isBuildComment ? 280 : WIDE_COLS.has(c.key) ? 240 : 110 }}>
                    {c.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sortedGroups.map(([groupName, groupRows]) => {
                const open = isOpen(groupName);
                // Worst overall status in this group for the header badge
                const worstOv = groupRows.reduce((worst, r) => {
                  const ov = getOverallStatus(r);
                  return getRank(ov) > getRank(worst) ? ov : worst;
                }, "");
                const worstColor = worstOv ? statusColor(worstOv) : C.muted;
                const hasBlocked = groupRows.some(r => getOverallBucket(r) === "Blocked");
                const groupBg = hasBlocked ? "#fef2f2" : "#e8eef7";
                return (
                  <React.Fragment key={groupName}>
                    {/* Group header row */}
                    <tr onClick={() => toggle(groupName)}
                      style={{ background: groupBg, cursor:"pointer", borderBottom:`1px solid ${C.border}` }}>
                      <td colSpan={COLS.length} style={{ padding:"9px 14px" }}>
                        <span style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
                          <span style={{ fontSize:13, userSelect:"none", lineHeight:1 }}>{open ? "▼" : "▶"}</span>
                          <span style={{ fontWeight:700, color:C.navy, fontSize:12 }}>{groupName}</span>
                          <span style={{ background:C.navyLight+"25", color:C.navyLight, borderRadius:10, padding:"1px 9px", fontSize:10, fontWeight:700 }}>
                            {groupRows.length} {groupRows.length === 1 ? "story" : "stories"}
                          </span>
                          {worstOv && worstOv !== "—" && (
                            <span style={{ background:worstColor+"20", color:worstColor, border:`1px solid ${worstColor}40`, borderRadius:10, padding:"1px 9px", fontSize:10, fontWeight:700 }}>
                              {worstOv}
                            </span>
                          )}
                          {hasBlocked && (
                            <span style={{ background:C.delayed+"20", color:C.delayed, border:`1px solid ${C.delayed}40`, borderRadius:10, padding:"1px 9px", fontSize:10, fontWeight:700 }}>
                              ⚠ Blocked
                            </span>
                          )}
                          <span style={{ color:C.muted, fontSize:10, marginLeft:"auto" }}>
                            {open ? "Click to collapse" : "Click to expand"}
                          </span>
                        </span>
                      </td>
                    </tr>
                    {/* Story rows */}
                    {open && groupRows.map((r, i) => (
                      <tr key={i} style={{ background: i%2===0 ? "#f9fbff" : C.white,
                        borderBottom:`1px solid ${C.border}`, verticalAlign:"top" }}>
                        {COLS.map(c => {
                          const v = String(r[c.key] || "—").trim();
                          const isStatus = c.key === K.derivedStatus || c.key === K.status
                                        || c.key === K.funcBuildStatus || c.key === K.techBuildStatus;
                          const isWide   = WIDE_COLS.has(c.key);
                          const sc = isStatus ? statusColor(v) : null;
                          return (
                            <td key={c.key} style={{
                              padding:"7px 10px",
                              color: C.text,
                              whiteSpace: isWide ? "normal" : "nowrap",
                              maxWidth: isWide ? 280 : 160,
                              overflow: isWide ? "visible" : "hidden",
                              textOverflow: isWide ? "unset" : "ellipsis",
                              wordBreak: isWide ? "break-word" : "normal",
                              lineHeight: isWide ? 1.4 : "normal",
                              fontSize: 11,
                            }} title={isWide ? undefined : v}>
                              {c.key === "__overall__"
                                ? (() => {
                                    const ov = getOverallStatus(r);
                                    const ovc = statusColor(ov);
                                    return ov === "—"
                                      ? <span style={{ color:C.muted }}>—</span>
                                      : <span style={{ background:ovc+"20", color:ovc, border:`1px solid ${ovc}40`,
                                          borderRadius:3, padding:"2px 8px", fontSize:10, fontWeight:700,
                                          whiteSpace:"nowrap" }}>{ov}</span>;
                                  })()
                                : c.isBuildComment
                                  ? <CommentCell value={r[c.key]} />
                                : isStatus && sc && v !== "—"
                                  ? <span style={{ background:sc+"20", color:sc, border:`1px solid ${sc}40`,
                                      borderRadius:3, padding:"2px 7px", fontSize:10, fontWeight:700,
                                      whiteSpace:"nowrap" }}>{v}</span>
                                  : isWide ? v : v.slice(0, 60)}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>

      </div>
    </div>
  );
}

// Sprint bubble colour — self-contained, no external deps
// Checks ALL possible status string variations since bucket names may vary
const sprintBubbleColor = (sd) => {
  if (!sd) return { bg:"#f1f5f9", color:"#94a3b8", border:"#cbd5e1" };
  const b = Number(sd.blocked    || sd.Blocked    || 0);
  const c = Number(sd.complete   || sd.Complete   || 0);
  const p = Number(sd.partial    || sd.Partial    || 0);
  const ip= Number(sd.inProgress || sd.InProgress || sd["in progress"] || sd["In Progress"] || 0);
  const n = Number(sd.notStarted || sd.NotStarted || sd["not started"] || sd["Not Started"] || 0);
  const a = Number(sd.na         || sd.NA         || sd["n/a"]         || sd["N/A"]         || 0);
  const total = b + c + p + ip + n + a;
  if (total === 0) return { bg:"#f1f5f9", color:"#94a3b8", border:"#cbd5e1" };
  if (b > 0)                            return { bg:"#fee2e2", color:"#b91c1c", border:"#fca5a5" }; // 🔴 Red — blocked
  if (c > 0 && (c + a) >= total)        return { bg:"#dbeafe", color:"#1d4ed8", border:"#93c5fd" }; // 🔵 Blue — all complete
  if (ip > 0 || p > 0)                  return { bg:"#dcfce7", color:"#15803d", border:"#86efac" }; // 🟢 Green — in progress
  if (n > 0)                            return { bg:"#f1f5f9", color:"#64748b", border:"#cbd5e1" }; // ⚫ Grey — not started
  return                                       { bg:"#eff6ff", color:"#3b82f6", border:"#bfdbfe" }; // 🔷 Mixed
};

// ─── COMPONENT CARDS TAB ─────────────────────────────────────────────────────
function ComponentCardsTab({ wp, raid, req, openModal }) {
  const [raidModal, setRaidModal] = useState(null);
  const [storyModal, setStoryModal] = useState(null);
  const [wpDrillModal, setWpDrillModal] = useState(null);
  if (!raid && !req && !wp) return <Empty label="Upload files to view Component Cards." />;

  // ── Same aliases + helpers as ScorecardTab ────────────────────────────────
  const COMP_ALIASES = {
    "carr": "Career Advancement Review",
    "career advancement reviiew": "Career Advancement Review",
    "career advancement review": "Career Advancement Review",
    "career advancement review (carr)": "Career Advancement Review",
    "career advancement readiness review": "Career Advancement Review",
    "career advancement readiness review (carr)": "Career Advancement Review",
    "expectation framework": "Expectations Framework",
    "expectations framework": "Expectations Framework",
  };
  const normaliseComp = (name) => {
    const key = String(name || "").toLowerCase().trim();
    if (COMP_ALIASES[key]) return COMP_ALIASES[key];
    for (const [alias, canonical] of Object.entries(COMP_ALIASES)) {
      if (key.includes(alias) || alias.includes(key)) return canonical;
    }
    return name;
  };

  const getCompRaid = (compName) => {
    if (!raid) return { open:0, delayed:0, issues:[], risks:[], openItems:[] };
    const normComp = normaliseComp(compName);
    const items = raid.items.filter(r => normaliseComp(String(r[raid.keys.component]||"")) === normComp);
    const open    = items.filter(r => { const s = String(r[raid.keys.status]||"").toLowerCase(); return s !== "complete" && s !== "deferred"; });
    const delayed = items.filter(r => String(r[raid.keys.status]||"").toLowerCase() === "delayed");
    const issues  = open.filter(r => String(r[raid.keys.type]||"").toLowerCase().includes("issue"));
    const risks   = open.filter(r => String(r[raid.keys.type]||"").toLowerCase().includes("risk"));
    return { open: open.length, delayed: delayed.length, issues, risks, openItems: open };
  };

  const wpWorstStatus = (rows) => {
    let worst = null, worstRank = -1;
    rows.forEach(r => {
      const sl = String(r["Default Status"]||r["Status"]||"").toLowerCase();
      const rank = sl.includes("off track")?4:sl.includes("on track")?3:sl.includes("not start")?2:sl.includes("complete")?1:0;
      if (rank > worstRank) { worstRank = rank; worst = String(r["Default Status"]||r["Status"]||""); }
    });
    return worst || "—";
  };

  const getCompWp = (compName) => {
    if (!wp) return null;
    const scopedRows = wp.allRows.filter(r =>
      String(r["Activity Grp - Lvl 1"]||"").trim() === "Technology - SAP Configuration & Build" &&
      String(r["Activity Grp - Lvl 2"]||"").trim() === "Component Build"
    );
    const normComp = normaliseComp(compName);
    const lvl3Rows = scopedRows.filter(r => normaliseComp(String(r["Activity Grp - Lvl 3"]||"").trim()) === normComp);
    if (!lvl3Rows.length) return null;
    const lvl3Names = Array.from(new Set(lvl3Rows.map(r => String(r["Activity Grp - Lvl 3"]||"").trim())));
    const subtreeRows = scopedRows.filter(r => lvl3Names.includes(String(r["Activity Grp - Lvl 3"]||"").trim()));
    const isLeafRow = r => { const c = r["Children"]; return !c || Number(c) === 0; };
    const lvl4PlusRows = subtreeRows.filter(r => Number(r["Lvl"]??0) >= 4);
    const isDesign = r => /design/i.test(String(r["Task Name"]||"")+String(r["Activity Grp - Lvl 4"]||""));
    const isBuild  = r => /build|develop|implement|code/i.test(String(r["Task Name"]||"")+String(r["Activity Grp - Lvl 4"]||""));
    const designLeaves = lvl4PlusRows.filter(r => isLeafRow(r) && isDesign(r));
    const buildLeaves  = lvl4PlusRows.filter(r => isLeafRow(r) && isBuild(r));
    const designStatus = designLeaves.length ? wpWorstStatus(designLeaves) : null;
    const buildStatus  = buildLeaves.length  ? wpWorstStatus(buildLeaves)  : null;
    const leafRows = subtreeRows.filter(isLeafRow);
    const pctValues = leafRows.map(r => {
      const v = r["% Complete"]??r["% complete"];
      if (v!==""&&v!=null&&!isNaN(Number(v))) return Number(v);
      const s = String(r["Default Status"]||r["Status"]||"").toLowerCase();
      if (s.includes("complete")) return 100;
      if (s.includes("on track")||s.includes("in progress")) return 50;
      if (s.includes("off track")||s.includes("delayed")) return 25;
      if (s.includes("not start")) return 0;
      return null;
    }).filter(v => v != null);
    const pctComplete = pctValues.length ? Math.round(pctValues.reduce((a,b)=>a+b,0)/pctValues.length) : null;

    // build drill rows for click-through
    const designLvl4Groups = Array.from(new Set(lvl4PlusRows.filter(isDesign).map(r=>String(r["Activity Grp - Lvl 4"]||"")).filter(Boolean)));
    const buildLvl4Groups  = Array.from(new Set(lvl4PlusRows.filter(isBuild).map(r=>String(r["Activity Grp - Lvl 4"]||"")).filter(Boolean)));
    const buildDrillRows = (groups) => {
      if (!groups.length) return [];
      const hdrs = subtreeRows.filter(r => Number(r["Lvl"]??0) === 3);
      const sub  = subtreeRows.filter(r => groups.includes(String(r["Activity Grp - Lvl 4"]||"").trim()));
      const seen = new Set();
      return [...hdrs,...sub].filter(r => { const id=r["Row ID"]||JSON.stringify(r); if(seen.has(id))return false; seen.add(id); return true; });
    };
    return { designStatus, buildStatus, pctComplete, designRows: buildDrillRows(designLvl4Groups), buildRows: buildDrillRows(buildLvl4Groups), allRows: subtreeRows };
  };

  const getCompReq = (compName) => {
    if (!req||!req.byComponent) return null;
    const normName = normaliseComp(compName);
    const key = Object.keys(req.byComponent).find(k => normaliseComp(k) === normName);
    if (!key) return null;
    const cd = req.byComponent[key];
    const sprintData = req.compBySprint ? (req.compBySprint[key]||{}) : {};
    const bs = req.compBuildStatus ? (req.compBuildStatus[key]||null) : null;
    return { total:cd.total, complete:cd.complete, partial:cd.partial, inProgress:cd.inProgress,
      notStarted:cd.notStarted, blocked:cd.blocked, na:cd.na||0, sprintData,
      funcDist:bs?bs.func:{}, techDist:bs?bs.tech:{}, rows:cd.rows };
  };

  // Sprint order
  const rawSprintOrder = req?.sprintOrder?.filter(s => s && s !== "No Sprint") || [];
  const sprintLabelMap = {};
  rawSprintOrder.forEach(sp => {
    const m = String(sp).toLowerCase().match(/^\s*(\d+)\.\s*s(\d+)/);
    if (m) { const n=parseInt(m[2]); if(n>=1&&n<=8) sprintLabelMap[sp]=`S${n}`; else sprintLabelMap[sp]=null; }
    else { const nm=String(sp).toLowerCase().match(/s(\d+)/); if(nm){const n=parseInt(nm[1]);if(n>=1&&n<=8)sprintLabelMap[sp]=`S${n}`;else sprintLabelMap[sp]=null;}else sprintLabelMap[sp]=null; }
  });
  const sprintOrder = [];
  const seenLabels = new Set();
  ["S1","S2","S3","S4","S5","S6","S7","S8"].forEach(lbl => {
    const raws = rawSprintOrder.filter(sp => sprintLabelMap[sp] === lbl);
    if (raws.length > 0 && !seenLabels.has(lbl)) { sprintOrder.push({ label:lbl, raws }); seenLabels.add(lbl); }
  });
  const getSprintData = (sprintData, entry) => {
    const c = { complete:0, partial:0, inProgress:0, notStarted:0, blocked:0, na:0, total:0 };
    entry.raws.forEach(raw => { const d=sprintData[raw]; if(d) Object.keys(c).forEach(k=>{c[k]+=(d[k]||0);}); });
    return c.total > 0 ? c : null;
  };

  // Component list
  const raidComps = raid ? Array.from(new Set(raid.items.map(r=>normaliseComp(String(r[raid.keys.component]||""))).filter(Boolean))).sort() : [];
  const reqComps  = req  ? Array.from(new Set(Object.keys(req.byComponent||{}).map(normaliseComp))).sort() : [];
  const wpComps   = wp   ? Array.from(new Set(wp.allRows.filter(r=>String(r["Activity Grp - Lvl 1"]||"").trim()==="Technology - SAP Configuration & Build"&&String(r["Activity Grp - Lvl 2"]||"").trim()==="Component Build").map(r=>normaliseComp(String(r["Activity Grp - Lvl 3"]||"").trim())).filter(Boolean))).sort() : [];
  const allComps  = Array.from(new Set([...raidComps,...reqComps,...wpComps])).sort();

  // Status colour helpers
  const sprintBubble = (sd) => {
    if (!sd) return { bg:"#f1f5f9", col:"#94a3b8" };
    if (sd.blocked > 0) return { bg:"#fee2e2", col:"#b91c1c" };
    if (sd.complete > 0 && (sd.complete+sd.na) >= sd.total) return { bg:"#dbeafe", col:"#1d4ed8" };
    if (sd.inProgress > 0 || sd.partial > 0) return { bg:"#dcfce7", col:"#15803d" };
    if (sd.notStarted > 0) return { bg:"#f1f5f9", col:"#64748b" };
    return { bg:"#f1f5f9", col:"#94a3b8" };
  };

  const wpPillStyle = (status) => {
    if (!status||status==="—") return null;
    const sl = status.toLowerCase();
    return {
      bg:    sl.includes("off track")?"#fee2e2":sl.includes("on track")?"#fef9e7":sl.includes("complete")?"#dbeafe":"#f1f5f9",
      col:   sl.includes("off track")?"#b91c1c":sl.includes("on track")?"#b45309":sl.includes("complete")?"#1d4ed8":"#64748b",
      border:sl.includes("off track")?"#fca5a5":sl.includes("on track")?"#fcd34d":sl.includes("complete")?"#93c5fd":"#cbd5e1",
    };
  };

  const consolidatedPill = (dist) => {
    if (!dist||!Object.keys(dist).length) return null;
    const s = Object.entries(dist).sort((a,b)=>b[1]-a[1]);
    const top = s[0][0].toLowerCase();
    const bg  = top.includes("block")?"#fee2e2":top.includes("complete")&&!top.includes("partial")?"#dbeafe":top.includes("progress")||top.includes("partial")?"#dcfce7":top.includes("not start")?"#f1f5f9":"#fef3c7";
    const col = top.includes("block")?"#991b1b":top.includes("complete")&&!top.includes("partial")?"#1d4ed8":top.includes("progress")||top.includes("partial")?"#166534":top.includes("not start")?"#475569":"#92400e";
    return { label: s[0][0], bg, col };
  };

  // Accent colour for card left border based on worst status
  const cardAccent = (rc, rq, cw) => {
    if (rc.delayed > 0) return C.delayed;
    if (rq?.blocked > 0) return C.delayed;
    if (cw?.buildStatus?.toLowerCase().includes("off track")) return C.delayed;
    if (cw?.designStatus?.toLowerCase().includes("off track")) return C.delayed;
    const pct = cw?.pctComplete;
    if (pct != null && pct >= 75) return C.complete;
    if (rq?.inProgress > 0) return C.onTrack;
    return "#94a3b8";
  };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>

      {/* Header summary row */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(120px,1fr))", gap:10 }}>
        {[
          ["Components", allComps.length, C.navyLight],
          ["Total Stories", req ? Object.values(req.byComponent||{}).reduce((s,d)=>s+d.total,0) : "—", C.navyLight],
          ["Blocked", req ? Object.values(req.byComponent||{}).reduce((s,d)=>s+(d.blocked||0),0) : "—", C.delayed],
          ["In Progress", req ? Object.values(req.byComponent||{}).reduce((s,d)=>s+(d.inProgress||0),0) : "—", C.onTrack],
          ["Complete", req ? Object.values(req.byComponent||{}).reduce((s,d)=>s+(d.complete||0),0) : "—", C.complete],
          ["Open RAIDs", raid ? raid.open.length : "—", C.gold],
          ["Delayed RAIDs", raid ? raid.delayed.length : "—", C.delayed],
        ].map(([lbl,val,col]) => (
          <div key={lbl} style={{ background:C.white, border:`1px solid ${C.border}`, borderTop:`3px solid ${col}`, borderRadius:7, padding:"10px 12px" }}>
            <div style={{ fontSize:9, color:C.muted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:3 }}>{lbl}</div>
            <div style={{ fontSize:22, fontWeight:800, color:col, lineHeight:1 }}>{val}</div>
          </div>
        ))}
      </div>

      {/* One card per component */}
      <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
        {allComps.map(comp => {
          const rc = getCompRaid(comp);
          const rq = getCompReq(comp);
          const cw = getCompWp(comp);
          const hasData = rc.open > 0 || rc.delayed > 0 || (rq && rq.total > 0) || cw;
          if (!hasData) return null;

          const accent = cardAccent(rc, rq, cw);
          const pct = cw?.pctComplete;
          const accentLight = accent === C.delayed ? "rgba(192,57,43,0.06)" : accent === C.complete ? "rgba(26,115,232,0.05)" : accent === C.onTrack ? "rgba(245,166,35,0.05)" : "rgba(148,163,184,0.03)";

          return (
            <div key={comp} style={{ background:C.white, border:`0.5px solid ${C.border}`, borderLeft:`4px solid ${accent}`, borderRadius:"0 10px 10px 0", overflow:"hidden", position:"relative" }}>
              {/* Tint wash */}
              <div style={{ position:"absolute", inset:0, background:accentLight, pointerEvents:"none" }} />

              <div style={{ display:"grid", gridTemplateColumns:"220px 160px 1fr 1fr 200px", gap:0, position:"relative" }}>

                {/* ── Col 1: Identity + RAID ── */}
                <div style={{ padding:"14px 16px", borderRight:`1px solid ${C.border}` }}>
                  <div style={{ fontSize:13, fontWeight:700, color:C.navy, marginBottom:8, lineHeight:1.3 }}>{comp}</div>

                  {/* RAID badges */}
                  <div style={{ display:"flex", flexWrap:"wrap", gap:5, marginBottom:6 }}>
                    {rc.open > 0 && (
                      <span onClick={()=>setRaidModal({title:`${comp} — Open RAIDs`, rows:rc.openItems})}
                        style={{ background:"#fef3c7", color:"#92400e", border:"1px solid #fcd34d", borderRadius:5, padding:"2px 8px", fontSize:10, fontWeight:600, cursor:"pointer" }}>
                        {rc.open} RAID open
                      </span>
                    )}
                    {rc.delayed > 0 && (
                      <span onClick={()=>setRaidModal({title:`${comp} — Delayed RAIDs`, rows:raid.items.filter(r=>normaliseComp(String(r[raid.keys.component]||""))===normaliseComp(comp)&&String(r[raid.keys.status]||"").toLowerCase()==="delayed")})}
                        style={{ background:"#fee2e2", color:"#991b1b", border:"1px solid #fca5a5", borderRadius:5, padding:"2px 8px", fontSize:10, fontWeight:600, cursor:"pointer" }}>
                        ⚠ {rc.delayed} delayed
                      </span>
                    )}
                    {rc.issues.length > 0 && (
                      <span onClick={()=>setRaidModal({title:`${comp} — Issues`, rows:rc.issues})}
                        style={{ background:"#fee2e2", color:"#991b1b", border:"1px solid #fca5a5", borderRadius:5, padding:"2px 8px", fontSize:10, cursor:"pointer" }}>
                        {rc.issues.length} issue{rc.issues.length>1?"s":""}
                      </span>
                    )}
                    {rc.risks.length > 0 && (
                      <span onClick={()=>setRaidModal({title:`${comp} — Risks`, rows:rc.risks})}
                        style={{ background:"#fef3c7", color:"#92400e", border:"1px solid #fcd34d", borderRadius:5, padding:"2px 8px", fontSize:10, cursor:"pointer" }}>
                        {rc.risks.length} risk{rc.risks.length>1?"s":""}
                      </span>
                    )}
                    {rc.open === 0 && <span style={{ color:C.muted, fontSize:10 }}>No open RAIDs</span>}
                  </div>
                </div>

                {/* ── Col 2: Workplan % + Design/Build status ── */}
                <div style={{ padding:"14px 16px", borderRight:`1px solid ${C.border}` }}>
                  <div style={{ fontSize:9, color:C.muted, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:6 }}>Workplan</div>
                  {cw ? (
                    <>
                      <div style={{ fontSize:28, fontWeight:800, color:accent, lineHeight:1, marginBottom:6 }}>
                        {pct != null ? `${pct}%` : "—"}
                      </div>
                      <div style={{ background:"#e2e8f0", borderRadius:3, height:5, overflow:"hidden", marginBottom:8 }}>
                        <div style={{ width:`${pct??0}%`, height:"100%", background:accent, borderRadius:3 }} />
                      </div>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        {cw.designStatus && (() => { const ps = wpPillStyle(cw.designStatus); return ps ? (
                          <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                            <span style={{ fontSize:9, color:C.muted, minWidth:36 }}>Design</span>
                            <span onClick={()=>setWpDrillModal({title:`${comp} — Design`, rows:cw.designRows})}
                              style={{ background:ps.bg, color:ps.col, border:`1px solid ${ps.border}`, borderRadius:4, padding:"1px 7px", fontSize:10, cursor:"pointer" }}>
                              {cw.designStatus}
                            </span>
                          </div>
                        ) : null; })()}
                        {cw.buildStatus && (() => { const ps = wpPillStyle(cw.buildStatus); return ps ? (
                          <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                            <span style={{ fontSize:9, color:C.muted, minWidth:36 }}>Build</span>
                            <span onClick={()=>setWpDrillModal({title:`${comp} — Build`, rows:cw.buildRows})}
                              style={{ background:ps.bg, color:ps.col, border:`1px solid ${ps.border}`, borderRadius:4, padding:"1px 7px", fontSize:10, cursor:"pointer" }}>
                              {cw.buildStatus}
                            </span>
                          </div>
                        ) : null; })()}
                      </div>
                    </>
                  ) : <span style={{ color:C.muted, fontSize:11 }}>No workplan data</span>}
                </div>

                {/* ── Col 3: Sprint bubbles ── */}
                <div style={{ padding:"14px 16px", borderRight:`1px solid ${C.border}` }}>
                  <div style={{ fontSize:9, color:C.muted, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:8 }}>Stories by Sprint</div>
                  {rq ? (
                    <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
                      {sprintOrder.map(entry => {
                        const sd = getSprintData(rq.sprintData, entry);
                        if (!sd) return null;
                        const { bg, col } = sprintBubble(sd);
                        const sprintRows = rq.rows.filter(r => entry.raws.some(raw => String(r[req.keys?.sprint]||"") === raw));
                        return (
                          <div key={entry.label} onClick={()=>setStoryModal({title:`${comp} — ${entry.label}`, rows:sprintRows})}
                            style={{ textAlign:"center", cursor:"pointer" }}>
                            <div style={{ fontSize:9, color:C.muted, marginBottom:2 }}>{entry.label}</div>
                            <span style={{ background:bg, color:col, borderRadius:5, padding:"3px 9px", fontSize:11, fontWeight:700, display:"inline-block" }}>{sd.total}</span>
                          </div>
                        );
                      })}
                      {/* Total */}
                      {rq.total > 0 && (
                        <div onClick={()=>setStoryModal({title:`${comp} — All Stories`, rows:rq.rows})}
                          style={{ textAlign:"center", cursor:"pointer", marginLeft:4 }}>
                          <div style={{ fontSize:9, color:C.muted, marginBottom:2 }}>Total</div>
                          <span style={{ background:C.navyLight+"20", color:C.navyLight, border:`1px solid ${C.navyLight}40`, borderRadius:5, padding:"3px 9px", fontSize:11, fontWeight:700, display:"inline-block" }}>{rq.total}</span>
                        </div>
                      )}
                    </div>
                  ) : <span style={{ color:C.muted, fontSize:11 }}>No story data</span>}
                </div>

                {/* ── Col 4: Func + Tech build status distribution ── */}
                <div style={{ padding:"14px 16px", borderRight:`1px solid ${C.border}` }}>
                  <div style={{ fontSize:9, color:C.muted, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:8 }}>Build Status</div>
                  {rq ? (
                    <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                      {[["Func", rq.funcDist], ["Tech", rq.techDist]].map(([label, dist]) => {
                        const pill = consolidatedPill(dist);
                        const total = Object.values(dist||{}).reduce((s,v)=>s+v,0);
                        return (
                          <div key={label} style={{ display:"flex", alignItems:"center", gap:8 }}>
                            <span style={{ fontSize:10, color:C.muted, minWidth:28, fontWeight:600 }}>{label}</span>
                            {pill ? (
                              <span onClick={()=>{ const drillRows=rq.rows.filter(r=>{const v=String(r[label==="Func"?req.keys?.funcBuildStatus:req.keys?.techBuildStatus]||"").toLowerCase();return v.includes(pill.label.toLowerCase().slice(0,6));}); if(drillRows.length) setStoryModal({title:`${comp} — ${label}: ${pill.label}`, rows:drillRows}); }}
                                style={{ background:pill.bg, color:pill.col, borderRadius:4, padding:"2px 8px", fontSize:10, cursor:"pointer" }}>
                                {pill.label} ({total})
                              </span>
                            ) : <span style={{ color:C.muted, fontSize:10 }}>—</span>}
                          </div>
                        );
                      })}
                    </div>
                  ) : <span style={{ color:C.muted, fontSize:11 }}>No build data</span>}
                </div>

                {/* ── Col 5: Story status breakdown ── */}
                <div style={{ padding:"14px 16px" }}>
                  <div style={{ fontSize:9, color:C.muted, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:8 }}>Story Breakdown</div>
                  {rq ? (
                    <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                      {[
                        ["Blocked",     rq.blocked,     "#fee2e2", "#991b1b"],
                        ["In Progress", rq.inProgress,  "#dcfce7", "#166534"],
                        ["Partial",     rq.partial,     "#dbeafe", "#1d4ed8"],
                        ["Not Started", rq.notStarted,  "#f1f5f9", "#475569"],
                        ["Complete",    rq.complete,    "#dbeafe", "#1d4ed8"],
                        ["N/A",         rq.na,          "#f1f5f9", "#7e22ce"],
                      ].filter(([,count]) => count > 0).map(([lbl, count, bg, col]) => (
                        <div key={lbl}
                          onClick={()=>{ const br=rq.rows.filter(r=>{const v=(r[req.keys?.funcBuildStatus]||r[req.keys?.techBuildStatus]||"").toString().toLowerCase();return v.includes(lbl.toLowerCase().slice(0,4));}); if(br.length) setStoryModal({title:`${comp} — ${lbl}`, rows:br}); }}
                          style={{ display:"flex", alignItems:"center", gap:7, cursor:"pointer" }}>
                          <div style={{ flex:1, background:"#e2e8f0", borderRadius:2, height:6, overflow:"hidden" }}>
                            <div style={{ width:`${rq.total>0?Math.round((count/rq.total)*100):0}%`, height:"100%", background:col, borderRadius:2 }} />
                          </div>
                          <span style={{ fontSize:10, color:col, fontWeight:600, minWidth:14, textAlign:"right" }}>{count}</span>
                          <span style={{ fontSize:10, color:C.muted, minWidth:60 }}>{lbl}</span>
                        </div>
                      ))}
                    </div>
                  ) : <span style={{ color:C.muted, fontSize:11 }}>No story data</span>}
                </div>

              </div>
            </div>
          );
        })}
      </div>

      {/* Modals */}
      {raidModal   && <RaidDrillModal  title={raidModal.title}  rows={raidModal.rows}  raidKeys={raid?.keys} onClose={()=>setRaidModal(null)} />}
      {storyModal  && <StoryDrillModal title={storyModal.title} rows={storyModal.rows} reqKeys={req?.keys}   onClose={()=>setStoryModal(null)} />}
      {wpDrillModal && <WorkplanDrillModal title={wpDrillModal.title} rows={wpDrillModal.rows} onClose={()=>setWpDrillModal(null)} />}
    </div>
  );
}

// ─── TEST SCENARIOS TAB ──────────────────────────────────────────────────────
function TestScenariosTab({ data, wp }) {
  const [modal, setModal] = useState(null);
  const [expandedSits, setExpandedSits] = useState({});

  if (!data) return <Empty label="Upload Test Scenarios file above to view this tab." />;

  const K = data.keys;

  const isReviewed = s => {
    const v = (String(s || "").replace(/^\d+\.\s*/, "").toLowerCase());
    return v.includes("reviewed") && !v.includes("ready") && !v.includes("updated") && !v.includes("not applicable");
  };

  // ── Parse SIT workplan data ───────────────────────────────────────────────
  const getSitWorkplanData = () => {
    if (!wp) return {};
    const result = {};
    const normLvl3 = s => String(s || "").replace(/\s+/g, " ").trim().toLowerCase();
    const extractSitNum = s => { const m = normLvl3(s).match(/test scenarios for sit\s*(\d+)/); return m ? m[1] : null; };
    const matchTask = (taskName, patterns) => patterns.some(p => taskName.toLowerCase().includes(p.toLowerCase()));
    const sitRows = {};
    wp.allRows.forEach(r => {
      const sitNum = extractSitNum(String(r["Activity Grp - Lvl 3"] || "").trim());
      if (!sitNum || Number(r["Lvl"] || 0) !== 4) return;
      const ch = r["Children"]; if (ch && Number(ch) !== 0) return;
      if (!sitRows[sitNum]) sitRows[sitNum] = [];
      sitRows[sitNum].push(r);
    });
    Object.entries(sitRows).forEach(([sitNum, rows]) => {
      const getTask = (...p) => rows.find(r => matchTask(String(r["Task Name"] || ""), p));
      const ti = r => r ? { finish: r["Finish"] || r["End Date"], start: r["Start"], status: String(r["Default Status"] || r["Status"] || "") } : null;
      const sdRow = getTask("SD Consulting Internal Review", "sd consulting");
      result[`SIT ${sitNum}`] = {
        sdConsulting: ti(sdRow),
        pmSd:    ti(getTask("PM SD Review")),
        dt:      ti(getTask("DT Review")),
        da:      ti(getTask("D&A Review")),
        pmTalent:ti(getTask("PM Talent Review")),
        signOff: ti(getTask("Sign Off", "Finalize and sign off")),
        funcTechDue: (sdRow && sdRow["Start"]) ? sdRow["Start"] : null,
      };
    });
    return result;
  };
  const sitWpData = getSitWorkplanData();

  const TEAMS = [
    { key: K.funcStatus,  label: "Func",    wpKey: "funcTech",    isFunc: true  },
    { key: K.techStatus,  label: "Tech",    wpKey: "funcTech",    isFunc: true  },
    { key: K.sdStatus,    label: "SD",      wpKey: "sdConsulting",isFunc: false },
    { key: K.dtStatus,    label: "DT",      wpKey: "dt",          isFunc: false },
    { key: K.daStatus,    label: "D&A",     wpKey: "da",          isFunc: false },
    { key: K.pmtStatus,   label: "PMT SD",  wpKey: "pmSd",        isFunc: false },
    { key: K.pmStatus,    label: "PM",      wpKey: "pmTalent",    isFunc: false },
  ].filter(t => t.key);

  // ── Build sub-process rows ────────────────────────────────────────────────
  const subprocessMap = {};
  data.activeRows.forEach(r => {
    const sp = String(r[K.subprocess] || "Unknown").trim();
    if (!subprocessMap[sp]) subprocessMap[sp] = [];
    subprocessMap[sp].push(r);
  });

  const getRowSits = rows => {
    const sits = new Set();
    rows.forEach(r => {
      const v = String(r[K.sitPlan] || "").trim();
      if (v && v !== "None" && v !== "nan")
        v.split(/\n|,/).map(s => s.trim()).filter(Boolean).forEach(s => sits.add(s));
    });
    return Array.from(sits).sort((a,b) => (parseInt(a.replace(/\D/g,""))||99)-(parseInt(b.replace(/\D/g,""))||99));
  };

  const spList = Object.entries(subprocessMap).map(([sp, spRows]) => {
    const sits = getRowSits(spRows);
    const sortedSits = [...sits].sort((a,b)=>(parseInt(a.replace(/\D/g,""))||99)-(parseInt(b.replace(/\D/g,""))||99));
    const primarySit = sortedSits[0] || "Unknown";
    const getWpTask = wpKey => { for(const s of sortedSits){ const v=sitWpData[s]?.[wpKey]; if(v!=null)return v; } return null; };

    const teamData = Object.fromEntries(TEAMS.map(t => {
      const rc = spRows.filter(r => isReviewed(r[t.key])).length;
      const pct = spRows.length > 0 ? rc/spRows.length : 0;
      const dueDate = t.isFunc ? getWpTask("funcTechDue") : getWpTask(t.wpKey)?.finish;
      const dLeft = dueDate != null ? daysUntil(dueDate) : null;
      // RAG logic
      const rag = rc === spRows.length && spRows.length > 0 ? "complete"
                : (rc === 0 && (dLeft == null || dLeft > 7)) ? "notStarted"
                : (pct >= 0.5 && (dLeft == null || dLeft > 7)) ? "onTrack"
                : (pct < 0.5 && dLeft != null && dLeft <= 7) ? "atRisk"
                : "onTrack";
      return [t.key, { rc, pct, dueDate, dLeft, rag }];
    }));

    const rags = TEAMS.map(t => teamData[t.key].rag);
    const overallRag = rags.includes("atRisk") ? "atRisk"
                     : rags.every(r => r === "complete") ? "complete"
                     : rags.some(r => r === "onTrack") ? "onTrack"
                     : "notStarted";

    return { sp, spRows, sits, primarySit, total: spRows.length, estCases: Math.round(spRows.reduce((s,r)=>s+(Number(r[K.estCases])||0),0)), teamData, overallRag };
  });

  // Group by primary SIT
  const bySit = {};
  spList.forEach(row => {
    if (!bySit[row.primarySit]) bySit[row.primarySit] = [];
    bySit[row.primarySit].push(row);
  });
  const sitGroups = Object.entries(bySit).sort((a,b)=>(parseInt(a[0].replace(/\D/g,""))||99)-(parseInt(b[0].replace(/\D/g,""))||99));

  const RAG_COLORS = {
    complete:   { bg:"#dcfce7", col:"#166534", label:"Complete" },
    onTrack:    { bg:"#fef9e7", col:"#b45309", label:"On Track" },
    atRisk:     { bg:"#fee2e2", col:"#991b1b", label:"At Risk"  },
    notStarted: { bg:"#f1f5f9", col:"#64748b", label:"Not Started" },
  };

  const isExpanded = sit => expandedSits[sit] !== false;
  const toggleSit = sit => setExpandedSits(p => ({...p, [sit]: !isExpanded(sit)}));

  // KPIs
  const kpiFuncRev = data.activeRows.filter(r => isReviewed(r[K.funcStatus])).length;
  const kpiReady   = data.activeRows.filter(r => { const s=(String(r[K.funcStatus]||"").replace(/^\d+\.\s*/,"")).toLowerCase(); return s.includes("ready")||s.includes("updated"); }).length;
  const kpiNotAppl = data.activeRows.filter(r => String(r[K.funcStatus]||"").replace(/^\d+\.\s*/,"").toLowerCase().includes("not applicable")).length;
  const totCases   = Math.round(spList.reduce((s,r)=>s+r.estCases,0));

  const ScenarioModal = ({ title, rows: mRows, onClose }) => (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:1000, display:"flex", alignItems:"center", justifyContent:"center" }} onClick={onClose}>
      <div style={{ background:C.white, borderRadius:10, width:"98%", maxWidth:1200, maxHeight:"90vh", display:"flex", flexDirection:"column", boxShadow:"0 24px 60px rgba(0,0,0,0.35)" }} onClick={e=>e.stopPropagation()}>
        <div style={{ background:C.headerBg, padding:"12px 20px", borderRadius:"10px 10px 0 0", display:"flex", justifyContent:"space-between", alignItems:"center", flexShrink:0 }}>
          <div style={{ color:"#fff", fontWeight:700, fontSize:13 }}>{title} <span style={{ opacity:.6, fontWeight:400, marginLeft:8 }}>({mRows.length})</span></div>
          <button onClick={onClose} style={{ background:"rgba(255,255,255,0.15)", border:"none", color:"#fff", borderRadius:5, padding:"5px 14px", cursor:"pointer", fontSize:13, fontWeight:600 }}>✕</button>
        </div>
        <div style={{ overflowY:"auto", flex:1 }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
            <thead style={{ position:"sticky", top:0, background:"#f0f4f8", zIndex:2 }}>
              <tr style={{ borderBottom:`2px solid ${C.border}` }}>
                {["ID","Scenario","Est. Cases","Sprint","SIT Plan","Func","Tech","SD","DT","D&A","PMT SD","PM"].map(h => (
                  <th key={h} style={{ padding:"8px 10px", textAlign:h==="ID"||h==="Scenario"?"left":"center", color:C.muted, fontWeight:700, whiteSpace:"nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {mRows.map((r,i) => (
                <tr key={i} style={{ background:i%2===0?C.white:"#f9fafb", borderBottom:`1px solid ${C.border}`, verticalAlign:"top" }}>
                  <td style={{ padding:"7px 10px", color:C.muted, fontWeight:600, whiteSpace:"nowrap" }}>{String(r[K.id]||"—")}</td>
                  <td style={{ padding:"7px 10px", color:C.text, maxWidth:260, wordBreak:"break-word" }}>{String(r[K.name]||"—")}</td>
                  <td style={{ padding:"7px 10px", textAlign:"center", fontWeight:700, color:C.navyLight }}>{r[K.estCases]||"—"}</td>
                  <td style={{ padding:"7px 10px", textAlign:"center", color:C.muted, whiteSpace:"nowrap" }}>
                    {String(r[K.sprintPlan]||"—").split("\n").map(s=>s.replace(/^\d+\.\s*/,"").match(/s\d+/i)?.[0]||"").filter(Boolean).join(", ")||"—"}
                  </td>
                  <td style={{ padding:"7px 10px", textAlign:"center", color:C.muted }}>{String(r[K.sitPlan]||"—").replace(/^\d+\.\s*/,"")}</td>
                  {TEAMS.map(t => {
                    const s = String(r[t.key]||"").replace(/^\d+\.\s*/,"").trim() || null;
                    const rv = isReviewed(r[t.key]);
                    const col = rv ? "#166534" : "#64748b";
                    const bg  = rv ? "#dcfce7"  : "#f1f5f9";
                    return <td key={t.key} style={{ padding:"7px 8px", textAlign:"center" }}>
                      {s ? <span style={{ background:bg, color:col, borderRadius:4, padding:"2px 6px", fontSize:10, fontWeight:600, whiteSpace:"nowrap" }}>{s}</span>
                         : <span style={{ color:C.muted }}>—</span>}
                    </td>;
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>

      {/* KPI tiles */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(130px,1fr))", gap:10 }}>
        {[
          ["Total Scenarios", data.total,    C.navyLight, () => setModal({ title:"All Scenarios", rows:data.activeRows })],
          ["Est. Test Cases", totCases,       C.navyLight, null],
          ["Func Reviewed",   kpiFuncRev,    "#166534",   () => setModal({ title:"Functionally Reviewed", rows:data.activeRows.filter(r=>isReviewed(r[K.funcStatus])) })],
          ["Ready for Review",kpiReady,      C.gold,      () => setModal({ title:"Ready for Review", rows:data.activeRows.filter(r=>{ const s=String(r[K.funcStatus]||"").replace(/^\d+\.\s*/,"").toLowerCase(); return s.includes("ready")||s.includes("updated"); }) })],
          ["Not Applicable",  kpiNotAppl,    "#6b21a8",   () => setModal({ title:"Not Applicable", rows:data.activeRows.filter(r=>String(r[K.funcStatus]||"").replace(/^\d+\.\s*/,"").toLowerCase().includes("not applicable")) })],
          ["Sub Processes",   spList.length, C.navyLight, null],
        ].map(([lbl, val, col, onClick]) => (
          <div key={lbl} onClick={onClick}
            onMouseEnter={e=>{ if(onClick) e.currentTarget.style.boxShadow="0 3px 10px rgba(0,0,0,0.10)"; }}
            onMouseLeave={e=>{ e.currentTarget.style.boxShadow="none"; }}
            style={{ background:C.white, border:`1px solid ${C.border}`, borderTop:`3px solid ${col}`, borderRadius:7, padding:"10px 12px", cursor:onClick?"pointer":"default" }}>
            <div style={{ fontSize:9, color:C.muted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:3 }}>{lbl}</div>
            <div style={{ fontSize:22, fontWeight:800, color:col, lineHeight:1 }}>{val}</div>
            {onClick && <div style={{ fontSize:9, color:C.accent, marginTop:3 }}>Details →</div>}
          </div>
        ))}
      </div>

      {/* Legend */}
      <div style={{ display:"flex", gap:14, flexWrap:"wrap", alignItems:"center", padding:"8px 14px", background:C.white, border:`1px solid ${C.border}`, borderRadius:7, fontSize:10 }}>
        <span style={{ color:C.muted, fontWeight:600 }}>Review status:</span>
        {Object.entries(RAG_COLORS).map(([k,{bg,col,label}]) => (
          <span key={k} style={{ display:"flex", alignItems:"center", gap:5 }}>
            <span style={{ width:18, height:18, borderRadius:4, background:bg, border:`1px solid ${col}30`, display:"inline-flex", alignItems:"center", justifyContent:"center" }}>
              <span style={{ width:8, height:8, borderRadius:2, background:col, display:"inline-block" }} />
            </span>
            <span style={{ color:C.text }}>{label}</span>
          </span>
        ))}
        {wp && <span style={{ color:C.muted, marginLeft:8 }}>Date shown = target from Workplan · Func &amp; Tech due = SD Consulting start</span>}
      </div>

      {/* Grouped by SIT */}
      {sitGroups.map(([sit, rows]) => {
        const sitData = sitWpData[sit];
        const signOffDate = sitData?.signOff?.finish;
        const signOffDl = signOffDate ? daysUntil(signOffDate) : null;
        const signOffCol = signOffDl == null ? C.muted : signOffDl < 0 ? C.delayed : signOffDl <= 14 ? C.gold : "#166534";
        const sitRags = rows.map(r => r.overallRag);
        const sitOverall = sitRags.includes("atRisk") ? "atRisk"
                         : sitRags.every(r => r === "complete") ? "complete"
                         : sitRags.some(r => r === "onTrack") ? "onTrack"
                         : "notStarted";
        const sitRagC = RAG_COLORS[sitOverall];
        const open = isExpanded(sit);
        const sitTotScenarios = rows.reduce((s,r)=>s+r.total,0);
        const sitTotCases = rows.reduce((s,r)=>s+r.estCases,0);

        return (
          <Card key={sit} style={{ padding:0 }}>
            {/* SIT header — clickable to expand/collapse */}
            <div onClick={()=>toggleSit(sit)}
              style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"11px 16px", cursor:"pointer", background:"#f8fafc", borderRadius:open?"10px 10px 0 0":10, borderBottom: open?`1px solid ${C.border}`:"none" }}
              onMouseEnter={e=>e.currentTarget.style.background="#eef4ff"}
              onMouseLeave={e=>e.currentTarget.style.background="#f8fafc"}>
              <div style={{ display:"flex", alignItems:"center", gap:12 }}>
                <span style={{ fontSize:14, userSelect:"none", color:C.muted }}>{open?"▼":"▶"}</span>
                <span style={{ fontWeight:700, color:C.navy, fontSize:13 }}>{sit}</span>
                <span style={{ background:C.navyLight+"20", color:C.navyLight, borderRadius:10, padding:"2px 9px", fontSize:10, fontWeight:700 }}>{rows.length} sub processes · {sitTotScenarios} scenarios · {sitTotCases} est. cases</span>
                {signOffDate && (
                  <span style={{ fontSize:11, color:signOffCol, fontWeight:600 }}>
                    Sign-off: {fmtDate(signOffDate)}{signOffDl!=null ? ` (${signOffDl<0?Math.abs(signOffDl)+"d over":signOffDl+"d"})` : ""}
                  </span>
                )}
              </div>
              <span style={{ background:sitRagC.bg, color:sitRagC.col, border:`1px solid ${sitRagC.col}40`, borderRadius:4, padding:"3px 10px", fontSize:10, fontWeight:700 }}>{sitRagC.label}</span>
            </div>

            {open && (
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
                  <thead>
                    <tr style={{ background:"#162f50" }}>
                      <th style={{ padding:"8px 12px", textAlign:"left", color:"#fff", fontWeight:700, fontSize:11, minWidth:170, borderRight:`1px solid rgba(255,255,255,0.1)` }}>Sub Process</th>
                      <th style={{ padding:"8px 10px", textAlign:"center", color:"#a8d8ff", fontWeight:700, fontSize:10, borderRight:`1px solid rgba(255,255,255,0.1)` }}>Scenarios</th>
                      {TEAMS.map(t => (
                        <th key={t.key} style={{ padding:"8px 8px", textAlign:"center", color:"#c4f1c4", fontWeight:700, fontSize:10, borderRight:`1px solid rgba(255,255,255,0.1)`, minWidth:80 }}>{t.label}</th>
                      ))}
                      <th style={{ padding:"8px 10px", textAlign:"center", color:"#fff", fontWeight:700, fontSize:10 }}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {/* SIT total row */}
                    <tr style={{ background:"#1a3a5c", borderBottom:`1px solid rgba(255,255,255,0.1)` }}>
                      <td style={{ padding:"7px 12px", color:"#fff", fontWeight:800, fontSize:11 }}>TOTAL</td>
                      <td style={{ padding:"7px 10px", textAlign:"center", color:"#a8d8ff", fontWeight:800 }}>{sitTotScenarios}</td>
                      {TEAMS.map(t => {
                        const tot = rows.reduce((s,r)=>s+(r.teamData[t.key]?.rc||0),0);
                        const allTot = rows.reduce((s,r)=>s+r.total,0);
                        const allRev = allTot > 0 && tot === allTot;
                        return <td key={t.key} style={{ padding:"7px 8px", textAlign:"center", color: allRev ? "#86efac" : "#c4f1c4", fontWeight:800 }}>{tot}</td>;
                      })}
                      <td style={{ padding:"7px 10px", textAlign:"center", color:"rgba(255,255,255,0.3)" }}>—</td>
                    </tr>

                    {/* Sub process rows */}
                    {rows.map((row, i) => {
                      const ragC = RAG_COLORS[row.overallRag];
                      return (
                        <tr key={row.sp}
                          style={{ background:i%2===0?C.white:"#f7f9fc", borderBottom:`1px solid ${C.border}`, cursor:"pointer" }}
                          onClick={()=>setModal({ title:row.sp, rows:row.spRows })}
                          onMouseEnter={e=>e.currentTarget.style.background="#eef4ff"}
                          onMouseLeave={e=>e.currentTarget.style.background=i%2===0?C.white:"#f7f9fc"}>

                          {/* Name */}
                          <td style={{ padding:"9px 12px", color:C.text, fontWeight:600, fontSize:11 }}>
                            <span style={{ display:"flex", alignItems:"center", gap:6 }}>
                              <span style={{ width:3, height:14, background:ragC.col, borderRadius:2, flexShrink:0 }} />
                              {row.sp}
                            </span>
                          </td>

                          {/* Scenario count */}
                          <td style={{ padding:"9px 10px", textAlign:"center", color:C.text, fontWeight:700 }}>{row.total}</td>

                          {/* Heatmap cells per team */}
                          {TEAMS.map(t => {
                            const { rc, rag, dueDate } = row.teamData[t.key];
                            const rc2 = RAG_COLORS[rag];
                            const dl = dueDate != null ? daysUntil(dueDate) : null;
                            const dateStr = dueDate != null ? fmtDate(dueDate) : null;
                            return (
                              <td key={t.key} style={{ padding:"6px 4px", textAlign:"center" }}>
                                <div style={{ display:"inline-flex", flexDirection:"column", alignItems:"center", gap:2,
                                  background:rc2.bg, borderRadius:6, padding:"4px 8px", minWidth:52,
                                  border:`1px solid ${rc2.col}30` }}>
                                  <span style={{ fontWeight:800, fontSize:14, color:rc2.col, lineHeight:1 }}>{rc}</span>
                                  {dateStr && <span style={{ fontSize:9, color:rc2.col, opacity:.75, whiteSpace:"nowrap" }}>{dateStr}</span>}
                                </div>
                              </td>
                            );
                          })}

                          {/* Status pill */}
                          <td style={{ padding:"9px 10px", textAlign:"center" }}>
                            <span style={{ background:ragC.bg, color:ragC.col, border:`1px solid ${ragC.col}40`, borderRadius:4, padding:"2px 8px", fontSize:10, fontWeight:700 }}>{ragC.label}</span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </Card>
        );
      })}

      {modal && <ScenarioModal title={modal.title} rows={modal.rows} onClose={()=>setModal(null)} />}
    </div>
  );
}


// ─── SCORECARD CLASSIC TAB (dark navy header, refined) ───────────────────────
function ScClassicWpPill({ status, onClick }) {
  if (!status || status === "—") return <span style={{ color: C.muted }}>—</span>;
  const sl = status.toLowerCase();
  const bg  = sl.includes("off track") ? "#fee2e2" : sl.includes("on track") ? "#fef9e7" : sl.includes("complete") ? "#dbeafe" : "#f1f5f9";
  const col = sl.includes("off track") ? "#b91c1c" : sl.includes("on track") ? "#b45309" : sl.includes("complete") ? "#1d4ed8" : "#64748b";
  const bdr = sl.includes("off track") ? "#fca5a5" : sl.includes("on track") ? "#fcd34d" : sl.includes("complete") ? "#93c5fd" : "#cbd5e1";
  return (
    <span onClick={onClick}
      style={{ background:bg, color:col, border:`1px solid ${bdr}`, borderRadius:4,
        padding:"2px 8px", fontSize:10, fontWeight:600,
        cursor: onClick ? "pointer" : "default", whiteSpace:"nowrap" }}>
      {status}
    </span>
  );
}

function ScClassicBuildPill({ dist }) {
  if (!dist || !Object.keys(dist).length) return <span style={{ color: C.muted }}>—</span>;
  const sorted = Object.entries(dist).sort((a, b) => b[1] - a[1]);
  const top = sorted[0][0].toLowerCase();
  const bg  = top.includes("block") ? "#fee2e2" : (top.includes("complete") && !top.includes("partial")) ? "#dbeafe" : (top.includes("progress") || top.includes("partial")) ? "#dcfce7" : top.includes("not start") ? "#f1f5f9" : "#fef3c7";
  const col = top.includes("block") ? "#991b1b" : (top.includes("complete") && !top.includes("partial")) ? "#1d4ed8" : (top.includes("progress") || top.includes("partial")) ? "#166534" : top.includes("not start") ? "#475569" : "#92400e";
  return (
    <span style={{ background:bg, color:col, borderRadius:4, padding:"2px 8px", fontSize:10, fontWeight:600, whiteSpace:"nowrap" }}>
      {sorted[0][0]}
    </span>
  );
}

function ScorecardClassicTab({ wp, raid, req, openModal }) {
  const [raidModal,    setRaidModal]    = useState(null);
  const [storyModal,   setStoryModal]   = useState(null);
  const [wpDrillModal, setWpDrillModal] = useState(null);
  if (!raid && !req && !wp) return <Empty label="Upload files to view Scorecard Classic." />;

  // ── Aliases ───────────────────────────────────────────────────────────────
  const COMP_ALIASES = {
    "carr": "Career Advancement Review",
    "career advancement reviiew": "Career Advancement Review",
    "career advancement review": "Career Advancement Review",
    "career advancement review (carr)": "Career Advancement Review",
    "career advancement readiness review": "Career Advancement Review",
    "career advancement readiness review (carr)": "Career Advancement Review",
    "expectation framework": "Expectations Framework",
    "expectations framework": "Expectations Framework",
  };
  const normaliseComp = name => {
    const key = String(name || "").toLowerCase().trim();
    if (COMP_ALIASES[key]) return COMP_ALIASES[key];
    for (const [alias, canonical] of Object.entries(COMP_ALIASES)) {
      if (key.includes(alias) || alias.includes(key)) return canonical;
    }
    return name;
  };

  // ── RAID helper ───────────────────────────────────────────────────────────
  const getCompRaid = compName => {
    if (!raid) return { open:0, delayed:0, issues:[], risks:[], openItems:[] };
    const norm = normaliseComp(compName);
    const items = raid.items.filter(r => normaliseComp(String(r[raid.keys.component] || "")) === norm);
    const open    = items.filter(r => { const s = String(r[raid.keys.status] || "").toLowerCase(); return s !== "complete" && s !== "deferred"; });
    const delayed = items.filter(r => String(r[raid.keys.status] || "").toLowerCase() === "delayed");
    return {
      open: open.length, delayed: delayed.length,
      issues:    open.filter(r => String(r[raid.keys.type] || "").toLowerCase().includes("issue")),
      risks:     open.filter(r => String(r[raid.keys.type] || "").toLowerCase().includes("risk")),
      openItems: open,
    };
  };

  // ── Workplan helper ───────────────────────────────────────────────────────
  const wpWorstStatus = rows => {
    let worst = null, wr = -1;
    rows.forEach(r => {
      const sl = String(r["Default Status"] || r["Status"] || "").toLowerCase();
      const rk = sl.includes("off track") ? 4 : sl.includes("on track") ? 3 : sl.includes("not start") ? 2 : sl.includes("complete") ? 1 : 0;
      if (rk > wr) { wr = rk; worst = String(r["Default Status"] || r["Status"] || ""); }
    });
    return worst || "—";
  };

  const getCompWp = compName => {
    if (!wp) return null;
    const scopedRows = wp.allRows.filter(r =>
      String(r["Activity Grp - Lvl 1"] || "").trim() === "Technology - SAP Configuration & Build" &&
      String(r["Activity Grp - Lvl 2"] || "").trim() === "Component Build"
    );
    // Try direct normalised match first, then sub-process mapping
    const normComp = normaliseComp(compName);
    const wpName = SUB_PROCESS_TO_WP[compName.toLowerCase().trim()] || normComp;
    const lvl3Rows = scopedRows.filter(r => {
      const lvl3 = normaliseComp(String(r["Activity Grp - Lvl 3"] || "").trim());
      return lvl3 === normComp || lvl3 === wpName || normaliseComp(lvl3) === normaliseComp(wpName);
    });
    if (!lvl3Rows.length) return null;
    const lvl3Names = Array.from(new Set(lvl3Rows.map(r => String(r["Activity Grp - Lvl 3"] || "").trim())));
    const subtreeRows = scopedRows.filter(r => lvl3Names.includes(String(r["Activity Grp - Lvl 3"] || "").trim()));
    const isLeafRow = r => { const c = r["Children"]; return !c || Number(c) === 0; };
    const lvl4Plus = subtreeRows.filter(r => Number(r["Lvl"] ?? 0) >= 4);
    const isDesign = r => /design/i.test(String(r["Task Name"] || "") + String(r["Activity Grp - Lvl 4"] || ""));
    const isBuild  = r => /build|develop|implement|code/i.test(String(r["Task Name"] || "") + String(r["Activity Grp - Lvl 4"] || ""));
    const dLeaves = lvl4Plus.filter(r => isLeafRow(r) && isDesign(r));
    const bLeaves = lvl4Plus.filter(r => isLeafRow(r) && isBuild(r));
    const leafRows = subtreeRows.filter(isLeafRow);
    const pctVals = leafRows.map(r => {
      const v = r["% Complete"] ?? r["% complete"];
      const s2 = String(v ?? "").replace("%","").trim();
      if (s2 !== "" && !isNaN(Number(s2))) { const n = Number(s2); return n <= 1 ? Math.round(n * 100) : Math.round(n); }
      const s = String(r["Default Status"] || r["Status"] || "").toLowerCase();
      if (s.includes("complete")) return 100;
      if (s.includes("on track") || s.includes("in progress")) return 50;
      if (s.includes("off track") || s.includes("delayed")) return 25;
      if (s.includes("not start")) return 0;
      return null;
    }).filter(v => v != null);
    const pct = pctVals.length ? Math.round(pctVals.reduce((a, b) => a + b, 0) / pctVals.length) : null;
    const dGroups = Array.from(new Set(lvl4Plus.filter(isDesign).map(r => String(r["Activity Grp - Lvl 4"] || "")).filter(Boolean)));
    const bGroups = Array.from(new Set(lvl4Plus.filter(isBuild).map(r => String(r["Activity Grp - Lvl 4"] || "")).filter(Boolean)));
    const makeDrill = groups => {
      if (!groups.length) return [];
      const hdrs = subtreeRows.filter(r => Number(r["Lvl"] ?? 0) === 3);
      const sub  = subtreeRows.filter(r => groups.includes(String(r["Activity Grp - Lvl 4"] || "").trim()));
      const seen = new Set();
      return [...hdrs, ...sub].filter(r => { const id = r["Row ID"] || JSON.stringify(r); if (seen.has(id)) return false; seen.add(id); return true; });
    };
    return { designStatus: dLeaves.length ? wpWorstStatus(dLeaves) : null, buildStatus: bLeaves.length ? wpWorstStatus(bLeaves) : null, pctComplete: pct, designRows: makeDrill(dGroups), buildRows: makeDrill(bGroups) };
  };

  // ── Requirements helper ───────────────────────────────────────────────────
  const getCompReq = compName => {
    if (!req || !req.byComponent) return null;
    const normName = normaliseComp(compName).toLowerCase();
    const keys = Object.keys(req.byComponent);

    // 1. Exact normalised match
    let key = keys.find(k => normaliseComp(k).toLowerCase() === normName);

    // 2. Partial match — comp name contains sub process name or vice versa
    if (!key) key = keys.find(k => {
      const nk = normaliseComp(k).toLowerCase();
      return normName.includes(nk) || nk.includes(normName);
    });

    // 3. Word overlap — share at least 2 significant words
    if (!key) {
      const stopWords = new Set(["the","and","for","of","in","a","an","to","from","with","by"]);
      const nameWords = normName.split(/\s+/).filter(w => w.length > 2 && !stopWords.has(w));
      key = keys.find(k => {
        const kWords = normaliseComp(k).toLowerCase().split(/\s+/).filter(w => w.length > 2 && !stopWords.has(w));
        const overlap = nameWords.filter(w => kWords.some(kw => kw.includes(w) || w.includes(kw)));
        return overlap.length >= 2;
      });
    }

    if (!key) return null;
    const cd = req.byComponent[key];
    const sprintData = req.compBySprint ? (req.compBySprint[key] || {}) : {};
    const bs = req.compBuildStatus ? (req.compBuildStatus[key] || null) : null;
    return { total:cd.total, complete:cd.complete, partial:cd.partial, inProgress:cd.inProgress,
      notStarted:cd.notStarted, blocked:cd.blocked, na:cd.na||0,
      sprintData, funcDist: bs ? bs.func : {}, techDist: bs ? bs.tech : {}, rows: cd.rows };
  };

  // ── Sprint order ──────────────────────────────────────────────────────────
  const rawSprintOrder = (req?.sprintOrder || []).filter(s => s && s !== "No Sprint");
  const sprintLabelMap = {};
  rawSprintOrder.forEach(sp => {
    const m = String(sp).toLowerCase().match(/^\s*(\d+)\.\s*s(\d+)/);
    if (m) { const n = parseInt(m[2]); sprintLabelMap[sp] = (n >= 1 && n <= 8) ? `S${n}` : null; }
    else { const nm = String(sp).toLowerCase().match(/s(\d+)/); if (nm) { const n = parseInt(nm[1]); sprintLabelMap[sp] = (n >= 1 && n <= 8) ? `S${n}` : null; } else sprintLabelMap[sp] = null; }
  });
  const sprintOrder = [];
  const seenLbls = new Set();
  ["S1","S2","S3","S4","S5","S6","S7","S8"].forEach(lbl => {
    const raws = rawSprintOrder.filter(sp => sprintLabelMap[sp] === lbl);
    if (raws.length > 0 && !seenLbls.has(lbl)) { sprintOrder.push({ label: lbl, raws }); seenLbls.add(lbl); }
  });
  const naSprintRaws = rawSprintOrder.filter(sp => /not.?applicable|n\/a for tech/i.test(sp));
  const hasNaCol = naSprintRaws.length > 0;

  const getSD = (sprintData, entry) => {
    const c = { complete:0, partial:0, inProgress:0, notStarted:0, blocked:0, na:0, total:0 };
    entry.raws.forEach(raw => { const d = sprintData[raw]; if (d) Object.keys(c).forEach(k => { c[k] += (d[k] || 0); }); });
    return c.total > 0 ? c : null;
  };
  const getNaData = sprintData => {
    const c = { complete:0, partial:0, inProgress:0, notStarted:0, blocked:0, na:0, total:0 };
    naSprintRaws.forEach(raw => { const d = sprintData[raw]; if (d) Object.keys(c).forEach(k => { c[k] += (d[k] || 0); }); });
    return c.total > 0 ? c : null;
  };

  // ── Component list ────────────────────────────────────────────────────────
  const raidComps = raid ? Array.from(new Set(raid.items.map(r => normaliseComp(String(r[raid.keys.component] || ""))).filter(Boolean))).sort() : [];
  const reqComps  = req  ? Array.from(new Set(Object.keys(req.byComponent || {}).map(normaliseComp))).sort() : [];
  const wpComps   = wp   ? Array.from(new Set(
    wp.allRows.filter(r => String(r["Activity Grp - Lvl 1"] || "").trim() === "Technology - SAP Configuration & Build" && String(r["Activity Grp - Lvl 2"] || "").trim() === "Component Build")
      .map(r => normaliseComp(String(r["Activity Grp - Lvl 3"] || "").trim())).filter(Boolean)
  )).sort() : [];
  const allComps = Array.from(new Set([...raidComps, ...reqComps, ...wpComps])).sort();
  const visComps = allComps.filter(comp => {
    const rc = getCompRaid(comp); const rq = getCompReq(comp); const cw = getCompWp(comp);
    return rc.open > 0 || rc.delayed > 0 || (rq && rq.total > 0) || cw;
  });

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <Card style={{ padding:0 }}>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
            <thead>
              <tr style={{ background:"#0a1f3d", borderBottom:"2px solid #2563eb" }}>
                <th style={{ padding:"8px 12px", color:"#fff", fontWeight:700, fontSize:10, textAlign:"left", minWidth:170, position:"sticky", left:0, background:"#0f2744", zIndex:2 }}>Component</th>
                <th colSpan={4} style={{ padding:"8px", color:"#fbbf24", fontWeight:700, fontSize:10, textAlign:"center", borderLeft:"1px solid rgba(255,255,255,0.2)" }}>RAID</th>
                {req && sprintOrder.map(e => <th key={e.label} style={{ padding:"8px 6px", color:"#7dd3fc", fontWeight:700, fontSize:10, textAlign:"center", minWidth:52 }}>{e.label}</th>)}
                {req && hasNaCol && <th style={{ padding:"8px 5px", color:"#7dd3fc", fontWeight:700, fontSize:9, textAlign:"center", minWidth:70, borderLeft:"1px solid rgba(255,255,255,0.15)" }}>N/A<br/>Tech</th>}
                {req && <th style={{ padding:"8px 6px", color:"#7dd3fc", fontWeight:700, fontSize:10, textAlign:"center", minWidth:46, borderLeft:"1px solid rgba(255,255,255,0.2)" }}>Total</th>}
                {req && <th style={{ padding:"8px 10px", color:"#86efac", fontWeight:700, fontSize:10, textAlign:"center", minWidth:120, borderLeft:"1px solid rgba(255,255,255,0.2)" }}>Func Build</th>}
                {req && <th style={{ padding:"8px 10px", color:"#86efac", fontWeight:700, fontSize:10, textAlign:"center", minWidth:120 }}>Tech Build</th>}
                {wp  && <th colSpan={3} style={{ padding:"8px 10px", color:"#fdba74", fontWeight:700, fontSize:10, textAlign:"center", minWidth:270, borderLeft:"1px solid rgba(255,255,255,0.2)" }}>Workplan</th>}
              </tr>
              <tr style={{ background:"#162f50" }}>
                <th style={{ padding:"5px 12px", color:"rgba(255,255,255,0.4)", fontSize:9, textAlign:"left", position:"sticky", left:0, background:"#162f50", zIndex:2 }} />
                <th style={{ padding:"5px 8px", color:"#fbbf24", fontSize:9, textAlign:"center", borderLeft:"1px solid rgba(255,255,255,0.1)" }}>Open</th>
                <th style={{ padding:"5px 8px", color:"#fbbf24", fontSize:9, textAlign:"center" }}>Delayed</th>
                <th style={{ padding:"5px 8px", color:"#fbbf24", fontSize:9, textAlign:"center" }}>Issues</th>
                <th style={{ padding:"5px 8px", color:"#fbbf24", fontSize:9, textAlign:"center", borderRight:"1px solid rgba(255,255,255,0.15)" }}>Risks</th>
                {req && sprintOrder.map(e => <th key={e.label} style={{ padding:"5px 8px", color:"rgba(255,255,255,0.45)", fontSize:9, textAlign:"center" }}>stories</th>)}
                {req && hasNaCol && <th style={{ padding:"5px 6px", color:"rgba(255,255,255,0.45)", fontSize:9, textAlign:"center", borderLeft:"1px solid rgba(255,255,255,0.1)" }}>stories</th>}
                {req && <th style={{ padding:"5px 8px", color:"rgba(255,255,255,0.45)", fontSize:9, textAlign:"center", borderLeft:"1px solid rgba(255,255,255,0.15)" }}>active</th>}
                {req && <th style={{ padding:"5px 8px", color:"#86efac", fontSize:9, textAlign:"center", borderLeft:"1px solid rgba(255,255,255,0.15)" }}>consolidated</th>}
                {req && <th style={{ padding:"5px 8px", color:"#86efac", fontSize:9, textAlign:"center" }}>consolidated</th>}
                {wp  && <th style={{ padding:"5px 8px", color:"#fdba74", fontSize:9, textAlign:"center", borderLeft:"1px solid rgba(255,255,255,0.2)" }}>Design</th>}
                {wp  && <th style={{ padding:"5px 8px", color:"#fdba74", fontSize:9, textAlign:"center" }}>Build</th>}
                {wp  && <th style={{ padding:"5px 8px", color:"#fdba74", fontSize:9, textAlign:"center" }}>% Complete</th>}
              </tr>
            </thead>
            <tbody>
              {/* TOTAL row */}
              {(() => {
                const tRaid = { open:0, delayed:0, issues:0, risks:0 };
                const tSprints = {};
                const tNa = { complete:0, partial:0, inProgress:0, notStarted:0, blocked:0, na:0, total:0 };
                let tTotal = 0;
                visComps.forEach(comp => {
                  const rc = getCompRaid(comp); const rq = getCompReq(comp);
                  tRaid.open += rc.open; tRaid.delayed += rc.delayed; tRaid.issues += rc.issues.length; tRaid.risks += rc.risks.length;
                  if (rq) {
                    tTotal += rq.total;
                    sprintOrder.forEach(e => {
                      const sd = getSD(rq.sprintData, e);
                      if (sd) { if (!tSprints[e.label]) tSprints[e.label] = { complete:0,partial:0,inProgress:0,notStarted:0,blocked:0,na:0,total:0 }; Object.keys(tSprints[e.label]).forEach(k => { tSprints[e.label][k] += (sd[k]||0); }); }
                    });
                    if (hasNaCol) { const nd = getNaData(rq.sprintData); if (nd) Object.keys(tNa).forEach(k => { tNa[k] += (nd[k]||0); }); }
                  }
                });
                const TH = ({ children, style }) => <td style={{ padding:"7px 8px", textAlign:"center", background:"#162f50", color:"#fff", fontWeight:800, fontSize:12, ...style }}>{children}</td>;
                return (
                  <tr style={{ borderBottom:`2px solid ${C.navyLight}`, position:"sticky", top:0, zIndex:3 }}>
                    <td style={{ padding:"7px 12px", background:"#162f50", color:"#fff", fontWeight:800, fontSize:12, position:"sticky", left:0, zIndex:4 }}>TOTAL ({visComps.length})</td>
                    <TH>{tRaid.open || "—"}</TH>
                    <TH>{tRaid.delayed || "—"}</TH>
                    <TH>{tRaid.issues || "—"}</TH>
                    <TH style={{ borderRight:"1px solid rgba(255,255,255,0.25)" }}>{tRaid.risks || "—"}</TH>
                    {req && sprintOrder.map(e => {
                      const sd = tSprints[e.label];
                      const { bg, color } = sprintBubbleColor(sd);
                      return <td key={e.label} style={{ padding:"7px 6px", textAlign:"center", background:"#162f50" }}>
                        {sd ? <span style={{ background:bg, color, borderRadius:4, padding:"2px 8px", fontSize:11, fontWeight:700 }}>{sd.total}</span> : <span style={{ color:"rgba(255,255,255,0.3)" }}>—</span>}
                      </td>;
                    })}
                    {req && hasNaCol && <td style={{ padding:"7px 6px", textAlign:"center", background:"#162f50", borderLeft:"1px solid rgba(255,255,255,0.15)" }}>
                      {tNa.total > 0 ? <span style={{ background:"#a855f720", color:"#e9d5ff", borderRadius:4, padding:"2px 8px", fontSize:11, fontWeight:700 }}>{tNa.total}</span> : <span style={{ color:"rgba(255,255,255,0.3)" }}>—</span>}
                    </td>}
                    {req && <TH style={{ borderLeft:"1px solid rgba(255,255,255,0.25)" }}>{tTotal || "—"}</TH>}
                    {req && <TH style={{ borderLeft:"1px solid rgba(255,255,255,0.25)" }}>—</TH>}
                    {req && <TH>—</TH>}
                    {wp  && <TH style={{ borderLeft:"1px solid rgba(255,255,255,0.25)" }}>—</TH>}
                    {wp  && <TH>—</TH>}
                    {wp  && <TH>—</TH>}
                  </tr>
                );
              })()}

              {/* Component rows */}
              {visComps.map((comp, i) => {
                const rc = getCompRaid(comp);
                const rq = getCompReq(comp);
                const cw = getCompWp(comp);
                const isDelayed = rc.delayed > 0;
                const rowBg = isDelayed ? "#fff5f5" : i % 2 === 0 ? C.white : "#f7f9fc";
                const delayedItems = raid ? raid.items.filter(r => normaliseComp(String(r[raid.keys.component] || "")) === normaliseComp(comp) && String(r[raid.keys.status] || "").toLowerCase() === "delayed") : [];

                return (
                  <tr key={comp} style={{ background:rowBg, borderBottom:`1px solid ${isDelayed ? "#fca5a5" : C.border}`, verticalAlign:"middle" }}>

                    {/* Name */}
                    <td style={{ padding:"8px 12px", color:C.navy, fontWeight:600, fontSize:11, position:"sticky", left:0, background:rowBg, zIndex:1, maxWidth:170, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={comp}>
                      {comp}
                    </td>

                    {/* RAID: Open */}
                    <td style={{ padding:"7px 8px", textAlign:"center" }}>
                      {rc.open > 0
                        ? <span onClick={() => setRaidModal({ title:`${comp} — Open RAIDs`, rows:rc.openItems })} style={{ cursor:"pointer", color:C.navyLight, fontWeight:700 }}>{rc.open}</span>
                        : <span style={{ color:C.muted }}>—</span>}
                    </td>

                    {/* RAID: Delayed */}
                    <td style={{ padding:"7px 8px", textAlign:"center" }}>
                      {rc.delayed > 0
                        ? <span onClick={() => setRaidModal({ title:`${comp} — Delayed`, rows:delayedItems })} style={{ background:C.delayed+"20", color:C.delayed, border:`1px solid ${C.delayed}40`, borderRadius:4, padding:"2px 7px", fontSize:10, fontWeight:700, cursor:"pointer" }}>⚠ {rc.delayed}</span>
                        : <span style={{ color:C.muted }}>—</span>}
                    </td>

                    {/* RAID: Issues */}
                    <td style={{ padding:"7px 8px", textAlign:"center" }}>
                      {rc.issues.length > 0
                        ? <span onClick={() => setRaidModal({ title:`${comp} — Issues`, rows:rc.issues })} style={{ cursor:"pointer", color:C.delayed, fontWeight:600 }}>{rc.issues.length}</span>
                        : <span style={{ color:C.muted }}>—</span>}
                    </td>

                    {/* RAID: Risks */}
                    <td style={{ padding:"7px 8px", textAlign:"center", borderRight:`1px solid ${C.border}` }}>
                      {rc.risks.length > 0
                        ? <span onClick={() => setRaidModal({ title:`${comp} — Risks`, rows:rc.risks })} style={{ cursor:"pointer", color:C.gold, fontWeight:600 }}>{rc.risks.length}</span>
                        : <span style={{ color:C.muted }}>—</span>}
                    </td>

                    {/* Sprint bubbles */}
                    {req && sprintOrder.map(e => {
                      const sd = rq ? getSD(rq.sprintData, e) : null;
                      const { bg, color } = sprintBubbleColor(sd);
                      const sprintRows = rq ? rq.rows.filter(r => e.raws.some(raw => String(r[req.keys?.sprint] || "") === raw)) : [];
                      return (
                        <td key={e.label} style={{ padding:"7px 6px", textAlign:"center" }}>
                          {sd
                            ? <span onClick={() => sprintRows.length && setStoryModal({ title:`${comp} — ${e.label}`, rows:sprintRows })} style={{ background:bg, color, borderRadius:5, padding:"3px 9px", fontSize:11, fontWeight:700, cursor:sprintRows.length?"pointer":"default", display:"inline-block" }}>{sd.total}</span>
                            : <span style={{ color:C.muted }}>—</span>}
                        </td>
                      );
                    })}

                    {/* N/A col */}
                    {req && hasNaCol && (() => {
                      const nd = rq ? getNaData(rq.sprintData) : null;
                      const naRows = rq ? rq.rows.filter(r => naSprintRaws.some(raw => String(r[req.keys?.sprint] || "") === raw)) : [];
                      return (
                        <td style={{ padding:"7px 6px", textAlign:"center", borderLeft:`1px solid ${C.border}` }}>
                          {nd
                            ? <span onClick={() => naRows.length && setStoryModal({ title:`${comp} — Not Applicable`, rows:naRows })} style={{ background:"#f3e8ff", color:"#6b21a8", borderRadius:5, padding:"3px 8px", fontSize:11, fontWeight:700, cursor:naRows.length?"pointer":"default" }}>{nd.total}</span>
                            : <span style={{ color:C.muted }}>—</span>}
                        </td>
                      );
                    })()}

                    {/* Total */}
                    {req && (
                      <td style={{ padding:"7px 8px", textAlign:"center", fontWeight:700, color:C.navyLight, borderLeft:`1px solid ${C.border}` }}>
                        {rq
                          ? <span onClick={() => setStoryModal({ title:`All Stories — ${comp}`, rows:rq.rows })} style={{ cursor:"pointer" }}>{rq.total}</span>
                          : <span style={{ color:C.muted }}>—</span>}
                      </td>
                    )}

                    {/* Func Build */}
                    {req && (
                      <td style={{ padding:"7px 10px", textAlign:"center", borderLeft:`1px solid ${C.border}` }}>
                        {rq ? <ScClassicBuildPill dist={rq.funcDist} /> : <span style={{ color:C.muted }}>—</span>}
                      </td>
                    )}

                    {/* Tech Build */}
                    {req && (
                      <td style={{ padding:"7px 10px", textAlign:"center" }}>
                        {rq ? <ScClassicBuildPill dist={rq.techDist} /> : <span style={{ color:C.muted }}>—</span>}
                      </td>
                    )}

                    {/* Design Status */}
                    {wp && (
                      <td style={{ padding:"7px 10px", textAlign:"center", borderLeft:`1px solid ${C.border}` }}>
                        {cw ? <ScClassicWpPill status={cw.designStatus} onClick={cw.designRows.length ? () => setWpDrillModal({ title:`${comp} — Design`, rows:cw.designRows }) : null} /> : <span style={{ color:C.muted }}>—</span>}
                      </td>
                    )}

                    {/* Build Status */}
                    {wp && (
                      <td style={{ padding:"7px 10px", textAlign:"center" }}>
                        {cw ? <ScClassicWpPill status={cw.buildStatus} onClick={cw.buildRows.length ? () => setWpDrillModal({ title:`${comp} — Build`, rows:cw.buildRows }) : null} /> : <span style={{ color:C.muted }}>—</span>}
                      </td>
                    )}

                    {/* % Complete with inline bar */}
                    {wp && (
                      <td style={{ padding:"7px 12px", minWidth:110 }}>
                        {cw && cw.pctComplete != null ? (
                          <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                            <div style={{ flex:1, background:"#e2e8f0", borderRadius:3, height:6, overflow:"hidden" }}>
                              <div style={{ width:`${cw.pctComplete}%`, height:"100%", background: cw.pctComplete >= 75 ? C.complete : cw.pctComplete >= 40 ? C.gold : C.delayed, borderRadius:3 }} />
                            </div>
                            <span style={{ fontSize:11, fontWeight:700, color:C.text, minWidth:32 }}>{cw.pctComplete}%</span>
                          </div>
                        ) : <span style={{ color:C.muted }}>—</span>}
                      </td>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      {raidModal    && <RaidDrillModal    title={raidModal.title}    rows={raidModal.rows}    raidKeys={raid?.keys} onClose={() => setRaidModal(null)} />}
      {storyModal   && <StoryDrillModal   title={storyModal.title}   rows={storyModal.rows}   reqKeys={req?.keys}   onClose={() => setStoryModal(null)} />}
      {wpDrillModal && <WorkplanDrillModal title={wpDrillModal.title} rows={wpDrillModal.rows}                       onClose={() => setWpDrillModal(null)} />}
    </div>
  );
}

function ScorecardTab({ wp, raid, req, openModal }) {
  const [raidModal, setRaidModal] = useState(null);
  const [storyModal, setStoryModal] = useState(null);
  const [wpDrillModal, setWpDrillModal] = useState(null);
  const [modalColConfig, setModalColConfig] = useState({
    raidId:    { label:"RAID ID",              visible:true,  width:90  },
    status:    { label:"Status",               visible:true,  width:90  },
    type:      { label:"Type",                 visible:true,  width:90  },
    component: { label:"Component",            visible:true,  width:130 },
    experience:{ label:"Experience",           visible:true,  width:90  },
    topic:     { label:"Topic",                visible:true,  width:90  },
    desc:      { label:"Description",          visible:true,  width:260 },
    comment:   { label:"Comments / Resolution",visible:true,  width:220 },
    owner:     { label:"Owner",                visible:true,  width:110 },
    team:      { label:"Primary Team (Owner)", visible:true,  width:140 },
    critPath:  { label:"Critical Path",        visible:true,  width:100 },
    dueDate:   { label:"Due Date",             visible:true,  width:85  },
  });
  if (!raid && !req && !wp) return <Empty label="Upload files to view Component Scorecard." />;

  // ── Component name aliases ────────────────────────────────────────────────
  // Map known abbreviations / alternate spellings to a single canonical name.
  // Keys are lowercase; value is the canonical display name.
  const COMP_ALIASES = {
    "carr":                                       "Career Advancement Review",
    "career advancement reviiew":                 "Career Advancement Review",
    "career advancement review":                  "Career Advancement Review",
    "career advancement review (carr)":           "Career Advancement Review",
    "career advancement readiness review":        "Career Advancement Review",
    "career advancement readiness review (carr)": "Career Advancement Review",
    "expectation framework":                      "Expectations Framework",
    "expectations framework":                     "Expectations Framework",
  };

  // Mapping from Requirements Sub Process names → Workplan Lvl3 component names
  // Used so getCompWp can find workplan data when given a sub process name
  const SUB_PROCESS_TO_WP = {
    "performance assessment":               "Performance Assessment",
    "firm contribution assessment":         "Firm Contribution Assessment",
    "360 insights":                         "360 Insights",
    "year end review":                      "Year End Review",
    "career advancement readiness review":  "Career Advancement Review",
    "expectations framework":               "Expectations Framework",
    "individual dashboard":                 "Individual Dashboard",
    "performance management dashboard":     "Performance Management Dashboard",
    "engagement leader dashboard":          "Engagement Leader Dashboard",
    "coach dashboard":                      "Coach Dashboard",
    "metrics":                              "Metrics",
    "notifications":                        "Notifications",
    "reports":                              "Reports",
    "foundation data model":                "Foundation Data Model",
    "rbp":                                  "RBP",
  };
  const normaliseComp = (name) => {
    const key = String(name || "").toLowerCase().trim();
    if (COMP_ALIASES[key]) return COMP_ALIASES[key];
    // Partial fallback: any alias key that is a substring of key, or vice versa
    for (const [alias, canonical] of Object.entries(COMP_ALIASES)) {
      if (key.includes(alias) || alias.includes(key)) return canonical;
    }
    return name;
  };

  // ── RAID helpers ──────────────────────────────────────────────────────────
  const getCompRaid = (compName) => {
    if (!raid) return { open:0, delayed:0, issues:[], risks:[], openItems:[] };
    const normComp = normaliseComp(compName);
    const items = raid.items.filter(r => {
      const c = normaliseComp(String(r[raid.keys.component]||""));
      return c === normComp;
    });
    const open    = items.filter(r => { const s = String(r[raid.keys.status]||"").toLowerCase(); return s !== "complete" && s !== "deferred"; });
    const delayed = items.filter(r => String(r[raid.keys.status]||"").toLowerCase() === "delayed");
    const issues  = open.filter(r => String(r[raid.keys.type]||"").toLowerCase().includes("issue"));
    const risks   = open.filter(r => String(r[raid.keys.type]||"").toLowerCase().includes("risk"));
    return { open: open.length, delayed: delayed.length, issues, risks, openItems: open };
  };

  // ── Workplan helpers ──────────────────────────────────────────────────────
  // Match comp name against Activity Grp - Lvl 3, then roll up children
  // to derive Design Status and Build Status from task names.
  // Priority order for worst-status rollup: Off Track > On Track > Not Started > Complete
  const WP_STATUS_RANK = { "off track": 4, "on track": 3, "not started": 2, "complete": 1 };
  const wpWorstStatus = (rows) => {
    let worst = null; let worstRank = -1;
    rows.forEach(r => {
      const s = String(r["Default Status"] || r["Status"] || "").trim();
      const sl = s.toLowerCase();
      const rank = sl.includes("off track") ? 4 : sl.includes("on track") ? 3 : sl.includes("not start") ? 2 : sl.includes("complete") ? 1 : 0;
      if (rank > worstRank) { worstRank = rank; worst = s; }
    });
    return worst || "—";
  };

  const getCompWp = (compName) => {
    if (!wp) return null;
    const cn = compName.toLowerCase().trim();
    const cnWords = cn.split(/[\s\-\/,]+/).filter(w => w.length > 2);

    // Find Lvl 3 rows whose name fuzzy-matches the component,
    // scoped to Lvl 1 = "Technology - SAP Configuration & Build" / Lvl 2 = "Component Build"
    const scopedRows = wp.allRows.filter(r =>
      String(r["Activity Grp - Lvl 1"] || "").trim() === "Technology - SAP Configuration & Build" &&
      String(r["Activity Grp - Lvl 2"] || "").trim() === "Component Build"
    );
    const lvl3Rows = scopedRows.filter(r => {
      const lvl3Normalised = normaliseComp(String(r["Activity Grp - Lvl 3"] || "").trim());
      const normComp = normaliseComp(compName);
      // Strict: only match if the normalised Lvl 3 value equals the normalised component name
      return lvl3Normalised === normComp;
    });
    if (!lvl3Rows.length) return null;

    // Unique matched Lvl 3 names
    const lvl3Names = Array.from(new Set(lvl3Rows.map(r => String(r["Activity Grp - Lvl 3"] || "").trim())));

    // All rows that belong to this Lvl 3 (Lvl 3 header rows + all Lvl 4+ descendants)
    // A row belongs if its Activity Grp - Lvl 3 is in our matched set
    const subtreeRows = scopedRows.filter(r => {
      const lvl3 = String(r["Activity Grp - Lvl 3"] || "").trim();
      return lvl3Names.includes(lvl3);
    });

    // For Lvl 4+ rows only — used for design/build keyword split
    const lvl4PlusRows = subtreeRows.filter(r => Number(r["Lvl"] ?? 0) >= 4);

    // Separate design vs build by scanning Task Name + Lvl 4 group name
    const isDesign = r => /design/i.test(String(r["Task Name"] || "") + String(r["Activity Grp - Lvl 4"] || ""));
    const isBuild  = r => /build|develop|implement|code/i.test(String(r["Task Name"] || "") + String(r["Activity Grp - Lvl 4"] || ""));

    // For design/build subtree drills: include Lvl 4 group headers + their children
    // Identify Lvl 4 group names that are design / build flavoured
    const designLvl4Groups = Array.from(new Set(
      lvl4PlusRows.filter(isDesign).map(r => String(r["Activity Grp - Lvl 4"] || "")).filter(Boolean)
    ));
    const buildLvl4Groups = Array.from(new Set(
      lvl4PlusRows.filter(isBuild).map(r => String(r["Activity Grp - Lvl 4"] || "")).filter(Boolean)
    ));

    // Build a helper that, given a set of Lvl4 group names, returns:
    //   - the Lvl 3 header row(s)  [so the hierarchy root is visible]
    //   - all rows whose Lvl 4 group is in the set  [the design/build subtree]
    // Rows are returned in original file order (preserved from wp.allRows).
    const buildDrillRows = (lvl4Groups) => {
      if (!lvl4Groups.length) return [];
      const lvl3HeaderRows = subtreeRows.filter(r => Number(r["Lvl"] ?? 0) === 3);
      const subtreeFiltered = subtreeRows.filter(r => {
        const l4 = String(r["Activity Grp - Lvl 4"] || "").trim();
        return lvl4Groups.includes(l4) || (Number(r["Lvl"] ?? 0) >= 4 && (isDesign(r) || isBuild(r)) && lvl4Groups.some(g => String(r["Activity Grp - Lvl 4"] || "").trim() === g));
      });
      // Deduplicate and preserve file order
      const seen = new Set();
      return [...lvl3HeaderRows, ...subtreeFiltered].filter(r => {
        const id = r["Row ID"] || JSON.stringify(r);
        if (seen.has(id)) return false;
        seen.add(id);
        return true;
      });
    };

    const designDrillRows = buildDrillRows(designLvl4Groups);
    const buildDrillRows_ = buildDrillRows(buildLvl4Groups);

    // For status/% calcs keep using the keyword-filtered subtrees (leaf rows only)
    const designSubtree = subtreeRows.filter(r => {
      const l4 = String(r["Activity Grp - Lvl 4"] || "").trim();
      return designLvl4Groups.includes(l4) || isDesign(r);
    });
    const buildSubtree = subtreeRows.filter(r => {
      const l4 = String(r["Activity Grp - Lvl 4"] || "").trim();
      return buildLvl4Groups.includes(l4) || isBuild(r);
    });

    // Worst status for pill colour — leaf rows only (ignore group header rows)
    const isLeafRow = r => { const c = r["Children"]; return !c || Number(c) === 0; };
    const designLeaves = designSubtree.filter(isLeafRow);
    const buildLeaves  = buildSubtree.filter(isLeafRow);
    const designStatus = designLeaves.length ? wpWorstStatus(designLeaves) : null;
    const buildStatus  = buildLeaves.length  ? wpWorstStatus(buildLeaves)  : null;

    // % Complete — prefer the Lvl 3 header row value (same source as Workplan tab)
    // Fall back to average of all leaf rows when header has no value
    const normPct = v => {
      const s = String(v ?? "").replace("%","").trim();
      if (!s || isNaN(Number(s))) return null;
      const n = Number(s); return n <= 1 ? Math.round(n * 100) : Math.round(n);
    };
    const lvl3HeaderRow = subtreeRows.find(r => Number(r["Lvl"] ?? 0) === 3);
    const headerPct = lvl3HeaderRow ? normPct(lvl3HeaderRow["% Complete"] ?? lvl3HeaderRow["% complete"]) : null;
    const leafRows = subtreeRows.filter(isLeafRow);
    const pctValues = leafRows.map(r => {
        const pv = normPct(r["% Complete"] ?? r["% complete"]);
        if (pv != null) return pv;
        const s = String(r["Default Status"] || r["Status"] || "").toLowerCase();
        if (s.includes("complete")) return 100;
        if (s.includes("on track") || s.includes("in progress")) return 50;
        if (s.includes("off track") || s.includes("delayed")) return 25;
        if (s.includes("not start")) return 0;
        return null;
      }).filter(v => v != null);
    const leafAvg = pctValues.length ? Math.round(pctValues.reduce((a,b) => a+b,0) / pctValues.length) : null;
    const pctComplete = headerPct != null ? headerPct : leafAvg;

    // If header shows 100% or all leaves complete, treat Build as Complete when
    // keyword split found no build-tagged tasks (avoids "Not started" on done components)
    const allLeavesComplete = leafRows.length > 0 && leafRows.every(r => {
      const s = String(r["Default Status"] || r["Status"] || "").toLowerCase();
      return s.includes("complete") || normPct(r["% Complete"] ?? r["% complete"]) === 100;
    });
    const resolvedBuildStatus = buildStatus ?? ((headerPct === 100 || allLeavesComplete) ? "Complete" : null);

    return {
      designStatus,
      buildStatus: resolvedBuildStatus,
      pctComplete,
      designRows:     designDrillRows,
      buildRows:      buildDrillRows_,
      allRows:        subtreeRows,
    };
  };

  // Helper: colour-code a workplan status string
  const wpStatusPill = (status, rows, label, onClick) => {
    if (!status || status === "—") return <span style={{ color:C.muted }}>—</span>;
    const sl = status.toLowerCase();
    const bg    = sl.includes("off track") ? "#fee2e2" : sl.includes("on track") ? "#fef9e7" : sl.includes("complete") ? "#dbeafe" : "#f1f5f9";
    const color = sl.includes("off track") ? "#b91c1c" : sl.includes("on track") ? "#b45309" : sl.includes("complete") ? "#1d4ed8" : "#64748b";
    const border= sl.includes("off track") ? "#fca5a5" : sl.includes("on track") ? "#fcd34d" : sl.includes("complete") ? "#93c5fd" : "#cbd5e1";
    return (
      <span onClick={onClick} title={label}
        style={{ background:bg, color, border:`1px solid ${border}`, borderRadius:4,
          padding:"2px 8px", fontSize:10, fontWeight:700, display:"inline-block",
          cursor: onClick ? "pointer" : "default", whiteSpace:"nowrap" }}>
        {status}
      </span>
    );
  };

  // ── Req helpers ───────────────────────────────────────────────────────────
  // Consolidated Status logic (mirrors Smartsheet formula):
  // Count stories where Status=Partial AND NOT Deprecated AND NOT Deferred
  // grouped by Build Cycle (sprint), using Func Build + Tech Build status values
  const getConsolidatedStatus = (items) => {
    if (!items || !items.length) return null;
    const K = req.keys;
    // Overall func/tech status — pick the "worst" status across all stories
    // Priority: Blocked > In Progress > Partial Build Complete > Not Started > Complete > N/A
    const STATUS_RANK = { "blocked":6, "in progress":5, "partial build complete":4, "not started":3, "complete":2, "n/a":1 };
    const getRank = v => {
      const s = String(v||"").toLowerCase().trim();
      for (const [k,r] of Object.entries(STATUS_RANK)) { if (s.includes(k)) return r; }
      return 0;
    };
    let funcStatus = "", techStatus = "", maxFuncRank = -1, maxTechRank = -1;
    items.forEach(r => {
      const fb = String(r[K.funcBuildStatus]||"").trim();
      const tb = String(r[K.techBuildStatus]||"").trim();
      if (fb && getRank(fb) > maxFuncRank) { maxFuncRank = getRank(fb); funcStatus = fb; }
      if (tb && getRank(tb) > maxTechRank) { maxTechRank = getRank(tb); techStatus = tb; }
    });
    return { funcStatus, techStatus };
  };

  const getCompReq = (compName) => {
    if (!req || !req.byComponent) return null;
    const normName = normaliseComp(compName);
    // Exact match on normalised name only
    let key = Object.keys(req.byComponent).find(k => normaliseComp(k) === normName);
    if (!key) return null;

    const cd = req.byComponent[key];
    const sprintData = req.compBySprint ? (req.compBySprint[key] || {}) : {};
    const bs = req.compBuildStatus ? (req.compBuildStatus[key] || null) : null;
    const consolidated = getConsolidatedStatus(cd.rows);

    return {
      total: cd.total,
      complete: cd.complete, partial: cd.partial,
      inProgress: cd.inProgress, notStarted: cd.notStarted,
      blocked: cd.blocked, na: cd.na || 0,
      sprintData,   // { sprintName: { complete, partial, inProgress, notStarted, blocked, na, total } }
      funcDist: bs ? bs.func : {},
      techDist: bs ? bs.tech : {},
      consolidated,
      rows: cd.rows
    };
  };

  // ── Build component list ──────────────────────────────────────────────────
  const raidComps = raid ? Array.from(new Set(raid.items.map(r => normaliseComp(String(r[raid.keys.component]||""))).filter(Boolean))).sort() : [];
  const reqComps  = req  ? Array.from(new Set(Object.keys(req.byComponent || {}).map(normaliseComp))).sort() : [];
  const wpComps   = wp   ? Array.from(new Set(
    wp.allRows
      .filter(r =>
        String(r["Activity Grp - Lvl 1"] || "").trim() === "Technology - SAP Configuration & Build" &&
        String(r["Activity Grp - Lvl 2"] || "").trim() === "Component Build"
      )
      .map(r => normaliseComp(String(r["Activity Grp - Lvl 3"] || "").trim()))
      .filter(Boolean)
  )).sort() : [];
  const allComps  = Array.from(new Set([...raidComps, ...reqComps, ...wpComps])).sort();

  // Sprint columns — from req sprintOrder, simplified to S1-S8 labels
  const rawSprintOrder = req && req.sprintOrder && req.sprintOrder.length > 0
    ? req.sprintOrder.filter(s => s && s !== "No Sprint")
    : [];

  // Map raw sprint names to short labels S1-S8, detect N/A sprint buckets
  const sprintLabelMap = {}; // rawName -> "S1"..."S8" or "na"
  const NA_SPRINT_PATTERNS = /not.?applicable|n\/a for tech|column20/i;

  rawSprintOrder.forEach(sp => {
    const s = String(sp).toLowerCase();
    // Match patterns like "1. S1...", "2. S2...", etc.
    const m = s.match(/^\s*(\d+)\.\s*s(\d+)/);
    if (m) {
      const num = parseInt(m[2]);
      if (num >= 1 && num <= 8) sprintLabelMap[sp] = `S${num}`;
      else sprintLabelMap[sp] = null; // S9+ excluded
    } else if (NA_SPRINT_PATTERNS.test(sp)) {
      sprintLabelMap[sp] = "na";
    } else {
      // fallback: try to extract just a number
      const nm = s.match(/s(\d+)/);
      if (nm) {
        const num = parseInt(nm[1]);
        if (num >= 1 && num <= 8) sprintLabelMap[sp] = `S${num}`;
        else sprintLabelMap[sp] = null;
      } else {
        sprintLabelMap[sp] = null; // exclude unknown
      }
    }
  });

  // Final sprint columns: S1-S8 in order (deduplicated by label)
  const sprintOrder = [];
  const seenLabels = new Set();
  ["S1","S2","S3","S4","S5","S6","S7","S8"].forEach(lbl => {
    // Find raw sprint(s) that map to this label
    const raws = rawSprintOrder.filter(sp => sprintLabelMap[sp] === lbl);
    if (raws.length > 0 && !seenLabels.has(lbl)) {
      sprintOrder.push({ label: lbl, raws });
      seenLabels.add(lbl);
    }
  });

  // N/A technical build sprints — combine "10. Not Applicable..." and "7. Not applicable..."
  // Also check funcBuildStatus/techBuildStatus values that contain "not applicable"
  const naSprintRaws = rawSprintOrder.filter(sp => sprintLabelMap[sp] === "na");
  const hasNaCol = naSprintRaws.length > 0;

  // Helper: get story count for a sprint label (sum across all raw sprint values for that label)
  const getSprintData = (sprintData, labelEntry) => {
    const combined = { complete:0, partial:0, inProgress:0, notStarted:0, blocked:0, na:0, total:0 };
    labelEntry.raws.forEach(raw => {
      const d = sprintData[raw];
      if (d) { Object.keys(combined).forEach(k => { combined[k] += (d[k]||0); }); }
    });
    return combined.total > 0 ? combined : null;
  };

  const getNaData = (sprintData) => {
    const combined = { complete:0, partial:0, inProgress:0, notStarted:0, blocked:0, na:0, total:0 };
    naSprintRaws.forEach(raw => {
      const d = sprintData[raw];
      if (d) { Object.keys(combined).forEach(k => { combined[k] += (d[k]||0); }); }
    });
    return combined.total > 0 ? combined : null;
  };

  const getNaRows = (allRows) => {
    if (!allRows || !req) return [];
    return allRows.filter(r => naSprintRaws.includes(String(r[req.keys.sprint]||"").trim()));
  };

  // Status colors
  const statusColor = s => {
    const v = String(s||"").toLowerCase();
    if (v.includes("complete") && !v.includes("partial")) return "#1d4ed8"; // blue
    if (v.includes("partial"))   return "#0369a1";                          // light blue
    if (v.includes("progress"))  return "#15803d";                          // green
    if (v.includes("block"))     return "#b91c1c";                          // red
    if (v.includes("n/a") || v === "na") return "#64748b";                  // grey
    return "#64748b";                                                        // grey (not started)
  };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>

      {/* ── KPI tiles ── */}
      {(() => {
        // Use pre-computed bucket counts from parseRequirements (already handles fallback to Status column)
        const compData    = req ? Object.values(req.byComponent || {}) : [];
        const allCompRows = compData.flatMap(d => d.rows || []);

        // statusBucket with fallback: func/tech build status → Status column
        const statusBucket = r => {
          const fb = String(r[req?.keys?.funcBuildStatus] || "").toLowerCase();
          const tb = String(r[req?.keys?.techBuildStatus] || "").toLowerCase();
          const rank = s => s.includes("block") ? 6 : s.includes("progress") ? 5 : s.includes("partial") ? 4 : s.includes("not start") ? 3 : s.includes("complete") ? 2 : s.includes("n/a") ? 1 : 0;
          const w = Math.max(rank(fb), rank(tb));
          if (w > 0) return w === 6 ? "blocked" : w === 5 ? "inProgress" : w === 4 ? "partial" : w === 3 ? "notStarted" : w === 2 ? "complete" : "na";
          // Fallback to the generic Status column (mirrors parseRequirements logic)
          const sv = String(r[req?.keys?.status] || "").toLowerCase();
          return sv.includes("block") ? "blocked" : sv.includes("progress") ? "inProgress" : sv.includes("partial") ? "partial" : sv.includes("complete") ? "complete" : "notStarted";
        };

        // Counts from pre-computed byComponent (most reliable — already uses same fallback)
        const preCount = {
          blocked:    compData.reduce((s,d)=>s+(d.blocked||0),0),
          inProgress: compData.reduce((s,d)=>s+(d.inProgress||0),0),
          partial:    compData.reduce((s,d)=>s+(d.partial||0),0),
          complete:   compData.reduce((s,d)=>s+(d.complete||0),0),
        };
        // Row arrays for drill-down modals (filtered using statusBucket with fallback)
        const storyRows = {
          all:        allCompRows,
          blocked:    allCompRows.filter(r => statusBucket(r) === "blocked"),
          inProgress: allCompRows.filter(r => statusBucket(r) === "inProgress"),
          partial:    allCompRows.filter(r => statusBucket(r) === "partial"),
          complete:   allCompRows.filter(r => statusBucket(r) === "complete"),
        };

        // Pre-compute RAID rows grouped by component for drill-down
        const raidOpenRows    = raid ? raid.open    : [];
        const raidDelayedRows = raid ? raid.delayed : [];

        const tiles = [
          {
            lbl: "Components",
            val: allComps.filter(c => { const rc=getCompRaid(c); const rq=getCompReq(c); const cw=getCompWp(c); return rc.open>0||rc.delayed>0||(rq&&rq.total>0)||cw; }).length,
            col: C.navyLight, onClick: null,
          },
          { lbl:"Total Stories", val: storyRows.all.length || "—",   col: C.navyLight, onClick: storyRows.all.length       ? () => setStoryModal({ title:"All Stories",           rows: storyRows.all        }) : null },
          { lbl:"Blocked",       val: preCount.blocked    || "—",   col: C.delayed,   onClick: preCount.blocked    > 0 ? () => setStoryModal({ title:"Blocked Stories",      rows: storyRows.blocked    }) : null },
          { lbl:"In Progress",   val: preCount.inProgress || "—",   col: C.onTrack,   onClick: preCount.inProgress > 0 ? () => setStoryModal({ title:"In Progress Stories",  rows: storyRows.inProgress }) : null },
          { lbl:"Partial",       val: preCount.partial    || "—",   col: "#0369a1",   onClick: preCount.partial    > 0 ? () => setStoryModal({ title:"Partial Build Stories", rows: storyRows.partial    }) : null },
          { lbl:"Complete",      val: preCount.complete   || "—",   col: C.complete,  onClick: preCount.complete   > 0 ? () => setStoryModal({ title:"Complete Stories",     rows: storyRows.complete   }) : null },
          { lbl:"Open RAIDs",    val: raidOpenRows.length || "—",         col: C.gold,      onClick: raidOpenRows.length    ? () => setRaidModal({ title:"Open RAIDs — All Components",    rows: raidOpenRows    }) : null },
          { lbl:"Delayed RAIDs", val: raidDelayedRows.length || "—",      col: C.delayed,   onClick: raidDelayedRows.length ? () => setRaidModal({ title:"Delayed RAIDs — All Components", rows: raidDelayedRows }) : null },
        ];

        return (
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(120px,1fr))", gap:10 }}>
            {tiles.map(({ lbl, val, col, onClick }) => (
              <div key={lbl} onClick={onClick}
                onMouseEnter={e => { if (onClick) e.currentTarget.style.boxShadow = "0 3px 10px rgba(0,0,0,0.10)"; }}
                onMouseLeave={e => { e.currentTarget.style.boxShadow = "none"; }}
                style={{ background:C.white, border:`1px solid ${C.border}`, borderTop:`3px solid ${col}`, borderRadius:7,
                  padding:"10px 12px", cursor: onClick ? "pointer" : "default", transition:"box-shadow .15s" }}>
                <div style={{ fontSize:9, color:C.muted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:3 }}>{lbl}</div>
                <div style={{ fontSize:22, fontWeight:800, color:col, lineHeight:1 }}>{val}</div>
                {onClick && <div style={{ fontSize:9, color:C.accent, marginTop:3 }}>Details →</div>}
              </div>
            ))}
          </div>
        );
      })()}

      {/* Info bar */}
      <Card style={{ padding:"10px 16px" }}>
        <div style={{ display:"flex", gap:20, flexWrap:"wrap", alignItems:"center" }}>
          <span style={{ fontSize:12, color:C.muted }}><b style={{ color:C.text }}>RAID:</b> <code style={{ background:"#f0f2f5", padding:"1px 5px", borderRadius:3 }}>Component</code> column</span>
          <span style={{ fontSize:12, color:C.muted }}><b style={{ color:C.text }}>Stories:</b> <code style={{ background:"#f0f2f5", padding:"1px 5px", borderRadius:3 }}>Sub Process</code> column · Sprint from <code style={{ background:"#f0f2f5", padding:"1px 5px", borderRadius:3 }}>Build Cycle</code> · Excl. Deprecated (5.) / Deferred (6.) via <code style={{ background:"#f0f2f5", padding:"1px 5px", borderRadius:3 }}>User Story Review Status (D&A)</code></span>
          <span style={{ fontSize:12, color:C.muted }}><b style={{ color:C.text }}>Workplan:</b> matched on <code style={{ background:"#f0f2f5", padding:"1px 5px", borderRadius:3 }}>Activity Grp - Lvl 3</code> · Design/Build status rolled up from child tasks · click pill to drill down</span>
          <div style={{ display:"flex", gap:16, flexWrap:"wrap", alignItems:"center" }}>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center" }}>
              <span style={{ fontSize:11, color:C.muted, fontWeight:600 }}>Sprint bubbles:</span>
              {[
                ["Complete",    "#1d4ed8", "#dbeafe", "#93c5fd"],
                ["Partial",     "#0369a1", "#e0f2fe", "#7dd3fc"],
                ["In Progress", "#15803d", "#dcfce7", "#86efac"],
                ["Blocked",     "#b91c1c", "#fee2e2", "#fca5a5"],
                ["Not Started", "#334155", "#f1f5f9", "#94a3b8"],
                ["N/A",         "#7e22ce", "#f3e8ff", "#d8b4fe"],
              ].map(([l, color, bg, border]) => (
                <span key={l} style={{ display:"inline-flex", alignItems:"center", gap:4,
                  background:bg, color:color, border:`1px solid ${border}`,
                  borderRadius:10, padding:"1px 8px", fontSize:11, fontWeight:500 }}>
                  {l}
                </span>
              ))}
            </div>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center" }}>
              <span style={{ fontSize:11, color:C.muted, fontWeight:600 }}>% bar colour:</span>
              {[
                ["Off Track",   C.delayed, "#fee2e2"],
                ["On Track",    C.gold,    "#fef9e7"],
                ["Complete",    C.green,   "#dcfce7"],
                ["Not Started", "#94a3b8", "#f1f5f9"],
              ].map(([l, col, bg]) => (
                <span key={l} style={{ display:"inline-flex", alignItems:"center", gap:5, fontSize:11 }}>
                  <span style={{ width:24, height:7, borderRadius:4, background:col, display:"inline-block" }} />
                  <span style={{ color:C.muted }}>{l}</span>
                </span>
              ))}
            </div>
          </div>
        </div>
      </Card>

      {/* Main table */}
      <Card style={{ padding:0 }}>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
            <thead>
              {/* Group header row */}
              <tr style={{ background:"var(--color-background-primary, #fff)", borderBottom:"0.5px solid #e2e8f0" }}>
                <th style={{ padding:"8px 10px", color:"transparent", fontSize:10, textAlign:"left", minWidth:150 }}> </th>
                <th colSpan={4} style={{ padding:"8px", background:"#fef3c7", color:"#92400e", fontWeight:700, fontSize:10, textAlign:"center", borderLeft:`2px solid #fcd34d` }}>RAID</th>
                {req && sprintOrder.length > 0 && <th colSpan={sprintOrder.length + (hasNaCol ? 1 : 0) + 1} style={{ padding:"8px", background:"#eff6ff", color:"#1e40af", fontWeight:700, fontSize:10, textAlign:"center", borderLeft:`2px solid #93c5fd` }}>User Stories</th>}
                {req && <th colSpan={2} style={{ padding:"8px", background:"#f0fdf4", color:"#166534", fontWeight:700, fontSize:10, textAlign:"center", borderLeft:`2px solid #86efac` }}>Build Status</th>}
                {wp  && <th colSpan={3} style={{ padding:"8px", background:"#fff7ed", color:"#c2410c", fontWeight:700, fontSize:10, textAlign:"center", borderLeft:`2px solid #fdba74` }}>Workplan</th>}
              </tr>
              {/* Sub-header row */}
              <tr style={{ background:"var(--color-background-secondary, #f8fafc)", borderBottom:`1.5px solid #cbd5e1` }}>
                <th style={{ padding:"6px 10px", color:C.navy, fontSize:10, fontWeight:700, textAlign:"left" }}>Component</th>
                <th style={{ padding:"5px 8px", color:"#92400e", fontSize:9, textAlign:"center" }}>Open</th>
                <th style={{ padding:"5px 8px", color:"#92400e", fontSize:9, textAlign:"center" }}>Delayed</th>
                <th style={{ padding:"5px 8px", color:"#92400e", fontSize:9, textAlign:"center" }}>Issues</th>
                <th style={{ padding:"5px 8px", color:"#92400e", fontSize:9, textAlign:"center" }}>Risks</th>
                {req && sprintOrder.map(entry => (
                  <th key={entry.label} style={{ padding:"5px 8px", color:"#1e40af", fontSize:9, textAlign:"center" }}>{entry.label}</th>
                ))}
                {req && hasNaCol && <th style={{ padding:"5px 6px", color:"#6b21a8", fontSize:9, textAlign:"center" }}>N/A</th>}
                {req && <th style={{ padding:"5px 8px", color:"#1e40af", fontSize:9, textAlign:"center" }} title="Total active stories. May exceed sum of sprint columns if some stories have no sprint assigned.">Total ⓘ</th>}
                {req && <th style={{ padding:"5px 8px", color:"#166534", fontSize:9, textAlign:"center" }}>Func</th>}
                {req && <th style={{ padding:"5px 8px", color:"#166534", fontSize:9, textAlign:"center" }}>Tech</th>}
                {wp  && <th style={{ padding:"5px 8px", color:"#c2410c", fontSize:9, textAlign:"center" }}>Design</th>}
                {wp  && <th style={{ padding:"5px 8px", color:"#c2410c", fontSize:9, textAlign:"center" }}>Build</th>}
                {wp  && <th style={{ padding:"5px 8px", color:"#c2410c", fontSize:9, textAlign:"center" }}>% Complete</th>}
              </tr>
            </thead>
            <tbody>
              {/* ── TOTAL ROW ── */}
              {(() => {
                const visComps = allComps.filter(comp => {
                  const rc = getCompRaid(comp);
                  const rq = getCompReq(comp);
                  const cw = getCompWp(comp);
                  return rc.open > 0 || rc.delayed > 0 || (rq && rq.total > 0) || (cw && (cw.designRows.length > 0 || cw.buildRows.length > 0));
                });
                const totRaid = { open:0, delayed:0, issues:0, risks:0 };
                const totSprints = {}; // { label: { complete,partial,inProgress,notStarted,blocked,na,total } }
                const totNa = { complete:0, partial:0, inProgress:0, notStarted:0, blocked:0, na:0, total:0 };
                let totTotal = 0;

                visComps.forEach(comp => {
                  const rc = getCompRaid(comp);
                  const rq = getCompReq(comp);
                  totRaid.open    += rc.open;
                  totRaid.delayed += rc.delayed;
                  totRaid.issues  += rc.issues.length;
                  totRaid.risks   += rc.risks.length;
                  if (rq) {
                    totTotal += rq.total;
                    sprintOrder.forEach(entry => {
                      const sd = getSprintData(rq.sprintData, entry);
                      if (sd) {
                        if (!totSprints[entry.label]) totSprints[entry.label] = { complete:0,partial:0,inProgress:0,notStarted:0,blocked:0,na:0,total:0 };
                        ["complete","partial","inProgress","notStarted","blocked","na","total"].forEach(k => { totSprints[entry.label][k] += (sd[k]||0); });
                      }
                    });
                    if (hasNaCol) {
                      const nd = getNaData(rq.sprintData);
                      if (nd) { ["complete","partial","inProgress","notStarted","blocked","na","total"].forEach(k => { totNa[k] += (nd[k]||0); }); }
                    }
                  }
                });

                const TH = ({ children, style }) => (
                  <td style={{ padding:"7px 8px", textAlign:"center", background:"#e8f0fe",
                    color:C.navy, fontWeight:800, fontSize:12, borderRight:`1px solid ${C.border}`,
                    ...style }}>
                    {children}
                  </td>
                );

                return (
                  <tr style={{ borderBottom:`2px solid ${C.navyLight}`, position:"sticky", top:0, zIndex:3 }}>
                    <td style={{ padding:"7px 10px", background:"#1e3a5f", color:"#fff", fontWeight:800,
                      fontSize:12, borderRight:`1px solid ${C.border}` }}>
                      TOTAL ({visComps.length} components)
                    </td>
                    {/* RAID totals */}
                    <TH>{totRaid.open || "—"}</TH>
                    <TH>{totRaid.delayed || "—"}</TH>
                    <TH>{totRaid.issues || "—"}</TH>
                    <TH style={{ borderRight:`1px solid rgba(255,255,255,0.3)` }}>{totRaid.risks || "—"}</TH>
                    {/* Sprint totals — colour-coded pills matching data rows */}
                    {req && sprintOrder.map(entry => {
                      const sd = totSprints[entry.label];
                      if (!sd || sd.total === 0) return <TH key={entry.label}>—</TH>;
                      const bc = sprintBubbleColor(sd);
                      return (
                        <td key={entry.label} style={{ padding:"7px 6px", textAlign:"center", background:"#e8f0fe",
                          fontWeight:800, fontSize:13, color:bc.color }}>
                          {sd.total}
                        </td>
                      );
                    })}
                    {/* N/A total */}
                    {req && hasNaCol && (() => {
                      const sd = totNa.total > 0 ? totNa : null;
                      const bc = sd ? sprintBubbleColor(sd) : null;
                      return (
                        <td style={{ padding:"7px 6px", textAlign:"center", background:"#e8f0fe",
                          borderLeft:`1px solid ${C.border}`, fontWeight:800, fontSize:13,
                          color: bc ? bc.color : C.muted }}>
                          {sd ? sd.total : "—"}
                        </td>
                      );
                    })()}
                    {/* Total stories */}
                    {req && (
                      <td style={{ padding:"7px 8px", textAlign:"center", background:"#e8f0fe",
                        borderLeft:`1px solid ${C.border}`, fontWeight:800, fontSize:13, color:C.navy }}>
                        {totTotal || "—"}
                      </td>
                    )}
                    {/* Func + Tech build status — blank */}
                    {req && <td style={{ background:"#e8f0fe", borderLeft:`1px solid rgba(255,255,255,0.15)` }} />}
                    {req && <td style={{ background:"#e8f0fe" }} />}
                    {/* Workplan Design + Build — blank in total row */}
                    {wp && <td style={{ background:"#e8f0fe", borderLeft:`1px solid rgba(255,255,255,0.2)` }} />}
                    {wp && <td style={{ background:"#e8f0fe" }} />}
                    {wp && <td style={{ background:"#e8f0fe" }} />}
                  </tr>
                );
              })()}

              {allComps.map((comp, i) => {
                const rc = getCompRaid(comp);
                const rq = getCompReq(comp);
                const cw = getCompWp(comp);
                const hasData = rc.open > 0 || rc.delayed > 0 || (rq && rq.total > 0) || (cw && (cw.designRows.length > 0 || cw.buildRows.length > 0));
                if (!hasData) return null;

                return (
                  <tr key={comp} style={{ background:i%2===0?C.white:"#f7f9fc", borderBottom:`1px solid ${C.border}`, verticalAlign:"middle" }}>

                    {/* Component name */}
                    <td style={{ padding:"8px 10px", color:C.text, fontWeight:600, borderRight:`1px solid ${C.border}`, whiteSpace:"nowrap", maxWidth:200, overflow:"hidden", textOverflow:"ellipsis" }} title={comp}>
                      {comp.slice(0,32)}
                    </td>

                    {/* RAID */}
                    <td style={{ padding:"6px 8px", textAlign:"center", cursor:rc.open>0?"pointer":"default" }}
                      onClick={()=>rc.open>0&&setRaidModal({title:`Open RAIDs — ${comp}`, rows:rc.openItems})}>
                      {rc.open>0
                        ? <span style={{ fontSize:12, fontWeight:800, background:"#fef3c7", color:"#b45309",
                            border:"2px solid #fcd34d", borderRadius:3, padding:"2px 10px",
                            display:"inline-block", minWidth:28, textAlign:"center", lineHeight:"18px" }}>
                            {rc.open}
                          </span>
                        : <span style={{ color:C.muted }}>—</span>}
                    </td>
                    <td style={{ padding:"6px 8px", textAlign:"center", cursor:rc.delayed>0?"pointer":"default" }}
                      onClick={()=>rc.delayed>0&&setRaidModal({title:`Delayed RAIDs — ${comp}`, rows:rc.openItems.filter(r=>String(r[raid.keys.status]||"").toLowerCase()==="delayed")})}>
                      {rc.delayed>0
                        ? <span style={{ fontSize:12, fontWeight:800, background:"#fee2e2", color:"#b91c1c",
                            border:"2px solid #fca5a5", borderRadius:3, padding:"2px 10px",
                            display:"inline-block", minWidth:28, textAlign:"center", lineHeight:"18px" }}>
                            {rc.delayed}
                          </span>
                        : <span style={{ color:C.muted }}>—</span>}
                    </td>
                    <td style={{ padding:"8px 8px", textAlign:"center" }}>
                      {rc.issues.length>0
                        ? <span onClick={()=>setRaidModal({title:`Open Issues — ${comp}`, rows:rc.issues})}
                            style={{ fontSize:12, fontWeight:800, background:"#fee2e2", color:"#b91c1c",
                              border:"2px solid #fca5a5", borderRadius:3, padding:"2px 10px",
                              display:"inline-block", minWidth:28, textAlign:"center",
                              lineHeight:"18px", cursor:"pointer" }}>
                            {rc.issues.length}
                          </span>
                        : <span style={{ color:C.muted }}>—</span>}
                    </td>
                    <td style={{ padding:"8px 8px", textAlign:"center", borderRight:`1px solid ${C.border}` }}>
                      {rc.risks.length>0
                        ? <span onClick={()=>setRaidModal({title:`Open Risks — ${comp}`, rows:rc.risks})}
                            style={{ fontSize:12, fontWeight:800, background:"#fef3c7", color:"#b45309",
                              border:"2px solid #fcd34d", borderRadius:3, padding:"2px 10px",
                              display:"inline-block", minWidth:28, textAlign:"center",
                              lineHeight:"18px", cursor:"pointer" }}>
                            {rc.risks.length}
                          </span>
                        : <span style={{ color:C.muted }}>—</span>}
                    </td>

                    {/* Sprint columns S1-S8 */}
                    {req && sprintOrder.map(entry => {
                      const sd = rq ? getSprintData(rq.sprintData, entry) : null;
                      if (!rq) return <td key={entry.label} style={{ padding:"8px 8px", textAlign:"center", color:C.muted }}>—</td>;
                      if (!sd) return <td key={entry.label} style={{ padding:"8px 8px", textAlign:"center", color:C.muted }}>—</td>;
                      const sprintRows = rq.rows.filter(r => entry.raws.includes(String(r[req.keys.sprint]||"").trim()));
                      const bc = sprintBubbleColor(sd);
                      const tip = `Complete:${sd.complete||0} Partial:${sd.partial||0} InProgress:${sd.inProgress||0} Blocked:${sd.blocked||0} NotStarted:${sd.notStarted||0}`;
                      return (
                        <td key={entry.label} style={{ padding:"5px 6px", textAlign:"center", cursor:"pointer" }}
                          onClick={()=>setStoryModal({title:`${comp} — ${entry.label}`, rows:sprintRows})}>
                          <span title={tip} style={{ fontSize:12, fontWeight:800,
                            background:bc.bg, color:bc.color, border:`2px solid ${bc.border}`,
                            borderRadius:12, padding:"2px 10px",
                            display:"inline-block", minWidth:28, textAlign:"center",
                            lineHeight:"18px" }}>
                            {sd.total}
                          </span>
                        </td>
                      );
                    })}
                    {/* Not Applicable for Technical Build — combined column */}
                    {req && hasNaCol && (() => {
                      const sd = rq ? getNaData(rq.sprintData) : null;
                      if (!rq || !sd) return <td style={{ padding:"8px 8px", textAlign:"center", color:C.muted, borderLeft:`1px solid ${C.border}` }}>—</td>;
                      const naRows = getNaRows(rq.rows);
                      const bc = sprintBubbleColor(sd);
                      const tip = `Complete:${sd.complete||0} Partial:${sd.partial||0} InProgress:${sd.inProgress||0} Blocked:${sd.blocked||0} NotStarted:${sd.notStarted||0}`;
                      return (
                        <td style={{ padding:"5px 6px", textAlign:"center", cursor:"pointer", borderLeft:`1px solid ${C.border}` }}
                          onClick={()=>setStoryModal({title:`${comp} — Not Applicable for Tech Build`, rows:naRows})}>
                          <span title={tip} style={{ fontSize:12, fontWeight:800,
                            background:bc.bg, color:bc.color, border:`2px solid ${bc.border}`,
                            borderRadius:12, padding:"2px 10px",
                            display:"inline-block", minWidth:28, textAlign:"center",
                            lineHeight:"18px" }}>
                            {sd.total}
                          </span>
                        </td>
                      );
                    })()}

                    {/* Total active stories */}
                    {req && (
                      <td style={{ padding:"7px 8px", textAlign:"center",
                        borderLeft:`1px solid ${C.border}`, cursor:rq?"pointer":"default",
                        fontWeight:700, fontSize:12, color:C.navy }}>
                        {rq ? (() => {
                          // Count stories assigned to a sprint (S1-S8 + N/A)
                          const assignedSprints = new Set([
                            ...sprintOrder.flatMap(e => e.raws),
                            ...naSprintRaws,
                          ]);
                          const unassignedRows = rq.rows.filter(r => {
                            const sp = String(r[req.keys.sprint]||"").trim();
                            return !sp || sp === "No Sprint" || !assignedSprints.has(sp);
                          });
                          const unassigned = unassignedRows.length;
                          return (
                            <span title={unassigned > 0 ? `${rq.total} total = sprint-assigned + ${unassigned} with no sprint` : `${rq.total} total`}>
                              <span style={{ cursor:"pointer" }}
                                onClick={() => setStoryModal({ title:`All Stories — ${comp}`, rows:rq.rows })}>
                                {rq.total}
                              </span>
                              {unassigned > 0 && (
                                <span
                                  onClick={e => { e.stopPropagation(); setStoryModal({ title:`No Sprint Assigned — ${comp}`, rows:unassignedRows }); }}
                                  title={`${unassigned} stories with no sprint assigned — click to view`}
                                  style={{ fontSize:9, color:"#b45309", fontWeight:700, marginLeft:4,
                                    cursor:"pointer", background:"#fef3c7", border:"1px solid #fcd34d",
                                    borderRadius:3, padding:"1px 4px" }}>
                                  +{unassigned}?
                                </span>
                              )}
                            </span>
                          );
                        })() : <span style={{ color:C.muted }}>—</span>}
                      </td>
                    )}

                    {/* Functional Build Status — coloured pills, each clickable */}
                    {req && (() => {
                      // Pill colour: Complete=blue, In Progress=green, Blocked=red, Partial=lightblue, N/A=purple, NotStarted=slate
                      const pillStyle = s => {
                        const v = String(s||"").toLowerCase();
                        if (v.includes("complete") && !v.includes("partial"))
                          return { bg:"#dbeafe", color:"#1d4ed8", border:"#93c5fd" };  // blue
                        if (v.includes("progress"))
                          return { bg:"#dcfce7", color:"#15803d", border:"#86efac" };  // green
                        if (v.includes("block"))
                          return { bg:"#fee2e2", color:"#b91c1c", border:"#fca5a5" };  // red
                        if (v.includes("partial"))
                          return { bg:"#e0f2fe", color:"#0369a1", border:"#7dd3fc" };  // light blue
                        if (v.includes("n/a") || v === "na" || v === "not applicable")
                          return { bg:"#f3e8ff", color:"#7e22ce", border:"#d8b4fe" };   // purple — N/A
                        if (v.includes("not start"))
                          return { bg:"#f1f5f9", color:"#334155", border:"#94a3b8" };   // dark grey — Not Started
                        return   { bg:"#f1f5f9", color:"#334155", border:"#94a3b8" };   // dark grey — fallback
                      };
                      return (
                        <td style={{ padding:"7px 8px", textAlign:"left", borderLeft:`1px solid ${C.border}`, verticalAlign:"top" }}>
                          {rq && Object.keys(rq.funcDist).length > 0 ? (
                            <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
                              {Object.entries(rq.funcDist).sort((a,b)=>b[1]-a[1]).map(([status, count]) => {
                                const ps = pillStyle(status);
                                const drillRows = rq.rows.filter(r =>
                                  String(r[req.keys.funcBuildStatus]||"").trim() === status);
                                return (
                                  <span key={status} title={status}
                                    onClick={()=>setStoryModal({title:`${comp} — Func: ${status}`, rows:drillRows})}
                                    style={{ background:ps.bg, color:ps.color, border:`1px solid ${ps.border}`,
                                      borderRadius:10, padding:"2px 9px", fontSize:11, fontWeight:500,
                                      cursor:"pointer", whiteSpace:"nowrap", display:"inline-flex",
                                      alignItems:"center", gap:4 }}>
                                    {count} <span style={{ fontSize:10, opacity:.7 }}>{status.slice(0,14)}</span>
                                  </span>
                                );
                              })}
                            </div>
                          ) : <span style={{ color:C.muted }}>—</span>}
                        </td>
                      );
                    })()}

                    {/* Tech Build Status — coloured pills, each clickable */}
                    {req && (() => {
                      const pillStyle = s => {
                        const v = String(s||"").toLowerCase();
                        if (v.includes("complete") && !v.includes("partial"))
                          return { bg:"#dbeafe", color:"#1d4ed8", border:"#93c5fd" };
                        if (v.includes("progress"))
                          return { bg:"#dcfce7", color:"#15803d", border:"#86efac" };
                        if (v.includes("block"))
                          return { bg:"#fee2e2", color:"#b91c1c", border:"#fca5a5" };
                        if (v.includes("partial"))
                          return { bg:"#e0f2fe", color:"#0369a1", border:"#7dd3fc" };
                        if (v.includes("n/a") || v === "na" || v === "not applicable")
                          return { bg:"#f3e8ff", color:"#7e22ce", border:"#d8b4fe" };   // purple — N/A
                        if (v.includes("not start"))
                          return { bg:"#f1f5f9", color:"#334155", border:"#94a3b8" };   // dark grey — Not Started
                        return   { bg:"#f1f5f9", color:"#334155", border:"#94a3b8" };   // dark grey — fallback
                      };
                      return (
                        <td style={{ padding:"7px 8px", textAlign:"left", verticalAlign:"top" }}>
                          {rq && Object.keys(rq.techDist).length > 0 ? (
                            <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
                              {Object.entries(rq.techDist).sort((a,b)=>b[1]-a[1]).map(([status, count]) => {
                                const ps = pillStyle(status);
                                const drillRows = rq.rows.filter(r =>
                                  String(r[req.keys.techBuildStatus]||"").trim() === status);
                                return (
                                  <span key={status} title={status}
                                    onClick={()=>setStoryModal({title:`${comp} — Tech: ${status}`, rows:drillRows})}
                                    style={{ background:ps.bg, color:ps.color, border:`1px solid ${ps.border}`,
                                      borderRadius:10, padding:"2px 9px", fontSize:11, fontWeight:500,
                                      cursor:"pointer", whiteSpace:"nowrap", display:"inline-flex",
                                      alignItems:"center", gap:4 }}>
                                    {count} <span style={{ fontSize:10, opacity:.7 }}>{status.slice(0,14)}</span>
                                  </span>
                                );
                              })}
                            </div>
                          ) : <span style={{ color:C.muted }}>—</span>}
                        </td>
                      );
                    })()}

                    {/* Workplan Design Status + Build Status from Lvl 3 children */}
                    {wp && (() => {
                      const cw = getCompWp(comp);
                      return (
                        <>
                          <td style={{ padding:"7px 8px", textAlign:"center", borderLeft:`1px solid ${C.border}`, verticalAlign:"middle" }}>
                            {cw ? wpStatusPill(
                              cw.designStatus,
                              cw.designRows,
                              `Design tasks: ${cw.designRows.length}`,
                              cw.designRows.length ? () => setWpDrillModal({ title:`${comp} — Design`, rows:cw.designRows }) : null
                            ) : <span style={{ color:C.muted }}>—</span>}
                          </td>
                          <td style={{ padding:"7px 8px", textAlign:"center", verticalAlign:"middle" }}>
                            {cw ? wpStatusPill(
                              cw.buildStatus,
                              cw.buildRows,
                              `Build tasks: ${cw.buildRows.length}`,
                              cw.buildRows.length ? () => setWpDrillModal({ title:`${comp} — Build`, rows:cw.buildRows }) : null
                            ) : <span style={{ color:C.muted }}>—</span>}
                          </td>
                          <td style={{ padding:"7px 8px", textAlign:"center", verticalAlign:"middle" }}>
                            {cw && cw.pctComplete != null ? (
                              <div style={{ display:"flex", alignItems:"center", gap:6, justifyContent:"center" }}>
                                <div style={{ width:60, background:"#e2e8f0", borderRadius:4, height:7, overflow:"hidden" }}>
                                  <div style={{ width:`${cw.pctComplete}%`, height:"100%", borderRadius:4,
                                    background: (() => {
                                      const s = String(cw.buildStatus || cw.designStatus || "").toLowerCase();
                                      if (s.includes("off track"))  return C.delayed;
                                      if (s.includes("on track"))   return C.gold;
                                      if (s.includes("complete"))   return C.green;
                                      return "#94a3b8"; // not started / unknown → grey
                                    })() }} />
                                </div>
                                <span style={{ color:C.text, fontWeight:700, fontSize:11 }}>{cw.pctComplete}%</span>
                              </div>
                            ) : <span style={{ color:C.muted }}>—</span>}
                          </td>
                        </>
                      );
                    })()}

                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      {/* RAID drill-down — same rich modal as RAID Analysis tab */}
      {raidModal && raid && (() => {
        const K = raid.keys;
        const teamKey = K.team || "Primary Team (Owner)";
        const statusCol = s => { const sl = String(s||"").toLowerCase(); return sl.includes("delay") ? C.delayed : sl.includes("complete") ? C.complete : C.onTrack; };
        const allModalTeams = Array.from(new Set(raidModal.rows.map(r => String(r[teamKey]||"").trim()).filter(Boolean))).sort();
        const allModalTypes = Array.from(new Set(raidModal.rows.map(r => String(r[K.type]||"").trim()).filter(Boolean))).sort();
        const allModalComps = Array.from(new Set(raidModal.rows.map(r => String(r[K.component]||"").trim()).filter(Boolean))).sort();
        return (
          <RaidKpiModal
            title={raidModal.title}
            rows={raidModal.rows}
            K={K} teamKey={teamKey}
            allTeams={allModalTeams} allTypes={allModalTypes} allComps={allModalComps}
            statusCol={statusCol}
            hideType={raidModal.hideType || false}
            hideStatus={raidModal.hideStatus || false}
            colConfig={modalColConfig}
            setColConfig={setModalColConfig}
            onClose={() => setRaidModal(null)}
          />
        );
      })()}
      {/* Story drill-down */}
      {storyModal && <StoryDrillModal title={storyModal.title} rows={storyModal.rows} reqKeys={req?.keys} onClose={()=>setStoryModal(null)} />}
      {/* Workplan hierarchy drill-down */}
      {wpDrillModal && <WorkplanDrillModal title={wpDrillModal.title} rows={wpDrillModal.rows} onClose={()=>setWpDrillModal(null)} />}
    </div>
  );
}

// ─── WORKPLAN TAB ────────────────────────────────────────────────────────────
function WorkplanTab({ wp, raid, openModal }) {
  const [subTab, setSubTab] = useState("workstream");
  const [wpModal, setWpModal] = useState(null);
  const [wsFilter, setWsFilter] = useState("All");
  const [sapFilter, setSapFilter] = useState("All");
  const [epFilter, setEpFilter] = useState("All");

  if (!wp) return <Empty label="Upload Workplan file above to view this tab." />;

  // ── Shared data calculations ─────────────────────────────────────────────

  const wsMap = {};
  wp.allRows.forEach(r => {
    const ws = String(r["Activity Grp - Lvl 1"] || r["Workstream"] || "").trim();
    // Robust leaf detection — Smartsheet API may return Children as string, number, null or empty
    const ch = r["Children"];
    const isLeaf = ch === "" || ch === null || ch === undefined || String(ch) === "0" || Number(ch) === 0;
    if (!ws || !isLeaf) return;
    if (!wsMap[ws]) wsMap[ws] = { total: 0, offTrack: 0, onTrack: 0, complete: 0, notStarted: 0, rows: [] };
    wsMap[ws].total++;
    wsMap[ws].rows.push(r);
    const s = String(r["Default Status"] || r["Status"] || "").toLowerCase();
    if (s.includes("off track")) wsMap[ws].offTrack++;
    else if (s.includes("on track")) wsMap[ws].onTrack++;
    else if (s.includes("complete")) wsMap[ws].complete++;
    else wsMap[ws].notStarted++;
  });
  const workstreams = Object.entries(wsMap).map(([name, d]) => {
    const pctVal = d.total > 0 ? Math.round((d.complete / d.total) * 100) : 0;
    const health = d.offTrack > 0 ? "Off Track" : d.complete === d.total && d.total > 0 ? "Complete" : d.onTrack > 0 ? "On Track" : "Not Started";
    return { name, d, pctVal, health };
  }).sort((a, b) => b.d.offTrack - a.d.offTrack);

  // Component RAG (same logic as Executive Summary)
  const compRows = wp.allRows.filter(r =>
    String(r["Activity Grp - Lvl 1"] || "").trim() === "Technology - SAP Configuration & Build" &&
    String(r["Activity Grp - Lvl 2"] || "").trim() === "Component Build"
  );
  const compNames = Array.from(new Set(compRows.map(r => String(r["Activity Grp - Lvl 3"] || "").trim()).filter(Boolean))).sort();

  const getCompStatus = (compName) => {
    const rows = compRows.filter(r => String(r["Activity Grp - Lvl 3"] || "").trim() === compName);
    const leaves = rows.filter(r => { const c = r["Children"]; return !c || Number(c) === 0; });
    const getS = r => String(r["Default Status"] || r["Status"] || "").toLowerCase();
    const normPct = v => {
      if (v === "" || v == null) return null;
      const s = String(v).replace("%", "").trim();
      if (s === "" || isNaN(Number(s))) return null;
      const n = Number(s);
      return n <= 1 ? Math.round(n * 100) : Math.round(n);
    };
    // A leaf is "done" if status includes complete OR pct = 100
    const isDone = r => getS(r).includes("complete") || normPct(r["% Complete"] ?? r["% complete"]) === 100;
    const hasOffTrack = leaves.some(r => getS(r).includes("off track"));
    const hasOnTrack  = leaves.some(r => getS(r).includes("on track"));
    const allComplete = leaves.length > 0 && leaves.every(r => isDone(r));
    const lvl3Header = rows.find(r => Number(r["Lvl"] ?? 0) === 3);
    const headerPct = lvl3Header ? normPct(lvl3Header["% Complete"] ?? lvl3Header["% complete"]) : null;
    // If header row shows 100%, treat as Complete regardless of leaf statuses
    const status = hasOffTrack ? "Off Track" : (allComplete || headerPct === 100) ? "Complete" : hasOnTrack ? "On Track" : "Not Started";
    const delayedCount = leaves.filter(r => getS(r).includes("off track")).length;
    const pctValues = leaves.map(r => normPct(r["% Complete"] ?? r["% complete"])).filter(v => v != null);
    const p = headerPct != null ? headerPct : (pctValues.length ? Math.round(pctValues.reduce((a,b) => a+b,0) / pctValues.length) : null);
    return { status, pct: p, rows, delayedCount };
  };

  // ── E&P children ─────────────────────────────────────────────────────────
  // Handle both column name casings — "Activity Grp - Lvl 0" and "Activity Grp - LVL 0"
  const getEpLvl0 = r => String(r["Activity Grp - Lvl 0"] || r["Activity Grp - LVL 0"] || "").trim();
  const getEpLvl1 = r => String(r["Activity Grp - Lvl 1"] || r["Activity Grp - LVL 1"] || "").trim();
  const getEpLvl2 = r => String(r["Activity Grp - Lvl 2"] || r["Activity Grp - LVL 2"] || "").trim();

  const epAllRows = wp.allRows.filter(r => getEpLvl0(r) === "E&P");

  // Unique Lvl 1 names that have children (group headers, not leaf tasks)
  const epChildNames = Array.from(new Set(
    epAllRows
      .filter(r => {
        const ch = r["Children"];
        const hasKids = ch && String(ch) !== "0" && Number(ch) !== 0;
        return getEpLvl1(r) && hasKids;
      })
      .map(r => getEpLvl1(r))
  )).filter(Boolean).sort();

  const getEpChildStatus = (childName) => {
    const subtree = epAllRows.filter(r => getEpLvl1(r) === childName);
    const isLeaf = r => { const c = r["Children"]; return c === "" || c === null || c === undefined || String(c) === "0" || Number(c) === 0; };
    const leaves = subtree.filter(isLeaf);
    const getS = r => String(r["Default Status"] || r["Status"] || "").toLowerCase();

    // Same logic as Workstream Status and SAP Config tabs:
    // Red = any leaf off track, Blue = all complete, Amber = any on track, Grey = not started
    const normPct = v => { if (v === "" || v == null) return null; const s = String(v).replace("%","").trim(); if (!s || isNaN(Number(s))) return null; const n = Number(s); return n <= 1 ? Math.round(n*100) : Math.round(n); };
    const isDone = r => getS(r).includes("complete") || normPct(r["% Complete"] ?? r["% complete"]) === 100;
    const hasOffTrack = leaves.some(r => getS(r).includes("off track"));
    const hasOnTrack  = leaves.some(r => getS(r).includes("on track"));
    const allComplete = leaves.length > 0 && leaves.every(r => isDone(r));
    const delayedCount = leaves.filter(r => getS(r).includes("off track")).length;

    // % Complete: use header row rollup first, fall back to leaf average
    const headerRow = subtree.find(r => { const ch = r["Children"]; return ch && String(ch) !== "0" && Number(ch) !== 0 && !getEpLvl2(r); }) || subtree[0];
    const headerPct = headerRow ? normPct(headerRow["% Complete"] ?? headerRow["% complete"]) : null;
    const status = hasOffTrack ? "Off Track" : (allComplete || headerPct === 100) ? "Complete" : hasOnTrack ? "On Track" : "Not Started";
    const leafPcts = leaves.map(r => normPct(r["% Complete"] ?? r["% complete"])).filter(v => v != null);
    const p = headerPct != null ? headerPct : (leafPcts.length ? Math.round(leafPcts.reduce((a,b)=>a+b,0)/leafPcts.length) : null);

    return { status, pct: p, rows: subtree, delayedCount };
  };

  const SUB_TABS = [
    { id: "workstream", label: "Workstream Status" },
    { id: "sapbuild",   label: "SAP Config & Build" },
    { id: "ep",         label: "E&P" },
  ];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>

      {/* Sub-tab bar */}
      <div style={{ background: "#43978F", borderBottom: `1px solid #357a73`, marginBottom: 16 }}>
        <div style={{ display: "flex", paddingLeft: 4 }}>
          {SUB_TABS.map(st => (
            <button key={st.id} onClick={() => setSubTab(st.id)} style={{
              padding: "10px 20px", border: "none", background: "transparent", cursor: "pointer",
              color: subTab === st.id ? "#fff" : "rgba(255,255,255,0.7)",
              borderBottom: `3px solid ${subTab === st.id ? "#fff" : "transparent"}`,
              fontWeight: subTab === st.id ? 700 : 500, fontSize: 13, transition: "all .12s",
            }}>{st.label}</button>
          ))}
        </div>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>

        {/* ── Workstream Status sub-tab ──────────────────────────────────────── */}
        {subTab === "workstream" && (
          <>
            {workstreams.length === 0 ? (
              <Empty label="No workstream data found in workplan." />
            ) : (
              <Card>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
                  <SecTitle title="Workstream Status by Workplan" color={C.navyLight} />
                  {/* Filter pills */}
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {[
                      { label: "All",         bg: C.navyLight, count: workstreams.length },
                      { label: "Delayed",     bg: "#b91c1c",   count: workstreams.filter(w => w.d.offTrack > 0).length },
                      { label: "On Track",    bg: "#92400e",   count: workstreams.filter(w => w.d.onTrack > 0 && w.d.offTrack === 0).length },
                      { label: "Not Started", bg: "#475569",   count: workstreams.filter(w => w.d.notStarted === w.d.total && w.d.total > 0).length },
                      { label: "Complete",    bg: "#1d4ed8",   count: workstreams.filter(w => w.d.complete === w.d.total && w.d.total > 0).length },
                    ].map(({ label, bg, count }) => {
                      const active = wsFilter === label;
                      return (
                        <button key={label} onClick={() => setWsFilter(label)}
                          style={{ display: "flex", alignItems: "center", gap: 5, padding: "4px 11px",
                            borderRadius: 20, border: `2px solid ${active ? bg : C.border}`,
                            background: active ? bg : C.white,
                            color: active ? "#fff" : C.muted,
                            cursor: "pointer", fontSize: 11, fontWeight: 700, transition: "all .15s" }}>
                          {label}
                          <span style={{ background: active ? "rgba(255,255,255,0.3)" : "#f1f5f9",
                            color: active ? "#fff" : C.text,
                            borderRadius: 10, padding: "1px 6px", fontSize: 10, fontWeight: 800 }}>
                            {count}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(155px, 1fr))", gap: 8 }}>
                  {workstreams
                    .filter(({ d }) => {
                      if (wsFilter === "All")         return true;
                      if (wsFilter === "Delayed")     return d.offTrack > 0;
                      if (wsFilter === "On Track")    return d.onTrack > 0 && d.offTrack === 0;
                      if (wsFilter === "Not Started") return d.notStarted === d.total && d.total > 0;
                      if (wsFilter === "Complete")    return d.complete === d.total && d.total > 0;
                      return true;
                    })
                    .map(({ name, d, pctVal }) => {
                    const hasDelayed  = d.offTrack > 0;
                    const allComplete = d.complete === d.total && d.total > 0;
                    const hasOnTrack  = d.onTrack > 0 && !hasDelayed;
                    const cellBg  = hasDelayed ? "#fee2e2" : allComplete ? "#dbeafe" : hasOnTrack ? "#fef3c7" : "#f1f5f9";
                    const textCol = hasDelayed ? "#b91c1c" : allComplete ? "#1d4ed8" : hasOnTrack ? "#92400e" : "#475569";
                    const borderC = hasDelayed ? "#fca5a5" : allComplete ? "#93c5fd" : hasOnTrack ? "#fcd34d" : "#e2e8f0";
                    const shortName = name.replace("Technology - SAP Configuration & Build", "SAP Config & Build").replace("Technology - ", "");
                    return (
                      <div key={name}
                        onClick={() => setWpModal({ title: name, rows: wp.allRows.filter(r => String(r["Activity Grp - Lvl 1"] || r["Workstream"] || "").trim() === name), initialFilter: d.offTrack > 0 ? "Off Track" : "All" })}
                        onMouseEnter={e => { e.currentTarget.style.boxShadow = "0 3px 10px rgba(0,0,0,0.12)"; e.currentTarget.style.transform = "translateY(-1px)"; }}
                        onMouseLeave={e => { e.currentTarget.style.boxShadow = "none"; e.currentTarget.style.transform = "none"; }}
                        style={{ background: cellBg, border: `1.5px solid ${borderC}`, borderRadius: 8, padding: "10px 12px", cursor: "pointer", transition: "box-shadow .15s, transform .15s", position: "relative" }}>
                        {hasDelayed && (
                          <div style={{ position: "absolute", top: 7, right: 7, background: C.delayed, color: "#fff", borderRadius: 8, padding: "1px 6px", fontSize: 9, fontWeight: 700 }}>⚠ {d.offTrack}</div>
                        )}
                        <div style={{ fontSize: 11, fontWeight: 700, color: textCol, marginBottom: 6, paddingRight: hasDelayed ? 36 : 0, lineHeight: 1.3 }}>{shortName}</div>
                        <div style={{ fontSize: 22, fontWeight: 800, color: textCol, lineHeight: 1, marginBottom: 6 }}>{pctVal ?? 0}%</div>
                        <div style={{ background: "rgba(0,0,0,0.08)", borderRadius: 3, height: 4, overflow: "hidden", marginBottom: 7 }}>
                          <div style={{ width: `${pctVal ?? 0}%`, height: "100%", background: textCol, borderRadius: 3, opacity: 0.65 }} />
                        </div>
                        <div style={{ fontSize: 10, color: textCol, opacity: 0.75 }}>{d.total} tasks · {d.complete} done</div>
                      </div>
                    );
                  })}
                </div>
                <div style={{ color: C.muted, fontSize: 10, marginTop: 8 }}>
                  Click a filter to narrow cards · Click any card to drill into tasks · ⚠ = off-track count
                </div>
              </Card>
            )}
          </>
        )}

        {/* ── SAP Config & Build sub-tab ─────────────────────────────────────── */}
        {subTab === "sapbuild" && (
          <>
            {compNames.length === 0 ? (
              <Empty label="No SAP Config & Build components found in workplan." />
            ) : (
              <Card>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
                  <SecTitle title="Component RAG Status — SAP Config & Build" color={C.navyLight} />
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {[
                      { label: "All",         bg: C.navyLight, count: compNames.length },
                      { label: "Off Track",   bg: "#b91c1c",   count: compNames.filter(n => getCompStatus(n).status === "Off Track").length },
                      { label: "On Track",    bg: "#92400e",   count: compNames.filter(n => getCompStatus(n).status === "On Track").length },
                      { label: "Not Started", bg: "#475569",   count: compNames.filter(n => getCompStatus(n).status === "Not Started").length },
                      { label: "Complete",    bg: "#1d4ed8",   count: compNames.filter(n => getCompStatus(n).status === "Complete").length },
                    ].map(({ label, bg, count }) => {
                      const active = sapFilter === label;
                      return (
                        <button key={label} onClick={() => setSapFilter(label)}
                          style={{ display: "flex", alignItems: "center", gap: 5, padding: "4px 11px",
                            borderRadius: 20, border: `2px solid ${active ? bg : C.border}`,
                            background: active ? bg : C.white, color: active ? "#fff" : C.muted,
                            cursor: "pointer", fontSize: 11, fontWeight: 700, transition: "all .15s" }}>
                          {label}
                          <span style={{ background: active ? "rgba(255,255,255,0.3)" : "#f1f5f9",
                            color: active ? "#fff" : C.text,
                            borderRadius: 10, padding: "1px 6px", fontSize: 10, fontWeight: 800 }}>
                            {count}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(140px, 1fr))", gap: 7 }}>
                  {compNames
                    .filter(name => {
                      if (sapFilter === "All") return true;
                      return getCompStatus(name).status === sapFilter;
                    })
                    .map(name => {
                    const { status, pct: p, rows, delayedCount } = getCompStatus(name);
                    const sl = status.toLowerCase();
                    const isOffTrack = sl.includes("off track");
                    const isComplete = sl.includes("complete");
                    const isOnTrack  = sl.includes("on track");
                    const cellBg  = isOffTrack ? "#fee2e2" : isComplete ? "#dbeafe" : isOnTrack ? "#fef3c7" : "#f1f5f9";
                    const textCol = isOffTrack ? "#b91c1c" : isComplete ? "#1d4ed8" : isOnTrack ? "#92400e" : "#475569";
                    const borderC = isOffTrack ? "#fca5a5" : isComplete ? "#93c5fd" : isOnTrack ? "#fcd34d" : "#e2e8f0";
                    const pctVal  = p ?? 0;
                    return (
                      <div key={name}
                        onClick={() => setWpModal({ title: name, rows })}
                        onMouseEnter={e => { e.currentTarget.style.boxShadow = "0 3px 10px rgba(0,0,0,0.12)"; e.currentTarget.style.transform = "translateY(-1px)"; }}
                        onMouseLeave={e => { e.currentTarget.style.boxShadow = "none"; e.currentTarget.style.transform = "none"; }}
                        style={{ background: cellBg, border: `1.5px solid ${borderC}`, borderRadius: 8, padding: "9px 11px", cursor: "pointer", transition: "box-shadow .15s, transform .15s", position: "relative" }}>
                        {isOffTrack && (
                          <div style={{ position: "absolute", top: 6, right: 6, background: C.delayed, color: "#fff", borderRadius: 7, padding: "1px 6px", fontSize: 9, fontWeight: 700 }}>⚠ {delayedCount}</div>
                        )}
                        <div style={{ fontSize: 10, fontWeight: 700, color: textCol, marginBottom: 5, paddingRight: isOffTrack ? 22 : 0, lineHeight: 1.3 }}>{name}</div>
                        <div style={{ fontSize: 20, fontWeight: 800, color: textCol, lineHeight: 1, marginBottom: 5 }}>{pctVal}%</div>
                        <div style={{ background: "rgba(0,0,0,0.08)", borderRadius: 3, height: 4, overflow: "hidden", marginBottom: 6 }}>
                          <div style={{ width: `${pctVal}%`, height: "100%", background: textCol, borderRadius: 3, opacity: 0.65 }} />
                        </div>
                        <span style={{ background: isOffTrack ? "#fee2e220" : isComplete ? "#dbeafe20" : isOnTrack ? "#fef3c720" : "#f1f5f920",
                          color: textCol, border: `1px solid ${textCol}40`, borderRadius: 4,
                          padding: "2px 7px", fontSize: 10, fontWeight: 700 }}>{status}</span>
                      </div>
                    );
                  })}
                </div>
                <div style={{ color: C.muted, fontSize: 10, marginTop: 8 }}>
                  Click a filter to narrow components · Click any card to drill into workplan tasks · ⚠ = off-track count
                </div>
              </Card>
            )}
          </>
        )}

        {/* ── E&P sub-tab ──────────────────────────────────────────────────── */}
        {subTab === "ep" && (
          <>
            {epChildNames.length === 0 ? (
              <Empty label="No E&P children found in workplan. Ensure E&P row exists at Lvl 2." />
            ) : (
              <Card>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
                  <SecTitle title="E&P — Component Status" color={C.navyLight} />
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {[
                      { label: "All",         bg: C.navyLight, count: epChildNames.length },
                      { label: "Off Track",   bg: "#b91c1c",   count: epChildNames.filter(n => getEpChildStatus(n).status.toLowerCase().includes("off track")).length },
                      { label: "On Track",    bg: "#92400e",   count: epChildNames.filter(n => { const s=getEpChildStatus(n).status.toLowerCase(); return s.includes("on track") && !s.includes("off"); }).length },
                      { label: "Not Started", bg: "#475569",   count: epChildNames.filter(n => getEpChildStatus(n).status.toLowerCase().includes("not start")).length },
                      { label: "Complete",    bg: "#1d4ed8",   count: epChildNames.filter(n => getEpChildStatus(n).status.toLowerCase().includes("complete")).length },
                    ].map(({ label, bg, count }) => {
                      const active = epFilter === label;
                      return (
                        <button key={label} onClick={() => setEpFilter(label)}
                          style={{ display: "flex", alignItems: "center", gap: 5, padding: "4px 11px",
                            borderRadius: 20, border: `2px solid ${active ? bg : C.border}`,
                            background: active ? bg : C.white, color: active ? "#fff" : C.muted,
                            cursor: "pointer", fontSize: 11, fontWeight: 700, transition: "all .15s" }}>
                          {label}
                          <span style={{ background: active ? "rgba(255,255,255,0.3)" : "#f1f5f9",
                            color: active ? "#fff" : C.text,
                            borderRadius: 10, padding: "1px 6px", fontSize: 10, fontWeight: 800 }}>
                            {count}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(155px, 1fr))", gap: 8 }}>
                  {epChildNames
                    .filter(name => {
                      if (epFilter === "All") return true;
                      const s = getEpChildStatus(name).status.toLowerCase();
                      if (epFilter === "Off Track")   return s.includes("off track");
                      if (epFilter === "On Track")    return s.includes("on track") && !s.includes("off");
                      if (epFilter === "Not Started") return s.includes("not start");
                      if (epFilter === "Complete")    return s.includes("complete");
                      return true;
                    })
                    .map(name => {
                      const { status, pct: p, rows, delayedCount } = getEpChildStatus(name);
                      const sl = status.toLowerCase();
                      const isOffTrack = sl.includes("off track");
                      const isComplete = sl.includes("complete");
                      const isOnTrack  = sl.includes("on track") && !isOffTrack;
                      // Same colours as Workstream Status and SAP Config tabs
                      const cellBg  = isOffTrack ? "#fee2e2" : isComplete ? "#dbeafe" : isOnTrack ? "#fef3c7" : "#f1f5f9";
                      const textCol = isOffTrack ? "#b91c1c" : isComplete ? "#1d4ed8" : isOnTrack ? "#92400e" : "#475569";
                      const borderC = isOffTrack ? "#fca5a5" : isComplete ? "#93c5fd" : isOnTrack ? "#fcd34d" : "#e2e8f0";
                      const pctVal  = p ?? 0;
                      return (
                        <div key={name}
                          onClick={() => setWpModal({ title: name, rows, initialFilter: isOffTrack ? "Off Track" : "All" })}
                          onMouseEnter={e => { e.currentTarget.style.boxShadow = "0 3px 10px rgba(0,0,0,0.12)"; e.currentTarget.style.transform = "translateY(-1px)"; }}
                          onMouseLeave={e => { e.currentTarget.style.boxShadow = "none"; e.currentTarget.style.transform = "none"; }}
                          style={{ background: cellBg, border: `1.5px solid ${borderC}`, borderRadius: 8, padding: "10px 12px", cursor: "pointer", transition: "box-shadow .15s, transform .15s", position: "relative" }}>

                          {/* Delayed badge — top right, same as other tabs */}
                          {delayedCount > 0 && (
                            <div style={{ position: "absolute", top: 7, right: 7, background: "#b91c1c", color: "#fff", borderRadius: 8, padding: "1px 6px", fontSize: 9, fontWeight: 700 }}>
                              ⚠ {delayedCount}
                            </div>
                          )}
                          <div style={{ fontSize: 12, fontWeight: 700, color: textCol, marginBottom: 6, paddingRight: delayedCount > 0 ? 36 : 0, lineHeight: 1.3 }}>{name}</div>
                          <div style={{ fontSize: 22, fontWeight: 800, color: textCol, lineHeight: 1, marginBottom: 6 }}>{pctVal}%</div>
                          <div style={{ background: "rgba(0,0,0,0.08)", borderRadius: 3, height: 4, overflow: "hidden", marginBottom: 7 }}>
                            <div style={{ width: `${pctVal}%`, height: "100%", background: textCol, borderRadius: 3, opacity: 0.65 }} />
                          </div>
                          <div style={{ fontSize: 10, color: textCol, opacity: 0.75 }}>
                            {rows.filter(r => { const c = r["Children"]; return c === "" || c === null || c === undefined || String(c) === "0" || Number(c) === 0; }).length} tasks
                          </div>
                        </div>
                      );
                    })}
                </div>
                <div style={{ color: C.muted, fontSize: 10, marginTop: 8 }}>
                  Click a filter to narrow · Click any card to drill into tasks · ⚠ = off-track count
                </div>
              </Card>
            )}
          </>
        )}

      </div>

      {wpModal && <WorkplanDrillModal title={wpModal.title} rows={wpModal.rows} initialFilter={wpModal.initialFilter} onClose={() => setWpModal(null)} />}
    </div>
  );
}


// ─── RAID TAB ────────────────────────────────────────────────────────────────
function RaidTab({ data, openModal }) {
  if (!data) return <Empty label="Upload the RAID Log to view this section." />;
  const { byPriority, byComponent, byTeam, keys: K } = data;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", gap: 11 }}>
        <KpiCard label="Open RAIDs" value={data.total} color={C.navyLight} onClick={() => openModal("All RAID Items", data.items)} />
        <KpiCard label="Open Issues" value={data.openIssues.length} color={C.delayed} onClick={() => openModal("Open Issues", data.openIssues)} />
        <KpiCard label="Open Risks"  value={data.openRisks.length}  color={C.gold}    onClick={() => openModal("Open Risks",  data.openRisks)} />
        <KpiCard label="On Track RAIDs" value={data.total - data.delayed.length} color={C.onTrack} onClick={() => openModal("On Track RAIDs", data.items.filter(r => !String(r[K.status] || "").toLowerCase().includes("delay")))} />
        <KpiCard label="Delayed RAIDs" value={data.delayed.length} color={C.delayed} onClick={() => openModal("Delayed RAIDs", data.delayed)} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
        <Card>
          <SecTitle title="Open RAID by Priority & Status" color={C.delayed} />
          <HSBar data={Object.entries(byPriority).map(([name, d]) => ({ name, onTrack: d.onTrack, delayed: d.delayed, rows: d.rows }))} valueKeys={["onTrack", "delayed"]} colors={[C.onTrack, C.delayed]} onBarClick={row => openModal(`Priority: ${row.name}`, row.rows)} />
          <Leg items={[{ label: "On Track", color: C.onTrack }, { label: "Delayed", color: C.delayed }]} />
        </Card>
        <Card>
          <SecTitle title="Open RAID by Team" color={C.navyLight} />
          <HSBar data={Object.entries(byTeam).sort((a, b) => (b[1].onTrack + b[1].delayed) - (a[1].onTrack + a[1].delayed)).slice(0, 10).map(([name, d]) => ({ name, onTrack: d.onTrack, delayed: d.delayed, rows: d.rows }))} valueKeys={["onTrack", "delayed"]} colors={[C.onTrack, C.delayed]} onBarClick={row => openModal(`Team: ${row.name}`, row.rows)} />
          <Leg items={[{ label: "On Track", color: C.onTrack }, { label: "Delayed", color: C.delayed }]} />
        </Card>
        <Card>
          <SecTitle title="Open RAID by Component" color={C.complete} />
          <HSBar data={Object.entries(byComponent).sort((a, b) => (b[1].onTrack + b[1].delayed) - (a[1].onTrack + a[1].delayed)).slice(0, 10).map(([name, d]) => ({ name, onTrack: d.onTrack, delayed: d.delayed, rows: d.rows }))} valueKeys={["onTrack", "delayed"]} colors={[C.onTrack, C.delayed]} onBarClick={row => openModal(`Component: ${row.name}`, row.rows)} />
          <Leg items={[{ label: "On Track", color: C.onTrack }, { label: "Delayed", color: C.delayed }]} />
        </Card>
      </div>
      <Card>
        <SecTitle title="RAID Item Detail" color={C.navyLight} />
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead><tr style={{ background: "#f0f4f8" }}>{[K.type, K.desc, K.status, K.priority, K.owner].filter(Boolean).map(c => <th key={c} style={{ textAlign: "left", padding: "8px 10px", color: C.muted, fontSize: 11, fontWeight: 700, borderBottom: `2px solid ${C.border}` }}>{c}</th>)}</tr></thead>
            <tbody>
              {data.items.slice(0, 50).map((r, i) => {
                const s = String(r[K.status] || "—"); const sc = SC[s] || C.muted;
                const p = String(r[K.priority] || "—"); const pc = SC[p] || C.muted;
                return (
                  <tr key={i} style={{ borderBottom: `1px solid ${C.border}`, background: i % 2 === 0 ? C.white : "#f9fafb", cursor: "pointer" }} onClick={() => openModal("Detail", [r])}>
                    {[K.type, K.desc, K.status, K.priority, K.owner].filter(Boolean).map(k => {
                      const v = String(r[k] || "—");
                      return <td key={k} style={{ padding: "7px 10px", maxWidth: 240 }} title={v}>
                        {k === K.status ? <span style={{ background: sc + "20", color: sc, border: `1px solid ${sc}40`, borderRadius: 4, padding: "2px 7px", fontSize: 10, fontWeight: 700 }}>{v}</span>
                          : k === K.priority ? <span style={{ background: pc + "20", color: pc, border: `1px solid ${pc}40`, borderRadius: 4, padding: "2px 7px", fontSize: 10, fontWeight: 700 }}>{v}</span>
                            : v.slice(0, 60)}
                      </td>;
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ─── BUILD MANAGEMENT TAB ────────────────────────────────────────────────────
function BuildTab({ data, openModal }) {
  if (!data) return <Empty label="Upload Requirements / User Stories to view Build Management." />;
  const { bySprint, byComponent, keys: K } = data;
  const sprintData = Object.entries(bySprint).sort((a, b) => String(a[0]).localeCompare(String(b[0]))).map(([name, d]) => ({ name, ...d }));
  const compData = Object.entries(byComponent).sort((a, b) => b[1].total - a[1].total).slice(0, 14).map(([name, d]) => ({ name, ...d }));
  const latest = sprintData[sprintData.length - 1];
  const latestTotal = latest ? latest.complete + latest.inProgress + latest.notStarted + latest.blocked : 0;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        <Card>
          <div style={{ background: C.headerBg, margin: "-16px -16px 14px", padding: "9px 16px", borderRadius: "8px 8px 0 0" }}>
            <span style={{ color: "#fff", fontWeight: 700, fontSize: 12 }}>Current Sprint {latest ? `— ${latest.name}` : ""}</span>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", gap: 9 }}>
            {[["Total", latestTotal, C.text, null], ["Complete", latest?.complete || 0, C.complete, "complete"], ["In Progress", latest?.inProgress || 0, C.inProgress, "inProgress"], ["Not Started", latest?.notStarted || 0, C.muted, "notStarted"], ["Blocked", latest?.blocked || 0, C.delayed, "blocked"]].map(([l, v, co, bk]) => (
              <KpiCard key={l} label={l} value={v} color={co} onClick={bk && latest?.rows?.length ? () => openModal(`${l} — ${latest.name}`, latest.rows.filter(r => { const s = String(r[K.status] || "").toLowerCase(); return bk === "complete" ? s.includes("done") || s.includes("complete") : bk === "inProgress" ? s.includes("progress") : bk === "blocked" ? s.includes("block") : !s.includes("done") && !s.includes("complete") && !s.includes("progress") && !s.includes("block"); })) : null} />
            ))}
          </div>
        </Card>
        <Card>
          <div style={{ background: C.headerBg, margin: "-16px -16px 14px", padding: "9px 16px", borderRadius: "8px 8px 0 0" }}>
            <span style={{ color: "#fff", fontWeight: 700, fontSize: 12 }}>Overall</span>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", gap: 9 }}>
            {[["Total", data.total, C.text, data.items], ["Complete", data.done.length, C.complete, data.done], ["In Progress", data.inProg.length, C.inProgress, data.inProg], ["Not Started", data.notStarted.length, C.muted, data.notStarted], ["Blockers", data.blocked.length, C.delayed, data.blocked]].map(([l, v, co, rows]) => (
              <KpiCard key={l} label={l} value={v} color={co} onClick={rows?.length ? () => openModal(l, rows) : null} />
            ))}
          </div>
        </Card>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        <Card>
          <div style={{ background: "#4a5568", margin: "-16px -16px 14px", padding: "7px 14px", borderRadius: "8px 8px 0 0" }}><span style={{ color: "#fff", fontWeight: 600, fontSize: 11 }}>Component Build Status</span></div>
          <HSBar data={compData} valueKeys={["complete", "inProgress", "notStarted", "blocked"]} colors={[C.complete, C.inProgress, C.notStarted, C.delayed]} onBarClick={row => openModal(`Component: ${row.name}`, row.rows)} />
          <Leg items={[{ label: "Complete", color: C.complete }, { label: "In Progress", color: C.inProgress }, { label: "Not Started", color: C.notStarted }, { label: "Blocked", color: C.delayed }]} />
        </Card>
        <Card>
          <div style={{ background: "#4a5568", margin: "-16px -16px 14px", padding: "7px 14px", borderRadius: "8px 8px 0 0" }}><span style={{ color: "#fff", fontWeight: 600, fontSize: 11 }}>User Story Build Status by Sprint</span></div>
          <HSBar data={sprintData} valueKeys={["complete", "inProgress", "notStarted", "blocked"]} colors={[C.complete, C.inProgress, C.notStarted, C.delayed]} onBarClick={row => openModal(`Sprint: ${row.name}`, row.rows)} />
          <Leg items={[{ label: "Complete", color: C.complete }, { label: "In Progress", color: C.inProgress }, { label: "Not Started", color: C.notStarted }, { label: "Blocked", color: C.delayed }]} />
        </Card>
      </div>
      {data.blocked.length > 0 && (
        <Card>
          <SecTitle title={`🚨 Blockers (${data.blocked.length})`} color={C.delayed} />
          <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
            {data.blocked.slice(0, 15).map((r, i) => (
              <div key={i} style={{ display: "flex", gap: 10, padding: "9px 13px", background: "#fff5f5", borderRadius: 6, border: `1px solid ${C.delayed}30`, cursor: "pointer" }} onClick={() => openModal("Blocker Detail", [r])}>
                <span style={{ color: C.delayed }}>🔴</span>
                <div style={{ flex: 1 }}>
                  <div style={{ color: C.text, fontSize: 12, fontWeight: 600 }}>{String(r[K.story] || "—").slice(0, 80)}</div>
                  {r[K.blocker] && <div style={{ color: C.muted, fontSize: 11, marginTop: 2 }}>Blocker: {String(r[K.blocker]).slice(0, 100)}</div>}
                </div>
                {r[K.sprint] && <span style={{ color: C.muted, fontSize: 11 }}>{r[K.sprint]}</span>}
              </div>
            ))}
          </div>
        </Card>
      )}
    </div>
  );
}

// ─── TESTING TAB ─────────────────────────────────────────────────────────────
function TestingTab({ data, openModal }) {
  if (!data) return <Empty label="Upload Workplan to view Testing activities." />;
  const rows = data.testRows || [];
  if (!rows.length) return <Empty label="No Testing workstream activities found in the Workplan." />;
  const sMap = {};
  rows.forEach(r => { const s = r["Default Status"] || r["Status"] || "Unknown"; sMap[s] = (sMap[s] || 0) + 1; });
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", gap: 11 }}>
        <KpiCard label="Total Test Activities" value={rows.length} color={C.navyLight} />
        {Object.entries(sMap).slice(0, 4).map(([s, v]) => (
          <KpiCard key={s} label={s} value={v} color={SC[s] || C.muted} onClick={() => openModal(`Testing — ${s}`, rows.filter(r => (r["Default Status"] || r["Status"] || "") === s), WP_COLS)} />
        ))}
      </div>
      <Card><SecTitle title="Testing Activities Detail" color={C.navyLight} /><ActivityTable rows={rows} /></Card>
    </div>
  );
}

// ─── CAPACITY TAB ─────────────────────────────────────────────────────────────
function CapacityTab({ data, openModal }) {
  if (!data) return <Empty label="Upload Capacity Planning sheet to view this section." />;
  const { sprintChart, keys: K } = data;
  const latest = sprintChart[sprintChart.length - 1] || {};
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 11 }}>
        <KpiCard label="Total Resources" value={data.total} color={C.navyLight} />
        <KpiCard label="Available Hours (Latest Sprint)" value={(latest.available || 0).toLocaleString()} color={C.green} />
        <KpiCard label="Planned Hours (Latest Sprint)" value={(latest.planned || 0).toLocaleString()} color={C.navyLight} />
        <KpiCard label="Surplus / Deficit" value={latest.diff != null ? (latest.diff >= 0 ? "+" : "") + latest.diff.toLocaleString() : "—"} color={latest.diff >= 0 ? C.green : C.delayed} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        <Card>
          <SecTitle title="Available vs. Planned by Sprint" color={C.navyLight} />
          <HSBar data={sprintChart.map(d => ({ name: d.name, available: Math.round(d.available), planned: Math.round(d.planned), rows: d.rows }))} valueKeys={["available", "planned"]} colors={[C.navyLight, C.gold]} onBarClick={row => openModal(`Sprint Capacity: ${row.name}`, row.rows)} />
          <Leg items={[{ label: "Available", color: C.navyLight }, { label: "Planned", color: C.gold }]} />
        </Card>
        <Card>
          <SecTitle title="Surplus / Deficit by Sprint" color={C.navyLight} />
          <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
            {sprintChart.map((d, i) => {
              const isD = d.diff < 0; const maxAbs = Math.max(...sprintChart.map(s => Math.abs(s.diff)), 1);
              return (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ minWidth: 70, fontSize: 11, color: C.text }}>{d.name}</div>
                  <div style={{ flex: 1, background: C.border, borderRadius: 3, height: 18, overflow: "hidden", cursor: "pointer" }} onClick={() => openModal(`Sprint: ${d.name}`, d.rows)}>
                    <div style={{ width: `${(Math.abs(d.diff) / maxAbs) * 100}%`, height: "100%", background: isD ? C.delayed : C.green, display: "flex", alignItems: "center", justifyContent: "flex-end", paddingRight: 5 }}>
                      <span style={{ color: "#fff", fontSize: 10, fontWeight: 700 }}>{isD ? "" : "+"}{d.diff}</span>
                    </div>
                  </div>
                  <div style={{ minWidth: 48, color: isD ? C.delayed : C.green, fontSize: 11, fontWeight: 700 }}>{isD ? "" : "+"}{d.diff}</div>
                </div>
              );
            })}
          </div>
          <Leg items={[{ label: "Surplus", color: C.green }, { label: "Deficit", color: C.delayed }]} />
        </Card>
      </div>
      <Card>
        <SecTitle title="Resource Detail" color={C.navyLight} />
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead><tr style={{ background: "#f0f4f8" }}>{[K.resource, K.sprint, K.workstream, K.available, K.planned].filter(Boolean).map(c => <th key={c} style={{ textAlign: "left", padding: "8px 10px", color: C.muted, fontSize: 11, fontWeight: 700, borderBottom: `2px solid ${C.border}` }}>{c}</th>)}</tr></thead>
            <tbody>
              {data.items.slice(0, 50).map((r, i) => (
                <tr key={i} style={{ borderBottom: `1px solid ${C.border}`, background: i % 2 === 0 ? C.white : "#f9fafb" }}>
                  {[K.resource, K.sprint, K.workstream, K.available, K.planned].filter(Boolean).map(k => <td key={k} style={{ padding: "7px 10px", color: C.text }}>{String(r[k] || "—").slice(0, 40)}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}