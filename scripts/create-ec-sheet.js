/**
 * create-ec-sheet.js
 *
 * Run this ONCE from your Windows machine (where Smartsheet API is accessible):
 *   cd pmt-dashboard-sync
 *   node scripts/create-ec-sheet.js
 *
 * What it does:
 *   1. Creates a new "PMT EC Classes" sheet in Smartsheet with all columns
 *   2. Seeds all 22 default EC classes derived from the workbook analysis
 *   3. Adds an "EC Class" column to the existing Test Scenarios sheet
 *   4. Prints the new sheet ID — add it to server/smartsheet.js as SHEETS.ec
 */

import { readFileSync } from "fs";
import { fileURLToPath } from "url";

// Read token from .env without requiring dotenv package
let TOKEN = process.env.SMARTSHEET_TOKEN;
if (!TOKEN) {
  try {
    const envPath = new URL("../.env", import.meta.url);
    const env = readFileSync(envPath, "utf8");
    const m = env.match(/^SMARTSHEET_TOKEN=(.+)$/m);
    if (m) TOKEN = m[1].trim();
  } catch(e) {}
}
if (!TOKEN) { console.error("ERROR: SMARTSHEET_TOKEN not found in .env or environment"); process.exit(1); }
const TEST_SHEET_ID = "2362069488717700"; // existing Test Scenarios sheet

const BASE = "https://api.smartsheet.com/2.0";

const headers = {
  "Authorization": `Bearer ${TOKEN}`,
  "Content-Type": "application/json",
};

async function api(method, path, body) {
  const res = await fetch(`${BASE}${path}`, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });
  const json = await res.json();
  if (!res.ok) throw new Error(`${res.status} ${JSON.stringify(json)}`);
  return json;
}

// ── EC Classes default data ───────────────────────────────────────────────────
const EC_CLASSES = [
  { id:"SF-1",    dim:"PA Form",             label:"CS Performance Assessment",           experiences:"Professional Services", regions:"All", businesses:"Consulting Services",    keyDiff:"5 rating dims; no Internal Client section; Audit Quality = conditional" },
  { id:"SF-2",    dim:"PA Form",             label:"Tax Performance Assessment",          experiences:"Professional Services", regions:"All", businesses:"Tax",                   keyDiff:"5 dims; conditional Audit Quality; adds Tax accreditation/CPA/EA/CE in YE form" },
  { id:"SF-3",    dim:"PA Form",             label:"A&A Performance Assessment",          experiences:"Professional Services", regions:"All", businesses:"Audit & Assurance",     keyDiff:"5 dims; Audit Quality = REQUIRED (not conditional); Q5B field on PMR dashboard" },
  { id:"SF-4",    dim:"PA Form",             label:"ES Performance Assessment",           experiences:"Professional Services", regions:"All", businesses:"Enterprise Solutions",  keyDiff:"5 dims; adds unique Internal Client Section 5 with email capture; AQ = conditional" },
  { id:"SF-5",    dim:"PA Form",             label:"PS Firm Contribution",                experiences:"Professional Services", regions:"All", businesses:"All",                   keyDiff:"Single shared FC form for all 4 PS sub-businesses" },
  { id:"SF-10",   dim:"PA Form",             label:"Intern Performance Assessment",       experiences:"Intern",                regions:"All", businesses:"All",                   keyDiff:"3 dims only; Offer & Staffing Decision section; no 2LR; no TL" },
  { id:"SF-11",   dim:"PA Form",             label:"Project Performance Assessment",      experiences:"Project",               regions:"All", businesses:"All",                   keyDiff:"2 dims; Staffing Question; Engagement Leader field" },
  { id:"SF-12",   dim:"PA Form",             label:"Operations Performance Assessment",   experiences:"Operations",            regions:"All", businesses:"All",                   keyDiff:"2 dims; Incident Reporting section; attestation field" },
  { id:"SF-13",   dim:"PA Form",             label:"Operations Firm Contribution",        experiences:"Operations",            regions:"All", businesses:"All",                   keyDiff:"Separate FC form; auto-complete triggered differently" },
  { id:"ROUTE-1", dim:"Route Map",           label:"Standard 7-Step Route Map",          experiences:"Professional Services, Project, Operations", regions:"All", businesses:"All", keyDiff:"7 steps: Request → TL Attest → TL Input → 2LR → TL Revision → Complete → Cancel" },
  { id:"ROUTE-2", dim:"Route Map",           label:"Intern 5-Step Route Map",            experiences:"Intern",                regions:"All", businesses:"All",                   keyDiff:"5 steps only — no 2LR step, no TL Revision step" },
  { id:"IND-1",   dim:"Individual Dashboard",label:"PS Individual Dashboard",            experiences:"Professional Services", regions:"All", businesses:"All",                   keyDiff:"5 PA dims on visualization; FC Summary + List; 2LR in popup" },
  { id:"IND-2",   dim:"Individual Dashboard",label:"Project Individual Dashboard",       experiences:"Project",               regions:"All", businesses:"All",                   keyDiff:"2 PA dims; no FC section" },
  { id:"IND-3",   dim:"Individual Dashboard",label:"Operations Individual Dashboard",    experiences:"Operations",            regions:"All", businesses:"All",                   keyDiff:"2 PA dims; FC Summary same structure as PS" },
  { id:"COACH-1", dim:"Coach Dashboard",     label:"PS Coach Dashboard",                 experiences:"Professional Services", regions:"All", businesses:"All",                   keyDiff:"5 dims; FC Summary + List; Performance Visual Group; ONL count" },
  { id:"COACH-2", dim:"Coach Dashboard",     label:"Intern Coach Dashboard",             experiences:"Intern",                regions:"All", businesses:"All",                   keyDiff:"Aggregate rating (3 dims); intern-specific PA popup; no FC" },
  { id:"COACH-3", dim:"Coach Dashboard",     label:"Operations Coach Dashboard",         experiences:"Operations",            regions:"All", businesses:"All",                   keyDiff:"2 dims; FC same structure as PS" },
  { id:"PMR-1",   dim:"PMR Dashboard",       label:"PS PMR Dashboard",                   experiences:"Professional Services", regions:"All", businesses:"All",                   keyDiff:"Q5B (A&A); Utilization; Sales; Margin/DNP; Tax/A&A accreditation; USI Integration %" },
  { id:"PMR-2",   dim:"PMR Dashboard",       label:"Intern PMR Dashboard",               experiences:"Intern",                regions:"All", businesses:"All",                   keyDiff:"Aggregate rating; Feedback Providers section; no FC" },
  { id:"PMR-3",   dim:"PMR Dashboard",       label:"Project PMR Dashboard",              experiences:"Project",               regions:"All", businesses:"All",                   keyDiff:"% CSH/TWH coverage; FC Completed/Impact; Late Time Reports; Resume Compliance" },
  { id:"YE-1",    dim:"Year-End Form",       label:"PS Year-End Input Form",             experiences:"Professional Services", regions:"All", businesses:"All",                   keyDiff:"All 101 TM + 59 Coach YE fields; Project/Intern/Ops not in scope" },
  { id:"YE-1a",   dim:"Year-End Form",       label:"Tax + A&A CPA Sub-Class",            experiences:"Professional Services", regions:"All", businesses:"Tax, Audit & Assurance",keyDiff:"Adds 15 CPA/licensure/accreditation fields on top of YE-1" },
  { id:"REG-1",   dim:"Region Process",      label:"USI Non-Audit Intern Offer Decision",experiences:"Intern",                regions:"USI", businesses:"All",                   keyDiff:"IN-010-130: separate subprocess for USI non-audit interns only" },
];

async function run() {
  console.log("PMT EC Sheet Setup\n==================");

  // 1. Create the EC Classes sheet
  console.log("\n[1/3] Creating PMT EC Classes sheet...");
  const sheetDef = {
    name: "PMT EC Classes",
    columns: [
      { title:"Class ID",          type:"TEXT_NUMBER", primary:true },
      { title:"Dimension",         type:"TEXT_NUMBER" },
      { title:"Label",             type:"TEXT_NUMBER" },
      { title:"Experiences",       type:"TEXT_NUMBER" },
      { title:"Regions",           type:"TEXT_NUMBER" },
      { title:"Businesses",        type:"TEXT_NUMBER" },
      { title:"Key Differentiator",type:"TEXT_NUMBER" },
      { title:"Active",            type:"CHECKBOX" },
      { title:"Notes",             type:"TEXT_NUMBER" },
      { title:"Last Updated",      type:"DATE" },
    ],
  };
  const created = await api("POST", "/sheets", sheetDef);
  const ecSheetId = created.result.id;
  console.log(`  ✓ Created sheet ID: ${ecSheetId}`);

  // 2. Fetch column IDs from the new sheet
  console.log("\n[2/3] Seeding EC class rows...");
  const sheet = await api("GET", `/sheets/${ecSheetId}`);
  const colMap = {};
  sheet.columns.forEach(c => { colMap[c.title] = c.id; });

  // Seed rows
  const rows = EC_CLASSES.map(cls => ({
    toBottom: true,
    cells: [
      { columnId: colMap["Class ID"],           value: cls.id },
      { columnId: colMap["Dimension"],           value: cls.dim },
      { columnId: colMap["Label"],               value: cls.label },
      { columnId: colMap["Experiences"],         value: cls.experiences },
      { columnId: colMap["Regions"],             value: cls.regions },
      { columnId: colMap["Businesses"],          value: cls.businesses },
      { columnId: colMap["Key Differentiator"],  value: cls.keyDiff },
      { columnId: colMap["Active"],              value: true },
    ],
  }));

  await api("POST", `/sheets/${ecSheetId}/rows`, rows);
  console.log(`  ✓ Seeded ${EC_CLASSES.length} EC class rows`);

  // 3. Add EC Class column to existing Test Scenarios sheet
  console.log("\n[3/3] Adding EC Class column to Test Scenarios sheet...");
  try {
    const testSheet = await api("GET", `/sheets/${TEST_SHEET_ID}`);
    const alreadyExists = testSheet.columns.some(c => c.title === "EC Class");
    if (alreadyExists) {
      console.log("  ⚠  EC Class column already exists — skipping");
    } else {
      await api("POST", `/sheets/${TEST_SHEET_ID}/columns`, {
        title: "EC Class",
        type: "TEXT_NUMBER",
        index: 3,
      });
      console.log("  ✓ EC Class column added at position 3");
    }
  } catch(e) {
    console.warn("  ⚠  Could not add column to Test sheet:", e.message);
  }

  console.log("\n==================");
  console.log("DONE. Add this to server/smartsheet.js SHEETS object:");
  console.log(`  ec: "${ecSheetId}",`);
  console.log("\nThen add 'ec' to the fetchAllSheets call to expose it via /api/data.");
}

run().catch(err => { console.error("FAILED:", err.message); process.exit(1); });
