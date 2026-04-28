/**
 * PMT Dashboard — Express backend
 *
 * Endpoints:
 *   GET  /api/data     → returns cached Smartsheet data (auto-refreshes on first call)
 *   POST /api/refresh  → forces a fresh fetch from Smartsheet, updates cache
 *   POST /api/update   → writes a single row back to Smartsheet
 *   GET  /api/health   → liveness check
 *
 * In production (NODE_ENV=production) the server also serves the Vite build
 * from ../client/dist so a single CF app handles both frontend and backend.
 */

import "dotenv/config";
import express from "express";
import cors from "cors";
import path from "path";
import { fileURLToPath } from "url";
import { fetchAllSheets, updateRow, EDITABLE } from "./smartsheet.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app  = express();
const PORT = process.env.PORT || 3001;

const TOKEN = process.env.SMARTSHEET_TOKEN;

// In-memory cache
let cache = null;       // { data, columnMaps, fetchedAt }
let refreshing = false; // prevent concurrent fetches

app.use(cors());
app.use(express.json());

// ── Helper ────────────────────────────────────────────────────────────────────
async function doRefresh() {
  if (refreshing) {
    // Wait for the in-flight refresh to finish
    await new Promise(resolve => {
      const iv = setInterval(() => { if (!refreshing) { clearInterval(iv); resolve(); } }, 200);
    });
    return;
  }
  refreshing = true;
  try {
    console.log("[server] Fetching from Smartsheet...");
    const { data, columnMaps } = await fetchAllSheets(TOKEN);
    cache = { data, columnMaps, fetchedAt: new Date().toISOString() };
    console.log("[server] Cache updated:", cache.fetchedAt);
  } finally {
    refreshing = false;
  }
}

// ── Routes ────────────────────────────────────────────────────────────────────


app.get("/api/health", (_req, res) => {
  res.json({
    status: "ok",
    tokenConfigured: !!TOKEN,
    cacheAge: cache ? Math.round((Date.now() - new Date(cache.fetchedAt)) / 1000) + "s" : null,
  });
});

app.get("/api/data", async (req, res) => {
  try {
    if (!cache) await doRefresh();
    res.json(cache.data);
  } catch (err) {
    console.error("[server] /api/data error:", err.message);
    res.status(502).json({ error: err.message });
  }
});

app.post("/api/refresh", async (req, res) => {
  try {
    await doRefresh();
    res.json({
      ok: true,
      fetchedAt: cache.fetchedAt,
      rowCounts: cache.data.meta?.rowCounts,
    });
  } catch (err) {
    console.error("[server] /api/refresh error:", err.message);
    res.status(502).json({ error: err.message });
  }
});

// POST /api/update
// Body: { sheet: "wp"|"raid", rowId: <number>, updates: { "Column Name": value, ... } }
app.post("/api/update", async (req, res) => {
  try {
    const { sheet, rowId, updates } = req.body;

    if (!sheet || !rowId || !updates || typeof updates !== "object") {
      return res.status(400).json({ error: "Request must include sheet, rowId, and updates object" });
    }
    if (!EDITABLE[sheet]) {
      return res.status(400).json({ error: `Sheet "${sheet}" does not support write-back` });
    }
    if (!cache?.columnMaps?.[sheet]) {
      return res.status(503).json({ error: "Cache not loaded — call /api/refresh first" });
    }

    await updateRow(sheet, rowId, updates, cache.columnMaps, TOKEN);
    console.log(`[server] Updated ${sheet} row ${rowId}:`, Object.keys(updates));

    // Patch in-memory cache so the UI reflects the change immediately
    if (cache.data[sheet]) {
      const row = cache.data[sheet].find(r => String(r._rowId) === String(rowId));
      if (row) Object.assign(row, updates);
    }

    res.json({ ok: true });
  } catch (err) {
    console.error("[server] /api/update error:", err.message);
    res.status(502).json({ error: err.message });
  }
});

// ── Serve React build in production ──────────────────────────────────────────
if (process.env.NODE_ENV === "production") {
  const dist = path.join(__dirname, "../client/dist");
  app.use(express.static(dist));
  // SPA fallback — any non-API route returns index.html
  app.get(/^(?!\/api).*/, (_req, res) => {
    res.sendFile(path.join(dist, "index.html"));
  });
}

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`[server] PMT Dashboard backend running on port ${PORT}`);
  console.log(`[server] Smartsheet token: ${TOKEN ? "configured ✓" : "NOT SET ⚠"}`);
  if (process.env.NODE_ENV !== "production") {
    console.log(`[server] API: http://localhost:${PORT}/api/data`);
  }
  // Pre-warm cache on startup so first page load is fast
  if (TOKEN) {
    doRefresh().catch(err => console.warn("[server] Pre-warm failed:", err.message));
  }
});
