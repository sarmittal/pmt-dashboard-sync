/**
 * PMT Dashboard — Express backend
 *
 * Endpoints:
 *   GET  /api/data     → returns cached Smartsheet data (auto-refreshes on first call)
 *   POST /api/refresh  → forces a fresh fetch from Smartsheet, updates cache
 *   GET  /api/health   → liveness check
 *
 * In production (NODE_ENV=production) the server also serves the Vite build
 * from ../client/dist so a single CF app handles both frontend and backend.
 */

import express from "express";
import cors from "cors";
import path from "path";
import { fileURLToPath } from "url";
import { fetchAllSheets } from "./smartsheet.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app  = express();
const PORT = process.env.PORT || 3001;

const TOKEN = process.env.SMARTSHEET_TOKEN;

// In-memory cache
let cache = null;       // { data, fetchedAt }
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
    const data = await fetchAllSheets(TOKEN);
    cache = { data, fetchedAt: new Date().toISOString() };
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
