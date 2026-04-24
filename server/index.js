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
import Anthropic from "@anthropic-ai/sdk";
import { fetchAllSheets } from "./smartsheet.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app  = express();
const PORT = process.env.PORT || 3001;

const TOKEN = process.env.SMARTSHEET_TOKEN;
const anthropic = process.env.ANTHROPIC_API_KEY
  ? new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })
  : null;

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

// Returns first-row column names per sheet — helps diagnose title mismatches
app.get("/api/debug/columns", (_req, res) => {
  if (!cache) return res.status(503).json({ error: "No data cached — call /api/data first" });
  const cols = {};
  for (const [sheet, rows] of Object.entries(cache.data)) {
    if (Array.isArray(rows) && rows.length > 0) cols[sheet] = Object.keys(rows[0]);
  }
  res.json(cols);
});

app.get("/api/health", (_req, res) => {
  res.json({
    status: "ok",
    tokenConfigured: !!TOKEN,
    aiConfigured: !!anthropic,
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

app.post("/api/chat", async (req, res) => {
  if (!anthropic) {
    return res.status(503).json({ error: "AI not configured — set ANTHROPIC_API_KEY" });
  }

  const { messages } = req.body;
  if (!Array.isArray(messages) || messages.length === 0) {
    return res.status(400).json({ error: "messages required" });
  }

  // Sanitize: remove empty content blocks — empty text blocks cause API 400 errors
  const clean = messages
    .map(m => ({ role: String(m.role), content: typeof m.content === "string" ? m.content.trim() : "" }))
    .filter(m => (m.role === "user" || m.role === "assistant") && m.content.length > 0);

  // Enforce alternating roles starting with "user"
  const valid = [];
  for (const m of clean) {
    if (valid.length === 0 && m.role !== "user") continue;
    if (valid.length > 0 && m.role === valid[valid.length - 1].role) continue;
    valid.push(m);
  }
  if (valid.length === 0) {
    return res.status(400).json({ error: "No valid messages — must start with a non-empty user message" });
  }

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");

  try {
    const stream = await anthropic.messages.create({
      model: "claude-opus-4-7",
      max_tokens: 1024,
      stream: true,
      system: "You are a helpful assistant embedded in a PMT (Project Management Tool) dashboard. " +
        "You help project teams understand their data including workplans, RAID logs (Risks/Actions/Issues/Decisions), " +
        "change requests, backlog items, and program health metrics sourced from Smartsheet. " +
        "Be concise, practical, and reference specific data when relevant.",
      messages: valid,
    });

    for await (const event of stream) {
      if (event.type === "content_block_delta" && event.delta.type === "text_delta") {
        res.write(`data: ${JSON.stringify({ text: event.delta.text })}\n\n`);
      }
    }
    res.write("data: [DONE]\n\n");
  } catch (err) {
    console.error("[server] /api/chat error:", err.message);
    if (!res.headersSent) {
      res.status(502).json({ error: err.message });
    } else {
      res.write(`data: ${JSON.stringify({ error: err.message })}\n\n`);
    }
  }
  res.end();
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
  console.log(`[server] Anthropic API key: ${anthropic ? "configured ✓" : "NOT SET — chat disabled ⚠"}`);
  if (process.env.NODE_ENV !== "production") {
    console.log(`[server] API: http://localhost:${PORT}/api/data`);
  }
  // Pre-warm cache on startup so first page load is fast
  if (TOKEN) {
    doRefresh().catch(err => console.warn("[server] Pre-warm failed:", err.message));
  }
});
