// Nova — backend server (documents, images, OCR, replies, server-side memory)
// Holds your secret API key and talks to Anthropic on Nova's behalf.
// The browser only ever talks to THIS server, never directly to Anthropic.

import express from "express";
import "dotenv/config";
import { unzipSync, strFromU8 } from "fflate";

const app = express();
app.use(express.json({ limit: "40mb" })); // room for attachments (base64)
app.use(express.static("public"));

// Extract plain text from a .docx (which is a zip containing word/document.xml).
function extractDocxText(buffer) {
  const files = unzipSync(new Uint8Array(buffer));
  const part = files["word/document.xml"];
  if (!part) return "";
  let xml = strFromU8(part)
    .replace(/<\/w:p>/g, "\n")
    .replace(/<w:tab\/>/g, "\t")
    .replace(/<w:br\/?>/g, "\n");
  return xml
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

const API_KEY = process.env.ANTHROPIC_API_KEY;
const MODEL = process.env.MAY_MODEL || "claude-sonnet-4-6";
const MEMORY_MODEL = process.env.MAY_MEMORY_MODEL || MODEL; // model that maintains long-term memory

// ---- Server-side memory storage ----
// Durable when Upstash env vars are set; otherwise falls back to an in-process
// cache (works, but resets when the free Render service sleeps or redeploys).
const UPSTASH_URL = process.env.UPSTASH_REDIS_REST_URL;
const UPSTASH_TOKEN = process.env.UPSTASH_REDIS_REST_TOKEN;
const MEM_KEY = process.env.NOVA_MEMORY_KEY || "nova_memory";
let memCache = ""; // in-process fallback / cache

async function upstash(command) {
  // command is a Redis command as an array, e.g. ["GET","nova_memory"]
  const r = await fetch(UPSTASH_URL, {
    method: "POST",
    headers: { Authorization: `Bearer ${UPSTASH_TOKEN}`, "content-type": "application/json" },
    body: JSON.stringify(command),
  });
  return r.json(); // { result: ... } or { error: ... }
}

async function loadMemory() {
  if (UPSTASH_URL && UPSTASH_TOKEN) {
    try {
      const j = await upstash(["GET", MEM_KEY]);
      memCache = typeof j?.result === "string" ? j.result : "";
    } catch (_) { /* keep cache on error */ }
  }
  return memCache;
}

async function saveMemory(text) {
  memCache = (text || "").toString();
  if (UPSTASH_URL && UPSTASH_TOKEN) {
    try { await upstash(["SET", MEM_KEY, memCache]); } catch (_) { /* keep cache */ }
  }
  return memCache;
}

// May's personality + skills live here on the server so they stay consistent.
const PERSONA = `You are Nova, a warm, sharp, and concise assistant who speaks aloud and can also read documents, emails, images, and screenshots.

VOICE: Your replies are read aloud, so keep them natural. For ordinary chat, 1-3 sentences, no markdown, no emoji, no bullet lists.

WHEN THE USER ATTACHES A DOCUMENT, EMAIL, IMAGE, OR SCREENSHOT:
- Summary requests: give the key points clearly and briefly.
- Read / extract / OCR requests: report faithfully what the text says. You can read text inside photos and scanned documents directly.
- Reply requests (to an email or message): draft a reply in a fitting tone. Begin with a short spoken lead-in like "Here's a draft you can read below," then give the full draft. If the intended tone or main points are unclear, ask one quick question first.

WEB: When a question needs current or factual info, use web search, then answer plainly in your own words.

LONG CONTENT: For something long like a full email draft, keep any preamble short (it gets read aloud) and put the substance in the body for the user to read on screen.`;

// Main conversation endpoint — accepts full message history (text + image/document/text blocks).
app.post("/api/ask", async (req, res) => {
  try {
    if (!API_KEY) {
      return res.status(500).json({ error: "Server is missing ANTHROPIC_API_KEY." });
    }
    const messages = Array.isArray(req.body?.messages) ? req.body.messages.slice(-16) : null;
    if (!messages || !messages.length) {
      return res.status(400).json({ error: "No messages provided." });
    }

    const memory = (await loadMemory()).trim();
    const system = memory
      ? PERSONA +
        "\n\nWHAT YOU REMEMBER ABOUT THIS USER (from past conversations — use it naturally, don't recite it back):\n" +
        memory
      : PERSONA;

    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: 1500,
        system,
        messages,
        tools: [{ type: "web_search_20250305", name: "web_search", max_uses: 5 }],
      }),
    });

    const data = await r.json();
    if (data.error) {
      return res.status(502).json({ error: data.error.message || "Anthropic API error." });
    }
    const reply = (data.content || [])
      .filter((b) => b.type === "text")
      .map((b) => b.text)
      .join(" ")
      .trim();

    res.json({ reply: reply || "Sorry, I didn't catch that — could you say it again?" });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Server error reaching the assistant." });
  }
});

// Extract plain text from an uploaded .docx so Nova can read Word documents.
// (Images and PDFs go straight to Claude as image/document blocks — no extraction needed.)
app.post("/api/extract", async (req, res) => {
  try {
    const { dataBase64, name } = req.body || {};
    if (!dataBase64) return res.status(400).json({ error: "No file data." });
    const buffer = Buffer.from(dataBase64, "base64");
    const text = extractDocxText(buffer);
    res.json({ text, name: name || "document.docx" });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Couldn't read that Word document." });
  }
});

// Read / set / clear Nova's memory directly (used by the memory panel).
app.get("/api/memory", async (_req, res) => {
  res.json({ memory: await loadMemory(), durable: !!(UPSTASH_URL && UPSTASH_TOKEN) });
});
app.post("/api/memory", async (req, res) => {
  const memory = await saveMemory((req.body?.memory || "").toString());
  res.json({ memory });
});
app.post("/api/memory/clear", async (_req, res) => {
  await saveMemory("");
  res.json({ memory: "" });
});

// Maintain a concise long-term memory of durable facts about the user.
// Called in the background after each exchange; stored on the server.
app.post("/api/remember", async (req, res) => {
  const current = await loadMemory();
  try {
    if (!API_KEY) return res.json({ memory: current });
    const exchange = (req.body?.exchange || "").toString().slice(0, 4000);
    if (!exchange.trim()) return res.json({ memory: current });

    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: MEMORY_MODEL,
        max_tokens: 400,
        system:
          "You maintain a short memory of durable facts about a user, for an assistant named Nova. " +
          "Given the CURRENT MEMORY and a NEW EXCHANGE, return an updated memory: a concise list of stable, " +
          "useful facts worth remembering long-term (name, location, job, key preferences, ongoing projects). " +
          "Keep existing facts, add new ones, correct outdated ones. Ignore one-off questions and small talk. " +
          "Keep it under 150 words. Output ONLY the memory text, with no preamble or commentary.",
        messages: [
          {
            role: "user",
            content:
              "CURRENT MEMORY:\n" + (current.trim() || "(empty)") +
              "\n\nNEW EXCHANGE:\n" + exchange +
              "\n\nUpdated memory:",
          },
        ],
      }),
    });
    const data = await r.json();
    if (data.error) return res.json({ memory: current });
    const updated = (data.content || [])
      .filter((b) => b.type === "text")
      .map((b) => b.text)
      .join(" ")
      .trim();
    const saved = await saveMemory(updated || current);
    res.json({ memory: saved });
  } catch (e) {
    res.json({ memory: current });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Nova is listening on http://localhost:${PORT}`));
