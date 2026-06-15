// Nova — backend server (documents, images, OCR, replies, server-side memory)
// Holds your secret API key and talks to Anthropic on Nova's behalf.
// The browser only ever talks to THIS server, never directly to Anthropic.

import express from "express";
import "dotenv/config";
import crypto from "crypto";
import { unzipSync, strFromU8 } from "fflate";
import { Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun, BorderStyle } from "docx";
import PDFDocument from "pdfkit";
import { readFileSync, existsSync } from "fs";

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

// ---- Branded document generation (Word + PDF) ----
const BRAND = {
  name: process.env.NOVA_BRAND_NAME || "Nova",
  tagline: process.env.NOVA_BRAND_TAGLINE || "",
  contact: process.env.NOVA_BRAND_CONTACT || "",
};
const C = { navy: "1e3a6b", gold: "c8881d", ink: "16233b", grey: "6b7589" };
const LOGO_PATH = "public/brand-logo.png";
function pngSize(buf) {
  if (buf.length > 24 && buf[0] === 0x89 && buf[1] === 0x50)
    return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  return null;
}
function getLogo() {
  try { if (existsSync(LOGO_PATH)) { const b = readFileSync(LOGO_PATH); const s = pngSize(b); if (s) return { buf: b, ...s }; } } catch (_) {}
  return null;
}
function parseBlocks(body) {
  return (body || "").replace(/\r/g, "").split(/\n{2,}/).map(s => s.trim()).filter(Boolean).map(p => {
    const m = p.match(/^#{1,3}\s+(.*)$/s);
    return m ? { type: "heading", text: m[1].trim() } : { type: "para", text: p };
  });
}
const niceDate = () => new Date().toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" });

async function buildDocx(body) {
  const kids = [];
  const logo = getLogo();
  if (logo) kids.push(new Paragraph({ children: [ new ImageRun({ data: logo.buf, transformation: { width: Math.round(logo.w * (48 / logo.h)), height: 48 } }) ] }));
  kids.push(new Paragraph({ children: [ new TextRun({ text: BRAND.name, bold: true, size: 34, color: C.navy }) ] }));
  if (BRAND.tagline) kids.push(new Paragraph({ children: [ new TextRun({ text: BRAND.tagline, size: 20, color: C.gold }) ] }));
  if (BRAND.contact) kids.push(new Paragraph({ children: [ new TextRun({ text: BRAND.contact, size: 16, color: C.grey }) ] }));
  kids.push(new Paragraph({ border: { bottom: { color: C.gold, space: 1, style: BorderStyle.SINGLE, size: 12 } }, spacing: { after: 180 } }));
  kids.push(new Paragraph({ children: [ new TextRun({ text: niceDate(), size: 18, color: C.grey }) ], spacing: { after: 200 } }));
  for (const b of parseBlocks(body)) {
    if (b.type === "heading") {
      kids.push(new Paragraph({ children: [ new TextRun({ text: b.text, bold: true, size: 26, color: C.navy }) ], spacing: { before: 160, after: 80 } }));
    } else {
      const lines = b.text.split("\n"); const runs = [];
      lines.forEach((ln, i) => { if (i > 0) runs.push(new TextRun({ break: 1 })); runs.push(new TextRun({ text: ln, size: 22, color: C.ink })); });
      kids.push(new Paragraph({ children: runs, spacing: { after: 160 }, alignment: AlignmentType.JUSTIFIED }));
    }
  }
  const doc = new Document({ sections: [ { properties: { page: { margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } } }, children: kids } ] });
  return Packer.toBuffer(doc);
}

function buildPdf(body) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", margin: 56 });
    const chunks = []; doc.on("data", c => chunks.push(c)); doc.on("end", () => resolve(Buffer.concat(chunks))); doc.on("error", reject);
    const W = 595, hex = h => "#" + h;
    doc.rect(0, 0, W, 10).fill(hex(C.navy));
    const logo = getLogo();
    if (logo) { try { doc.image(logo.buf, W - 56 - logo.w * (40 / logo.h), 30, { height: 40 }); } catch (_) {} }
    doc.fillColor(hex(C.navy)).font("Helvetica-Bold").fontSize(22).text(BRAND.name, 56, 44);
    if (BRAND.tagline) doc.fillColor(hex(C.gold)).font("Helvetica").fontSize(11).text(BRAND.tagline, 56, doc.y);
    if (BRAND.contact) doc.fillColor(hex(C.grey)).font("Helvetica").fontSize(9).text(BRAND.contact, 56, doc.y);
    doc.moveDown(0.6);
    const ly = doc.y; doc.moveTo(56, ly).lineTo(W - 56, ly).lineWidth(1.5).strokeColor(hex(C.gold)).stroke();
    doc.moveDown(0.7);
    doc.fillColor(hex(C.grey)).font("Helvetica").fontSize(10).text(niceDate());
    doc.moveDown(0.7);
    for (const b of parseBlocks(body)) {
      if (b.type === "heading") { doc.moveDown(0.3).fillColor(hex(C.navy)).font("Helvetica-Bold").fontSize(13).text(b.text); doc.moveDown(0.2); }
      else { doc.fillColor(hex(C.ink)).font("Helvetica").fontSize(11).text(b.text, { align: "justify", lineGap: 2 }); doc.moveDown(0.6); }
    }
    doc.end();
  });
}

const API_KEY = process.env.ANTHROPIC_API_KEY;
const MODEL = process.env.MAY_MODEL || "claude-sonnet-4-6";
const MEMORY_MODEL = process.env.MAY_MEMORY_MODEL || MODEL; // model that maintains long-term memory

// ---- Passcode lock ----
// Set NOVA_PASSCODE in the environment to require a passcode. If it's unset,
// Nova stays open (so you can't accidentally lock yourself out before setting it).
const PASSCODE = (process.env.NOVA_PASSCODE || "").toString();
const SESSION_TOKEN = PASSCODE
  ? crypto.createHash("sha256").update(PASSCODE + ":nova-session-v1").digest("hex")
  : "";

function safeEqual(a, b) {
  const x = Buffer.from(String(a)), y = Buffer.from(String(b));
  return x.length === y.length && crypto.timingSafeEqual(x, y);
}
function authed(req) {
  if (!PASSCODE) return true; // no passcode configured => open
  return safeEqual(req.headers["x-nova-auth"] || "", SESSION_TOKEN);
}
function guard(req, res, next) {
  if (authed(req)) return next();
  res.status(401).json({ error: "Locked — enter the passcode." });
}

// Does this Nova require a passcode? (so the page knows whether to show the lock)
app.get("/api/authmode", (_req, res) => res.json({ required: !!PASSCODE }));

// Exchange the passcode for a session token the browser stores.
app.post("/api/login", (req, res) => {
  if (!PASSCODE) return res.json({ ok: true, token: "" });
  const given = (req.body?.passcode || "").toString();
  if (!safeEqual(given, PASSCODE)) return res.status(401).json({ ok: false });
  res.json({ ok: true, token: SESSION_TOKEN });
});

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

// ---- Tasks / daily checklist ----
const TASKS_KEY = process.env.NOVA_TASKS_KEY || "nova_tasks";
let tasksCache = []; // in-process fallback

// Local "today" in Johannesburg, as YYYY-MM-DD
function todayStr() {
  return new Date().toLocaleDateString("en-CA", { timeZone: "Africa/Johannesburg" });
}
async function loadTasks() {
  if (UPSTASH_URL && UPSTASH_TOKEN) {
    try {
      const j = await upstash(["GET", TASKS_KEY]);
      if (typeof j?.result === "string" && j.result) tasksCache = JSON.parse(j.result);
    } catch (_) { /* keep cache */ }
  }
  return Array.isArray(tasksCache) ? tasksCache : [];
}
async function saveTasks(list) {
  tasksCache = Array.isArray(list) ? list : [];
  if (UPSTASH_URL && UPSTASH_TOKEN) {
    try { await upstash(["SET", TASKS_KEY, JSON.stringify(tasksCache)]); } catch (_) {}
  }
  return tasksCache;
}
// A task: { id, text, daily:boolean, doneDate:"YYYY-MM-DD"|null }
// Daily tasks count as "done" only if completed today; otherwise they're due again.
function withStatus(t) {
  const today = todayStr();
  const done = t.daily ? t.doneDate === today : !!t.doneDate;
  return { ...t, done };
}

// Tools Nova can call to manage the checklist by voice/chat.
const TASK_TOOLS = [
  {
    name: "add_task",
    description: "Add a task to the user's daily checklist. Use whenever the user asks to add, note, remember, or put something on their checklist or to-do list. For things they do every day, set daily=true so it resets each morning.",
    input_schema: { type: "object", properties: {
      text: { type: "string", description: "The task text" },
      daily: { type: "boolean", description: "True if it repeats every day" }
    }, required: ["text"] },
  },
  {
    name: "complete_task",
    description: "Mark a task on the checklist as done. Match by the words the user uses.",
    input_schema: { type: "object", properties: { text: { type: "string", description: "Text (or part) of the task to mark done" } }, required: ["text"] },
  },
  {
    name: "remove_task",
    description: "Delete a task from the checklist. Match by the words the user uses.",
    input_schema: { type: "object", properties: { text: { type: "string", description: "Text (or part) of the task to delete" } }, required: ["text"] },
  },
];
function findTask(list, q) {
  const s = (q || "").toLowerCase().trim();
  return list.find(t => t.text.toLowerCase() === s)
    || list.find(t => t.text.toLowerCase().includes(s))
    || list.find(t => s.includes(t.text.toLowerCase()));
}
async function toolAddTask(input) {
  const text = (input?.text || "").toString().trim();
  if (!text) return "No task text given.";
  const list = await loadTasks();
  list.push({ id: "t" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), text, daily: !!input?.daily, doneDate: null });
  await saveTasks(list);
  return `Added "${text}"${input?.daily ? " as a daily task" : ""} to the checklist.`;
}
async function toolCompleteTask(input) {
  const list = await loadTasks();
  const t = findTask(list, input?.text);
  if (!t) return `No task matching "${input?.text}" was found.`;
  t.doneDate = todayStr();
  await saveTasks(list);
  return `Marked "${t.text}" as done.`;
}
async function toolRemoveTask(input) {
  let list = await loadTasks();
  const t = findTask(list, input?.text);
  if (!t) return `No task matching "${input?.text}" was found.`;
  list = list.filter(x => x.id !== t.id);
  await saveTasks(list);
  return `Removed "${t.text}" from the checklist.`;
}
async function runTaskTool(name, input) {
  if (name === "add_task") return toolAddTask(input);
  if (name === "complete_task") return toolCompleteTask(input);
  if (name === "remove_task") return toolRemoveTask(input);
  return null;
}

// May's personality + skills live here on the server so they stay consistent.
const PERSONA = `You are Nova, a warm, sharp, and concise assistant who speaks aloud and can also read documents, emails, images, and screenshots.

VOICE: Your replies are read aloud, so keep them natural. For ordinary chat, 1-3 sentences, no markdown, no emoji, no bullet lists.

WHEN THE USER ATTACHES A DOCUMENT, EMAIL, IMAGE, OR SCREENSHOT:
- Summary requests: give the key points clearly and briefly.
- Read / extract / OCR requests: report faithfully what the text says. You can read text inside photos and scanned documents directly.
- Reply requests (to an email or message): draft a reply in a fitting tone. Begin with a short spoken lead-in like "Here's a draft you can read below," then give the full draft. If the intended tone or main points are unclear, ask one quick question first.

WEB: When a question needs current or factual info, use web search, then answer plainly in your own words.

LONG CONTENT: For something long like a full email draft, keep any preamble short (it gets read aloud) and put the substance in the body for the user to read on screen.

DAILY CHECKLIST: The user has a real, saved checklist built into the app (the "Tasks" button). You can SEE today's checklist below when it exists, and you can MANAGE it with your tools: use add_task when they ask to add/note/remember a task (set daily=true for things they do every day), complete_task when they say something's done, and remove_task to delete one. After using a tool, confirm briefly and naturally in one sentence (e.g. "Added that to your checklist"). If they ask what's due, just tell them from the list below. Never write out checklists or HTML/code for tasks in chat.`;

// Main conversation endpoint — accepts full message history (text + image/document/text blocks).
app.post("/api/ask", guard, async (req, res) => {
  try {
    if (!API_KEY) {
      return res.status(500).json({ error: "Server is missing ANTHROPIC_API_KEY." });
    }
    const messages = Array.isArray(req.body?.messages) ? req.body.messages.slice(-16) : null;
    if (!messages || !messages.length) {
      return res.status(400).json({ error: "No messages provided." });
    }

    const memory = (await loadMemory()).trim();
    let system = memory
      ? PERSONA +
        "\n\nWHAT YOU REMEMBER ABOUT THIS USER (from past conversations — use it naturally, don't recite it back):\n" +
        memory
      : PERSONA;

    // Make Nova aware of today's checklist so she can report what's still due.
    try {
      const tasks = (await loadTasks()).map(withStatus);
      if (tasks.length) {
        const due = tasks.filter(t => !t.done).map(t => "- " + t.text + (t.daily ? " (daily)" : ""));
        const done = tasks.filter(t => t.done).map(t => "- " + t.text);
        system += "\n\nTODAY'S CHECKLIST (" + todayStr() + "). If the user asks what's due/outstanding/left, tell them from this list, naturally and briefly.";
        system += "\nSTILL DUE:\n" + (due.length ? due.join("\n") : "(nothing — all done!)");
        if (done.length) system += "\nDONE TODAY:\n" + done.join("\n");
      }
    } catch (_) {}

    const tools = [
      { type: "web_search_20250305", name: "web_search", max_uses: 5 },
      ...TASK_TOOLS,
    ];

    let convo = messages.slice();
    let reply = "";
    let tasksChanged = false;

    for (let step = 0; step < 5; step++) {
      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-api-key": API_KEY,
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({ model: MODEL, max_tokens: 1500, system, messages: convo, tools }),
      });
      const data = await r.json();
      if (data.error) {
        return res.status(502).json({ error: data.error.message || "Anthropic API error." });
      }
      const textPart = (data.content || []).filter(b => b.type === "text").map(b => b.text).join(" ").trim();

      if (data.stop_reason === "tool_use") {
        const results = [];
        for (const b of (data.content || [])) {
          if (b.type !== "tool_use") continue;
          const out = await runTaskTool(b.name, b.input); // null for non-task (server) tools
          if (out != null) { tasksChanged = true; results.push({ type: "tool_result", tool_use_id: b.id, content: out }); }
        }
        if (!results.length) { reply = textPart; break; } // nothing for us to run
        convo.push({ role: "assistant", content: data.content });
        convo.push({ role: "user", content: results });
        continue; // loop for Nova's spoken confirmation
      }
      reply = textPart;
      break;
    }

    res.json({ reply: reply || "Done.", tasksChanged });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Server error reaching the assistant." });
  }
});

// Extract plain text from an uploaded .docx so Nova can read Word documents.
// (Images and PDFs go straight to Claude as image/document blocks — no extraction needed.)
app.post("/api/extract", guard, async (req, res) => {
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
app.get("/api/memory", guard, async (_req, res) => {
  res.json({ memory: await loadMemory(), durable: !!(UPSTASH_URL && UPSTASH_TOKEN) });
});
app.post("/api/memory", guard, async (req, res) => {
  const memory = await saveMemory((req.body?.memory || "").toString());
  res.json({ memory });
});
app.post("/api/memory/clear", guard, async (_req, res) => {
  await saveMemory("");
  res.json({ memory: "" });
});

// Maintain a concise long-term memory of durable facts about the user.
// Called in the background after each exchange; stored on the server.
app.post("/api/remember", guard, async (req, res) => {
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

// Generate a branded Word or PDF document from text.
app.post("/api/document", guard, async (req, res) => {
  try {
    const body = (req.body?.body || "").toString();
    const format = (req.body?.format || "docx").toString();
    const title = (req.body?.title || "").toString();
    if (!body.trim()) return res.status(400).json({ error: "Nothing to put in the document." });
    const slug = (title || body).replace(/[#*]/g, "").trim().split(/\s+/).slice(0, 6).join("-")
      .replace(/[^a-zA-Z0-9-]/g, "").slice(0, 50) || "nova-document";
    if (format === "pdf") {
      const buf = await buildPdf(body);
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", `attachment; filename="${slug}.pdf"`);
      return res.send(buf);
    }
    const buf = await buildDocx(body);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${slug}.docx"`);
    return res.send(buf);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Couldn't create the document." });
  }
});

// ---- Tasks endpoints ----
app.get("/api/tasks", guard, async (_req, res) => {
  const list = await loadTasks();
  res.json({ tasks: list.map(withStatus), today: todayStr() });
});
app.post("/api/tasks", guard, async (req, res) => {
  const text = (req.body?.text || "").toString().trim();
  if (!text) return res.status(400).json({ error: "Empty task." });
  const daily = !!req.body?.daily;
  const list = await loadTasks();
  list.push({ id: "t" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), text, daily, doneDate: null });
  await saveTasks(list);
  res.json({ tasks: list.map(withStatus), today: todayStr() });
});
app.post("/api/tasks/toggle", guard, async (req, res) => {
  const id = (req.body?.id || "").toString();
  const list = await loadTasks();
  const t = list.find(x => x.id === id);
  if (t) { const today = todayStr(); const isDone = t.daily ? t.doneDate === today : !!t.doneDate; t.doneDate = isDone ? null : today; }
  await saveTasks(list);
  res.json({ tasks: list.map(withStatus), today: todayStr() });
});
app.post("/api/tasks/delete", guard, async (req, res) => {
  const id = (req.body?.id || "").toString();
  let list = await loadTasks();
  list = list.filter(x => x.id !== id);
  await saveTasks(list);
  res.json({ tasks: list.map(withStatus), today: todayStr() });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Nova is listening on http://localhost:${PORT}`));
