// Nova — backend server (documents, images, OCR, replies, server-side memory)
// Holds your secret API key and talks to Anthropic on Nova's behalf.
// The browser only ever talks to THIS server, never directly to Anthropic.

import express from "express";
import "dotenv/config";
import crypto from "crypto";
import { unzipSync, strFromU8 } from "fflate";
import webpush from "web-push";
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
const MODEL = process.env.MAY_MODEL || "claude-haiku-4-5-20251001";
const MEMORY_MODEL = process.env.MAY_MEMORY_MODEL || MODEL; // model that maintains long-term memory
const MAX_MSGS = parseInt(process.env.NOVA_MAX_MSGS || "20", 10); // cap recent turns sent to the API to control cost
const ANTHROPIC_URL = process.env.ANTHROPIC_BASE_URL || "https://api.anthropic.com/v1/messages";

// Call Anthropic with a hard timeout so a slow/stalled request can never hang the server.
async function callAnthropic(payload, timeoutMs) {
  const ms = Number(process.env.ANTHROPIC_TIMEOUT_MS) || timeoutMs || 40000;
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), ms);
  try {
    const r = await fetch(ANTHROPIC_URL, {
      method: "POST",
      headers: { "content-type": "application/json", "x-api-key": API_KEY, "anthropic-version": "2023-06-01" },
      body: JSON.stringify(payload),
      signal: ctrl.signal,
    });
    return await r.json();
  } finally {
    clearTimeout(timer);
  }
}

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

// Open health check — visit /api/health in a browser to see what's live (no secrets).
app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    version: "2026-06-15-diag-v3",
    model: MODEL,
    hasKey: !!API_KEY,
    passcodeSet: !!PASSCODE,
    upstash: !!(UPSTASH_URL && UPSTASH_TOKEN),
    webSearch: process.env.NOVA_WEB_SEARCH !== "off",
  });
});

// Self-test the AI connection — visit /api/selftest?key=YOURPASSCODE in a browser.
// Does one tiny real call (no tools) and reports success + how long it took.
app.get("/api/selftest", async (req, res) => {
  if (PASSCODE && (req.query.key || "") !== PASSCODE) return res.status(401).json({ ok: false, error: "Add ?key=YOUR_PASSCODE to the URL." });
  if (!API_KEY) return res.json({ ok: false, error: "No ANTHROPIC_API_KEY set on the server." });
  const t0 = Date.now();
  try {
    const data = await callAnthropic({ model: MODEL, max_tokens: 16, messages: [{ role: "user", content: "Say OK." }] }, 15000);
    const ms = Date.now() - t0;
    if (data.error) return res.json({ ok: false, ms, error: data.error.message || "API error", type: data.error.type });
    const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join(" ").trim();
    res.json({ ok: true, ms, model: MODEL, reply: text });
  } catch (e) {
    res.json({ ok: false, ms: Date.now() - t0, error: e.name === "AbortError" ? "Timed out after 15s" : String(e.message || e) });
  }
});

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
  if (name === "add_calendar_event") {
    const r = await createCalendarEvent(input || {});
    if (r.ok) return "Added to your Google Calendar" + (r.invited && r.invited.length ? (" and invited " + r.invited.join(", ")) : "") + ".";
    return "I couldn't add that to your calendar: " + r.error;
  }
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

DAILY CHECKLIST: The user has a real, saved checklist built into the app (the "Tasks" button). You can SEE today's checklist below when it exists, and you can MANAGE it with your tools: use add_task when they ask to add/note/remember a task (set daily=true for things they do every day), complete_task when they say something's done, and remove_task to delete one. After using a tool, confirm briefly and naturally in one sentence (e.g. "Added that to your checklist"). If they ask what's due, just tell them from the list below. Never write out checklists or HTML/code for tasks in chat.

CALENDAR: When Google Calendar is connected, you can add events with the add_calendar_event tool, and invite people by passing their email addresses as attendees (they get an invite automatically). Use it when the user asks to schedule, book, or set up a meeting/event at a specific time. Work out the exact date (YYYY-MM-DD) from words like "tomorrow" or "next Tuesday" using today's date, and always include the year. If they want to invite someone but haven't given an email, ask for it. Today's date is provided below. After adding, confirm in one short sentence. If a calendar action fails because it isn't connected, tell them to tap "Connect Google Calendar" in the Tasks panel.

WHATSAPP: You can't read or send WhatsApp messages directly, but you're great at drafting replies. When the user wants to reply to someone on WhatsApp, write the message in their voice; a "Send on WhatsApp" button appears under your reply so they can send it themselves.

SENDABLE DRAFTS: Whenever your reply contains a message the user will SEND to someone else (a WhatsApp reply, an email body, a text message), wrap ONLY that sendable message between the markers [[SEND]] and [[/SEND]]. Put your own chit-chat, options, or notes OUTSIDE the markers. Put nothing inside the markers except the exact words to send — no "Hi, here's a draft", no sign-off from you, no quotation marks around it. Example: Sure! [[SEND]]Hi Sam, confirming our 10am meeting tomorrow. See you then.[[/SEND]] Want me to make it warmer? — This way the user can copy or WhatsApp just the message itself. If your reply is NOT a message to send on, don't use the markers.`;

// Main conversation endpoint — accepts full message history (text + image/document/text blocks).
// Core chat pipeline, shared by /api/ask and the diagnostic /api/asktest.
async function generateReply(messages) {
  const memory = (await loadMemory()).trim();
  let system = memory
    ? PERSONA +
      "\n\nWHAT YOU REMEMBER ABOUT THIS USER (from past conversations — use it naturally, don't recite it back):\n" +
      memory
    : PERSONA;

  // Always tell Nova today's date (Johannesburg) and weekday, so she can resolve "tomorrow", "next Tuesday", etc.
  const todayJ = todayStr();
  const weekday = new Date(todayJ + "T12:00:00").toLocaleDateString("en-GB", { weekday: "long" });
  system += "\n\nTODAY is " + weekday + ", " + todayJ + " (Africa/Johannesburg). Use this for any relative dates.";

  try {
    const tasks = (await loadTasks()).map(withStatus);
    if (tasks.length) {
      const due = tasks.filter(t => !t.done).map(t => "- " + t.text + (t.due ? " [at " + t.due + "]" : "") + (t.daily ? " (daily)" : ""));
      const done = tasks.filter(t => t.done).map(t => "- " + t.text);
      system += "\n\nTODAY'S CHECKLIST (" + todayStr() + "). If the user asks what's due/outstanding/left, tell them from this list, naturally and briefly.";
      system += "\nSTILL DUE:\n" + (due.length ? due.join("\n") : "(nothing — all done!)");
      if (done.length) system += "\nDONE TODAY:\n" + done.join("\n");
    }
  } catch (_) {}

  const tools = [];
  if (process.env.NOVA_WEB_SEARCH !== "off") tools.push({ type: "web_search_20250305", name: "web_search", max_uses: 5 });
  tools.push(...TASK_TOOLS);
  if (GOOGLE_ON) tools.push({
    name: "add_calendar_event",
    description: "Add an event to the user's Google Calendar, optionally inviting people by email (they receive an invite). Use when the user asks to schedule, book, or set up a meeting or event at a specific date and time. Always include the year in the date.",
    input_schema: {
      type: "object",
      properties: {
        title: { type: "string", description: "Short title of the event" },
        date: { type: "string", description: "Event date as YYYY-MM-DD" },
        time: { type: "string", description: "Start time in 24-hour HH:MM" },
        durationMins: { type: "number", description: "Length in minutes (default 30)" },
        attendees: { type: "string", description: "Comma-separated email addresses to invite (optional)" },
        description: { type: "string", description: "Optional notes/agenda" },
      },
      required: ["title", "date", "time"],
    },
  });

  let convo = messages.slice(-MAX_MSGS);
  let reply = "";
  let tasksChanged = false;
  let lastStop = "";

  for (let step = 0; step < 4; step++) {
    const data = await callAnthropic({ model: MODEL, max_tokens: 1500, system, messages: convo, tools }, 30000);
    if (data.error) { const err = new Error(data.error.message || "Anthropic API error."); err.apiType = data.error.type; throw err; }
    lastStop = data.stop_reason || "";
    const textPart = (data.content || []).filter(b => b.type === "text").map(b => b.text).join(" ").trim();

    if (data.stop_reason === "pause_turn") { convo.push({ role: "assistant", content: data.content }); continue; }

    if (data.stop_reason === "tool_use") {
      const results = [];
      for (const b of (data.content || [])) {
        if (b.type !== "tool_use") continue;
        const out = await runTaskTool(b.name, b.input);
        if (out != null) { tasksChanged = true; results.push({ type: "tool_result", tool_use_id: b.id, content: out }); }
      }
      if (!results.length) { reply = textPart; break; }
      convo.push({ role: "assistant", content: data.content });
      convo.push({ role: "user", content: results });
      continue;
    }
    reply = textPart;
    break;
  }
  return { reply: reply || "Done.", tasksChanged, lastStop };
}

app.post("/api/ask", guard, async (req, res) => {
  try {
    if (!API_KEY) return res.status(500).json({ error: "Server is missing ANTHROPIC_API_KEY." });
    const messages = Array.isArray(req.body?.messages) ? req.body.messages.slice(-16) : null;
    if (!messages || !messages.length) return res.status(400).json({ error: "No messages provided." });
    const out = await generateReply(messages);
    res.json({ reply: out.reply, tasksChanged: out.tasksChanged });
  } catch (e) {
    console.error(e);
    res.status(502).json({ error: e.message || "Server error reaching the assistant." });
  }
});

// Diagnostic: runs the FULL chat pipeline (tools + system) from a browser.
// Visit /api/asktest?key=YOURPASSCODE&q=hello  — shows the result or the exact error.
app.get("/api/asktest", async (req, res) => {
  if (PASSCODE && (req.query.key || "") !== PASSCODE) return res.status(401).json({ ok: false, error: "Add ?key=YOUR_PASSCODE to the URL." });
  const q = (req.query.q || "hello").toString();
  const t0 = Date.now();
  try {
    const out = await generateReply([{ role: "user", content: q }]);
    res.json({ ok: true, ms: Date.now() - t0, q, reply: out.reply, tasksChanged: out.tasksChanged, lastStop: out.lastStop });
  } catch (e) {
    res.json({ ok: false, ms: Date.now() - t0, q, error: e.message || String(e), apiType: e.apiType || null });
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

    const data = await callAnthropic({
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
    }, 20000);
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
  const due = (req.body?.due || "").toString().trim() || null; // "HH:MM" for daily, or "YYYY-MM-DDTHH:MM" for one-off
  const list = await loadTasks();
  list.push({ id: "t" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), text, daily, due, doneDate: null });
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

// Build an iCalendar (.ics) event with an alarm, for a task that has a due time.
function icsEscape(s) { return (s || "").replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n"); }
function pad(n) { return String(n).padStart(2, "0"); }
function buildICS(task) {
  if (!task.due) return null;
  // Determine the wall-clock start (floating local time, so the phone rings at that clock time).
  let datePart, timePart;
  if (task.daily) {
    // "HH:MM" repeating daily, first occurrence today (Johannesburg date)
    datePart = todayStr().replace(/-/g, "");
    timePart = task.due.replace(":", "") + "00";
  } else {
    // "YYYY-MM-DDTHH:MM"
    const [d, t] = task.due.split("T");
    if (!d || !t) return null;
    datePart = d.replace(/-/g, "");
    timePart = t.replace(":", "") + "00";
  }
  const dtStart = datePart + "T" + timePart; // floating local time
  const now = new Date();
  const stamp = now.getUTCFullYear() + pad(now.getUTCMonth() + 1) + pad(now.getUTCDate()) + "T" +
    pad(now.getUTCHours()) + pad(now.getUTCMinutes()) + pad(now.getUTCSeconds()) + "Z";
  const lines = [
    "BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Nova//Checklist//EN", "CALSCALE:GREGORIAN",
    "BEGIN:VEVENT",
    "UID:" + task.id + "@nova.naicker.cc",
    "DTSTAMP:" + stamp,
    "DTSTART:" + dtStart,
    "DURATION:PT15M",
    "SUMMARY:" + icsEscape(task.text),
  ];
  if (task.daily) lines.push("RRULE:FREQ=DAILY");
  lines.push("BEGIN:VALARM", "ACTION:DISPLAY", "DESCRIPTION:" + icsEscape(task.text), "TRIGGER:PT0S", "END:VALARM");
  lines.push("END:VEVENT", "END:VCALENDAR");
  return lines.join("\r\n");
}

// Download a task as a calendar event. Token passed as ?t= since this opens directly in the browser.
app.get("/api/tasks/ics", async (req, res) => {
  if (PASSCODE && (req.query.t || "") !== SESSION_TOKEN) return res.status(401).send("Locked.");
  const id = (req.query.id || "").toString();
  const list = await loadTasks();
  const task = list.find(x => x.id === id);
  if (!task) return res.status(404).send("Task not found.");
  const ics = buildICS(task);
  if (!ics) return res.status(400).send("This task has no time set.");
  const fname = (task.text || "reminder").replace(/[^a-zA-Z0-9]+/g, "-").slice(0, 40) || "reminder";
  res.setHeader("Content-Type", "text/calendar; charset=utf-8");
  res.setHeader("Content-Disposition", `inline; filename="${fname}.ics"`);
  res.send(ics);
});

// ---- Push notifications ----
const VAPID_PUBLIC = process.env.NOVA_VAPID_PUBLIC || "";
const VAPID_PRIVATE = process.env.NOVA_VAPID_PRIVATE || "";
const VAPID_SUBJECT = process.env.NOVA_VAPID_SUBJECT || "mailto:nova@naicker.cc";
const PUSH_ON = !!(VAPID_PUBLIC && VAPID_PRIVATE);
if (PUSH_ON) { try { webpush.setVapidDetails(VAPID_SUBJECT, VAPID_PUBLIC, VAPID_PRIVATE); } catch (e) { console.error("VAPID setup failed", e); } }
const SUBS_KEY = process.env.NOVA_SUBS_KEY || "nova_push_subs";
let subsCache = [];
async function loadSubs() {
  if (UPSTASH_URL && UPSTASH_TOKEN) {
    try { const j = await upstash(["GET", SUBS_KEY]); if (typeof j?.result === "string" && j.result) subsCache = JSON.parse(j.result); } catch (_) {}
  }
  return Array.isArray(subsCache) ? subsCache : [];
}
async function saveSubs(list) {
  subsCache = Array.isArray(list) ? list : [];
  if (UPSTASH_URL && UPSTASH_TOKEN) { try { await upstash(["SET", SUBS_KEY, JSON.stringify(subsCache)]); } catch (_) {} }
  return subsCache;
}
async function sendPush(payload) {
  const subs = await loadSubs();
  if (!subs.length) return { sent: 0 };
  let sent = 0; const survivors = [];
  for (const s of subs) {
    try { await webpush.sendNotification(s, JSON.stringify(payload)); sent++; survivors.push(s); }
    catch (e) { if (e.statusCode === 404 || e.statusCode === 410) { /* expired — drop it */ } else { survivors.push(s); } }
  }
  if (survivors.length !== subs.length) await saveSubs(survivors);
  return { sent };
}

// Give the browser the public key it needs to subscribe.
app.get("/api/push/key", (_req, res) => res.json({ key: VAPID_PUBLIC, enabled: PUSH_ON }));

// Save a device's push subscription.
app.post("/api/push/subscribe", guard, async (req, res) => {
  const sub = req.body?.subscription;
  if (!sub || !sub.endpoint) return res.status(400).json({ error: "No subscription." });
  const subs = await loadSubs();
  if (!subs.find(s => s.endpoint === sub.endpoint)) { subs.push(sub); await saveSubs(subs); }
  res.json({ ok: true, count: subs.length });
});

// Send a test notification to all devices.
app.post("/api/push/test", guard, async (_req, res) => {
  if (!PUSH_ON) return res.status(400).json({ error: "Push not configured on the server." });
  const r = await sendPush({ title: "Nova", body: "Push notifications are working 🎉" });
  res.json(r);
});

// The pinger calls this every few minutes. It nudges any task that's due and not done,
// re-nudging every ~10 minutes until completed. Token via ?key= so a cron service can call it.
app.get("/api/push/check", async (req, res) => {
  if (PASSCODE && (req.query.key || "") !== PASSCODE) return res.status(401).json({ error: "Add ?key=PASSCODE" });
  if (!PUSH_ON) return res.json({ ok: false, error: "push off" });
  const now = new Date();
  const todJ = todayStr();
  // current Johannesburg wall-clock minutes
  const hm = now.toLocaleTimeString("en-GB", { timeZone: "Africa/Johannesburg", hour: "2-digit", minute: "2-digit" });
  const list = await loadTasks();
  let nudged = 0, changed = false;
  for (const t of list) {
    if (!t.due) continue;
    const isDone = t.daily ? t.doneDate === todJ : !!t.doneDate;
    if (isDone) continue;
    // figure out the task's due time-of-day and whether it's reached
    let dueHM = null, dueDate = null;
    if (t.daily) { dueHM = t.due; dueDate = todJ; }
    else { const [d, tm] = (t.due || "").split("T"); dueDate = d; dueHM = tm; }
    if (!dueHM) continue;
    if (dueDate !== todJ) continue;          // only nudge on the due date (daily = today)
    if (hm < dueHM) continue;                 // not yet due
    const last = t.lastNudge || 0;
    if (now.getTime() - last < 10 * 60 * 1000) continue; // throttle: every 10 min
    await sendPush({ title: "Nova reminder", body: t.text, tag: t.id });
    t.lastNudge = now.getTime(); changed = true; nudged++;
  }
  if (changed) await saveTasks(list);
  res.json({ ok: true, nudged, time: hm });
});

// ---- Google Calendar integration ----
const GOOGLE_CLIENT_ID = process.env.NOVA_GOOGLE_CLIENT_ID || "";
const GOOGLE_CLIENT_SECRET = process.env.NOVA_GOOGLE_CLIENT_SECRET || "";
const GOOGLE_REDIRECT = process.env.NOVA_GOOGLE_REDIRECT || "https://nova.naicker.cc/api/google/callback";
const GOOGLE_ON = !!(GOOGLE_CLIENT_ID && GOOGLE_CLIENT_SECRET);
const G_AUTH = (process.env.GOOGLE_AUTH_BASE || "https://accounts.google.com") + "/o/oauth2/v2/auth";
const G_TOKEN = (process.env.GOOGLE_OAUTH_BASE || "https://oauth2.googleapis.com") + "/token";
const G_API = process.env.GOOGLE_API_BASE || "https://www.googleapis.com";
const GKEY = process.env.NOVA_GOOGLE_KEY || "nova_google";
const GCAL_SCOPE = "https://www.googleapis.com/auth/calendar.events";
const TZONE = "Africa/Johannesburg";
let googleCache = null;

async function loadGoogle() {
  if (UPSTASH_URL && UPSTASH_TOKEN) {
    try { const j = await upstash(["GET", GKEY]); if (typeof j?.result === "string" && j.result) googleCache = JSON.parse(j.result); } catch (_) {}
  }
  return googleCache;
}
async function saveGoogle(obj) {
  googleCache = obj;
  if (UPSTASH_URL && UPSTASH_TOKEN) { try { await upstash(["SET", GKEY, JSON.stringify(obj)]); } catch (_) {} }
  return obj;
}
async function googleAccessToken() {
  const g = await loadGoogle();
  if (!g || !g.refresh_token) return null;
  if (g.access_token && g.expiry && Date.now() < g.expiry - 60000) return g.access_token;
  const body = new URLSearchParams({ client_id: GOOGLE_CLIENT_ID, client_secret: GOOGLE_CLIENT_SECRET, refresh_token: g.refresh_token, grant_type: "refresh_token" });
  const r = await fetch(G_TOKEN, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body });
  const j = await r.json();
  if (j.access_token) { g.access_token = j.access_token; g.expiry = Date.now() + (j.expires_in || 3600) * 1000; await saveGoogle(g); return g.access_token; }
  return null;
}

app.get("/api/google/status", guard, async (_req, res) => {
  const g = await loadGoogle();
  res.json({ configured: GOOGLE_ON, connected: !!(g && g.refresh_token) });
});

app.get("/api/google/connect", (req, res) => {
  if (!GOOGLE_ON) return res.status(400).send("Google Calendar isn't configured on the server yet.");
  if (PASSCODE && (req.query.t || "") !== SESSION_TOKEN) return res.status(401).send("Locked.");
  const u = new URL(G_AUTH);
  u.searchParams.set("client_id", GOOGLE_CLIENT_ID);
  u.searchParams.set("redirect_uri", GOOGLE_REDIRECT);
  u.searchParams.set("response_type", "code");
  u.searchParams.set("scope", GCAL_SCOPE);
  u.searchParams.set("access_type", "offline");
  u.searchParams.set("prompt", "consent");
  u.searchParams.set("state", SESSION_TOKEN || "nostate");
  res.redirect(u.toString());
});

app.get("/api/google/callback", async (req, res) => {
  try {
    if (PASSCODE && (req.query.state || "") !== SESSION_TOKEN) return res.status(401).send("State mismatch.");
    const code = (req.query.code || "").toString();
    if (!code) return res.redirect("/?gcal=error");
    const body = new URLSearchParams({ code, client_id: GOOGLE_CLIENT_ID, client_secret: GOOGLE_CLIENT_SECRET, redirect_uri: GOOGLE_REDIRECT, grant_type: "authorization_code" });
    const r = await fetch(G_TOKEN, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body });
    const j = await r.json();
    if (j.refresh_token) {
      await saveGoogle({ refresh_token: j.refresh_token, access_token: j.access_token, expiry: Date.now() + (j.expires_in || 3600) * 1000 });
    } else if (j.access_token) {
      const g = (await loadGoogle()) || {}; g.access_token = j.access_token; g.expiry = Date.now() + (j.expires_in || 3600) * 1000; await saveGoogle(g);
    } else {
      return res.redirect("/?gcal=error");
    }
    res.redirect("/?gcal=connected");
  } catch (e) { console.error(e); res.redirect("/?gcal=error"); }
});

async function createCalendarEvent({ title, date, time, durationMins, attendees, description }) {
  const token = await googleAccessToken();
  if (!token) return { ok: false, error: "Google Calendar isn't connected." };
  if (!title || !date || !time) return { ok: false, error: "I need a title, date and time." };
  const dur = Number(durationMins) || 30;
  const [hh, mm] = time.split(":").map(Number);
  const endMins = hh * 60 + mm + dur;
  const eh = String(Math.floor(endMins / 60) % 24).padStart(2, "0"), em = String(endMins % 60).padStart(2, "0");
  let atts = [];
  if (Array.isArray(attendees)) atts = attendees;
  else if (typeof attendees === "string" && attendees.trim()) atts = attendees.split(/[,;\s]+/).filter(Boolean);
  const ev = {
    summary: title, description: description || "",
    start: { dateTime: date + "T" + time + ":00", timeZone: TZONE },
    end: { dateTime: date + "T" + eh + ":" + em + ":00", timeZone: TZONE },
  };
  if (atts.length) ev.attendees = atts.map(e => ({ email: e }));
  const r = await fetch(G_API + "/calendar/v3/calendars/primary/events?sendUpdates=all", {
    method: "POST", headers: { authorization: "Bearer " + token, "content-type": "application/json" }, body: JSON.stringify(ev),
  });
  const j = await r.json();
  if (j.id) return { ok: true, link: j.htmlLink, invited: atts };
  return { ok: false, error: (j.error && j.error.message) || "Calendar API error." };
}

app.post("/api/calendar/add", guard, async (req, res) => {
  res.json(await createCalendarEvent(req.body || {}));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Nova is listening on http://localhost:${PORT}`));
