# May — your voice assistant (v2)

A voice-driven AI assistant that listens, **reads documents/emails/images**, does **OCR on photos and scans**, searches the web, **summarises**, and **drafts replies**.
The browser handles voice and the camera; a small Node server holds your API key and talks to Anthropic.

```
may-app/
├─ server.js          ← backend: chat + Word-doc text extraction
├─ package.json
├─ .env.example       ← copy to .env and add your key
└─ public/
   └─ index.html      ← May's interface (voice, chat, upload, camera)
```

## What's new in v2
- **Attach files** (paperclip) — images, PDFs, Word docs (.docx), and text files.
- **Take a picture** (camera) — snap a document or screen and she reads it.
- **OCR** — she reads text inside photos and scanned PDFs directly (powered by Claude's vision).
- **Summarise / extract / suggest a reply** — quick-action buttons appear when you attach something.
- **Drag & drop** files anywhere onto the page.

## Run her locally
```bash
cd may-app
npm install            # now also installs "mammoth" for Word docs
npm start
```
Open http://localhost:3000.

> Upgrading from v1? Replace your old `server.js`, `package.json`, and `public/index.html`
> with these, then run `npm install` again (for the new dependency) and `npm start`.
> Your existing `.env` key keeps working.

## Notes & limits
- PDFs: up to ~100 pages / ~32 MB per request (Anthropic limit). Big files cost more tokens.
- Images are read by Claude's vision — great for photos of letters, receipts, screenshots, whiteboards.
- The camera and microphone need **HTTPS** (or localhost), which you get automatically once hosted.

## Hosting on your domain
Same as before — deploy the folder to any Node host (Render, Railway, a VPS, or Axxess
cloud/VPS or cPanel "Setup Node.js App"), set the `ANTHROPIC_API_KEY` environment variable,
and point your domain at it. See the hosting steps you were given, or ask.

## Make her yours
- Personality & skills: edit `PERSONA` in `server.js`.
- Model: set `MAY_MODEL` in `.env` (`claude-opus-4-8` = smartest; `claude-haiku-4-5` = cheapest).
- Look: everything visual is in `public/index.html`.

## Keep your key safe
Your key lives only in `.env` / host environment variables — never in `public/`.
Don't commit `.env`. For a public site, add rate limiting on `/api/ask` and `/api/extract`.
