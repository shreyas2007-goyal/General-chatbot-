# 💬 General Chatbot — Multi-Provider Edition

A single-file Streamlit chatbot that lets you talk to **five different AI backends**
from one interface, **upload and chat about multiple files of almost any type**
(including images, PowerPoint, and Excel), and **keeps a permanent, searchable
history** of every past conversation.

## Features

- 🎨 **Custom dark + warm-coral theme**: chat bubbles with avatars, timestamps,
  and color-coded provider identity (not the default Streamlit look)
- 🔌 **Switch providers on the fly**: Ollama (local), Gemini, OpenAI, Anthropic, Groq —
  organized in a clean collapsible sidebar
- 📄 **Multi-document Q&A**: upload as many files as you want at once (PDF, Word,
  PowerPoint, Excel, images, TXT, Markdown, Python) — all stay active together,
  the same way ChatGPT handles multiple attachments
- 🖼️ **Vision support**: images, and rendered PowerPoint slides / Excel sheets,
  are sent directly to vision-capable models so they can actually "see" them
- 🕘 **Persistent chat history**: every conversation is saved to a local SQLite
  file and survives closing/reopening the app — browse past chats in a
  ChatGPT-style "Recent" list, titled automatically from your first message
- ➕ **New chat button**: start fresh anytime without losing the previous conversation
- 🔎 **Auto web search**: detects when your question needs current info and tries
  a live search before answering (see honest limitations below)
- 🔑 API keys entered in the sidebar (never hardcoded, never saved to disk)
- 🪶 Single Python file — easy to read, easy to extend

## Project Structure

```
general_chatbot/
├── app.py             # the entire app
├── requirements.txt   # Python dependencies
├── README.md          # this file
└── chat_history.db    # created automatically on first run — your saved chats
```

## Setup

### 1. Install Python dependencies

```bash
pip install -r requirements.txt --break-system-packages
```

(Drop `--break-system-packages` if you're using a virtual environment, e.g. your `penv`.)

### 2. Install LibreOffice (required for PowerPoint/Excel visual rendering)

PowerPoint and Excel files get their text extracted either way, but to let
vision-capable models actually **see** slides and sheets (charts, colors, layout),
the app converts them to images using LibreOffice running headless.

- **Windows/Mac:** download from https://www.libreoffice.org/download
- **Linux:** `sudo apt install libreoffice` (or your distro's equivalent)

After installing, confirm it's on your PATH:

```bash
soffice --version
```

If this command isn't found, the app will still work — it just falls back to
text-only extraction for PPTX/XLSX files and shows a warning explaining why.

### 3. (Optional) Set up Ollama for local models

If you want to use the local "Ollama" provider:

```bash
# Install Ollama: https://ollama.com/download
ollama serve              # start the Ollama server
ollama pull llama3.2       # a solid general-purpose text model
ollama pull llava          # a vision-capable model, if you want image support locally
```

The app auto-detects which models you have pulled and lists them in the dropdown.
**Only vision-capable Ollama models** (names containing `llava`, `vision`,
`bakllava`, or `moondream`) can actually understand images — if you pick a
text-only model like `llama3.2`, the sidebar shows a note explaining that
uploaded images won't be understood until you switch models.

### 4. Get API keys (only needed for the cloud providers you actually use)

| Provider  | Where to get a key |
|-----------|---------------------|
| Gemini    | https://aistudio.google.com/apikey |
| OpenAI    | https://platform.openai.com/api-keys |
| Anthropic | https://console.anthropic.com/settings/keys |
| Groq      | https://console.groq.com/keys |

You paste these into the sidebar when you select that provider — they live only in
your browser session for that run, nothing is written to disk. (Gemini, OpenAI,
and Anthropic's standard models all support vision out of the box. On Groq,
only specific models like `meta-llama/llama-4-scout-17b-16e-instruct` support vision.)

### 5. Run the app

```bash
streamlit run app.py
```

It'll open at `http://localhost:8501`.

## How to Use

1. In the sidebar, open **🔌 Provider & model** and pick a provider and model.
2. If using a cloud provider, **paste your API key** in that same section.
3. *(Optional)* Open **📄 Documents** and upload one or more files — PDF, DOCX,
   PPTX, XLSX, images, TXT, MD, or PY. They all stay active together; ask
   things like "summarize all of these" or "what does the chart on slide 3 show?"
4. **Type your message** and hit enter.
5. Each reply shows a colored avatar matching the active provider, a timestamp,
   and a "live search used" indicator when auto web search kicked in.
6. Click **➕ New chat** anytime to start fresh — your previous conversation is
   automatically saved and reappears under **🕘 Recent** in the sidebar, titled
   from its first message. Click any past chat to reopen it, or the ✕ next to
   it to delete it permanently.

## How Multi-Document Chat Works

Every file you upload is processed immediately and added to an "active documents"
list shown in the sidebar (with a ✕ to remove any single one, or all at once).
On every message you send, the app combines:

- **Extracted text** from every active document (PDF/DOCX/PPTX/XLSX/TXT/MD/PY),
  labeled by filename, into one combined context block
- **Images** from every active document (standalone image uploads, plus rendered
  slides/sheets) sent directly to the model — *only if* your selected
  provider/model actually supports vision (see the Vision section below)

This is a **direct text + image injection** approach, not true RAG with vector
search/chunking — simplest and most reliable for a handful of small-to-medium
files, but it won't scale well to huge documents or dozens of files at once.
Each document's text is capped at ~15,000 characters, and a max of 10 documents
can be active simultaneously (both adjustable in the constants at the top of `app.py`).

## How Vision Support Works

- **Images** (PNG/JPG): sent directly to the model, no text extraction needed.
- **PowerPoint (PPTX)**: text is extracted with python-pptx (titles + body text
  per slide) *and*, if LibreOffice is installed, every slide is also rendered as
  an image so the model can see charts, diagrams, and layout — not just text.
- **Excel (XLSX)**: text is extracted with openpyxl (every cell, sheet by sheet)
  *and*, if LibreOffice is installed, each sheet is rendered as an image too —
  useful for colored cells, conditional formatting, or embedded charts.
- **Vision-capable providers**: Gemini, OpenAI, and Anthropic all support image
  input on their standard models, no extra setup needed. Ollama only supports
  it on specific multimodal models (`llava`, `llama3.2-vision`, etc.) — using a
  text-only Ollama model with image uploads will just skip the images (with a
  note in the chat) and proceed using extracted text only. Groq similarly only
  supports vision on specific models (e.g. `meta-llama/llama-4-scout-17b-16e-instruct`).

## How Persistent Chat History Works

Every conversation is stored in `chat_history.db`, a SQLite file created
automatically next to `app.py` on first run. After every message exchange, the
current chat is saved (or updated) under a unique ID. The **🕘 Recent** section
in the sidebar lists every saved chat, most-recently-updated first, titled from
the first message — click one to reload it, or the ✕ to delete it permanently.
This means your chats survive closing the app, restarting your computer, etc.
If you ever want to start over completely, just delete `chat_history.db` — a
fresh one will be created automatically next time you run the app.

## How Auto Web Search Works

When you send a message, the app scans it for keywords that suggest you want
current information — words like "latest," "today," "current," "news," "weather,"
or a recent year. If matched, it tries a live search via DuckDuckGo's results
page and feeds the top snippets into the model's context before it answers.

**⚠️ Important — please read this before relying on it:**

This uses web scraping, not an official search API, because DuckDuckGo doesn't
offer a free key-based search API. That means:

- **It's against DuckDuckGo's Terms of Service** for automated/non-personal use.
  This was a deliberate trade-off to avoid requiring you to sign up for yet
  another API key — but it's worth knowing the legal status is a gray area.
- **It will fail unpredictably.** DuckDuckGo rate-limits and blocks scrapers,
  sometimes within a handful of requests, sometimes after many. There's no
  published threshold.
- **It can break entirely** if DuckDuckGo changes their HTML page structure,
  since the scraper depends on specific CSS class names that aren't a stable
  public contract.
- **The app handles failure gracefully** — if the search fails for any reason,
  it just answers from the model's own knowledge instead of crashing.

### If you want reliable current info instead

Swap `web_search_duckduckgo()` for a real free-tier search API — the best option
is **Tavily** (1,000 free searches/month, no credit card, built specifically for
feeding LLMs). It takes about 2 minutes to get a key at tavily.com, and the only
code change needed is inside that one function — everything else (the keyword
trigger, the context injection, the UI) stays the same.

## Known Limitations (Honest Scope)

- Multi-document context is **combined text + image injection**, not real RAG —
  fine for a handful of files, not built for large document libraries.
- PPTX/XLSX visual rendering depends on LibreOffice being installed and on your
  system PATH — without it, you still get full text extraction, just no images.
- Ollama vision detection is a **name-based heuristic** (checks if the model name
  contains "llava", "vision", etc.) — if you rename a model unusually, the
  detection might miss it.
- API keys are stored only in Streamlit's session state (in-memory, per browser
  tab) — they are *not* saved anywhere, so you'll need to re-enter them each
  time you restart the app. (Chat history persists; API keys do not.)
- The web search feature is best-effort scraping — see the dedicated section above.
- Each document's extracted text is capped at ~15,000 characters to keep
  requests manageable; very long documents get truncated.

## Extending This Project

Ideas if you want to build on this for your portfolio:
- Add streaming responses (token-by-token) instead of waiting for the full reply
- Add a "system prompt" field in the sidebar to customize the assistant's persona
- Swap the text-injection approach for real chunking + embeddings (true RAG)
- Add full-text search across saved chat history, not just titles
- Support audio/video file uploads with transcription
- Add per-document toggles so you can choose which active documents apply to
  a given question, instead of always combining all of them
