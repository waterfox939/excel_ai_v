# Excel AI Agent

Claude Code, but for Excel.

An AI assistant (powered by Claude, Anthropic's API) that works alongside you
inside Excel — not a chatbot you upload files to. It lives in a task pane
next to your spreadsheet, sees whatever you've selected, and can read or
write data on **any sheet in the workbook**, live, while you're working.
Originally built around order records, shipping, UPS invoices, and payment
summaries, but general-purpose for any Excel workflow.

You can also just talk to it from a terminal (the original CLI), or from a
floating popup window (an earlier, now-secondary interface) — see below.

> **Before you point this at real work, read
> [Your data](#️-your-data--please-read-before-using-this-on-real-work).**
> Your spreadsheet contents are sent to Anthropic'"'"'s API, and Claude can read
> every sheet in an open workbook. Use it at your own discretion.

---

## Install (Windows or macOS)

**You do not need to install Python, Node.js, or anything else.** Download
one file, unzip it, double-click it.

1. Go to the [**Releases page**](../../releases) and download the zip for
   your machine:
   - Windows → `excel-ai-agent-windows-x64.zip`
   - macOS → `excel-ai-agent-macos-arm64.zip`
2. **Windows only — unblock it first.** Right-click the downloaded zip →
   **Properties** → tick **Unblock** at the bottom → OK. *Then* extract it.
   Windows tags anything downloaded from the internet, and doing this before
   extracting saves you from a SmartScreen warning on every file inside.
3. Extract the zip anywhere you like (Documents is fine — but see the note
   below about not moving it afterwards).
4. Run it:
   - Windows → double-click **`excel-ai-agent.exe`**
   - macOS → open Terminal in that folder and run `./excel-ai-agent`
5. The first run asks you to **paste your own Anthropic API key** (get one at
   [console.anthropic.com](https://console.anthropic.com/settings/keys)).
   Input is hidden. It's saved to `~/.excel-ai-agent/config.json`, outside
   the app folder, and you won't be asked again.
6. **Leave that window open.** It *is* the app — closing it stops the add-in
   from working.
7. Open Excel. The add-in is under **Home → Add-ins → My Add-ins → Excel AI
   Agent**. Select some cells and start typing.

That's the whole install. Everything after step 1 happens offline on your
own machine.

### What the first run does automatically

- Generates a self-signed HTTPS certificate for `localhost` and trusts it for
  **your user account only** — no admin password, no system-wide changes.
  Office refuses to load a task pane over plain HTTP, so this is required.
- Registers the add-in with Excel (a registry entry on Windows, a file copy
  on macOS).
- Starts the local server on `https://localhost:8765`.

All three are idempotent — re-running after an update is safe and won't
re-prompt.

### Things that will look alarming but are normal

- **"Windows protected your PC" (SmartScreen).** The `.exe` isn't code-signed
  — signing certificates cost a few hundred dollars a year. Click **More
  info → Run anyway**, or do the Unblock step above to avoid it entirely.
- **A one-time certificate confirmation dialog.** Windows asks before adding
  any certificate to your trust store. Say yes — if you decline, the task
  pane won't load, though re-running the app will offer it again.
- **Antivirus grumbling.** Apps bundled this way (PyInstaller) get flagged as
  false positives fairly often.
- **Don't move the folder after the first run.** Excel is pointed at the copy
  of the manifest made during setup. If you do move it, just run the app once
  more from its new home.

### What "no downloads" does and doesn't mean

You need no *developer tools* — no Python, no Node, no package managers, no
build step. You do still need an internet connection while using it: the task
pane loads Office.js from Microsoft's CDN (required for all Office add-ins),
and the whole point is talking to Claude's API.

---

## ⚠️ Your data — please read before using this on real work

**Use this at your own discretion, especially with sensitive or confidential
files.** This is a personal project provided as-is, with no warranty (see
[LICENSE](LICENSE)). You are responsible for deciding what is appropriate to
put through it.

Specifically, be aware that:

- **What you send leaves your machine.** Your selected cells, anything Claude
  reads via `read_range`, your questions, and every file you attach are sent
  to **Anthropic's API** over the internet in order to generate a reply. They
  do not touch any server run by this project — but "runs locally" means the
  *app* is local, not that your data stays on your computer.
- **Claude can see more than your selection.** The `list_sheets`,
  `get_used_range`, and `read_range` tools let it read **any sheet in the open
  workbook** on its own initiative, without asking first. If a workbook has a
  tab you would not want sent to an API, don't have it open.
- **Writes modify your real workbook.** `write_range` always asks first via an
  Apply/Reject card, but once you click Apply it overwrites those cells
  immediately. Your only safety net is Excel's own undo. **Back up anything
  you can't afford to lose**, and prefer working on a copy.
- **Your API key is stored in plain text** at `~/.excel-ai-agent/config.json`
  (Windows: `C:\Users\<you>\.excel-ai-agent\config.json`). The app restricts
  the file to your user account, but it is not encrypted. Anyone with access
  to your account can read it. Revoke it from the
  [Anthropic console](https://console.anthropic.com/settings/keys) if you
  suspect it has leaked.
- **Check your obligations.** If you work with regulated, client-confidential,
  personal, or otherwise restricted data, confirm your organisation's policy
  and Anthropic's terms before pointing this at it. Some workplaces prohibit
  sending business data to third-party AI services entirely.
- **AI output can be wrong.** Claude misreads spreadsheets, miscounts rows,
  and does arithmetic incorrectly. Check anything that matters before relying
  on it — particularly numbers it writes back into your workbook.

If in doubt, try it on a dummy copy of your workbook first.

---

## How it runs on your device

Everything runs **locally on your own machine** — there's no cloud service,
no shared server, and nothing of yours is uploaded anywhere except the
messages you send to Anthropic's Claude API directly (using your own API
key, billed to your own account).

Concretely, two things run side by side on your computer:

1. **A small local backend** (Python, compiled into the download) that talks
   to the Claude API on your behalf. It binds to `127.0.0.1` and is never
   reachable from outside your machine.
2. **A frontend** that you actually interact with — the Excel task pane,
   the CLI, or the popup window — which talks to that local backend over
   `localhost`.

Your spreadsheet data, your questions, and any files/images you attach are
sent only to Anthropic's API (to generate a response) and never touch any
server operated by this project.

## Supported devices

| Interface | Platform | Requirements |
|---|---|---|
| **Excel Add-in** (primary) | Windows + macOS | Excel desktop. Nothing else. |
| **CLI** | Windows / macOS / Linux | Python 3 (run from source) |
| **Desktop popup** (secondary) | macOS built/tested; Electron is cross-platform | Node.js (run from source) |

Linux has no packaged build. The CLI works there from source.

## What Claude can do in your workbook

Four tools, executed live against your open workbook via `Excel.run` — never
against a file on disk:

- `list_sheets`, `get_used_range`, `read_range` — run immediately, read-only.
- `write_range` — **always** shows an Apply/Reject card first. Nothing is
  written to the workbook until you click Apply.

The task pane reads whatever range you currently have selected, can inspect
**any other sheet by name** (not just the active one), and accepts attached
images, PDFs, CSVs, and XLSX files.

### In the task pane

- **Every proposed write is shown as a diff first.** Before you approve
  anything, the card lists each cell that changes with its current value
  beside the new one, so you can see exactly what Apply will do. Cells whose
  value wouldn't change are counted separately rather than cluttering the
  list.
- **Formatted replies** — headings, lists, bold, and code blocks render
  properly instead of showing raw markdown.
- **Stop button** to interrupt a reply mid-stream, and **New** to clear the
  conversation.
- **Live tool activity**, so you can see when Claude is reading a range
  rather than staring at a blank pane.
- **Follows your system light/dark setting**, and Shift+Enter inserts a
  newline while Enter sends.

---

## Versions

**v2 (current) — standalone.** Download a zip from Releases, double-click it,
paste your API key. No Python, no Node, no build step. Windows and macOS.

**v1 — developer setup only.** Preserved at the [`AI`](../../tree/AI) tag. It
needed a cloned repo, a Python virtualenv, Node and npm, and two terminals
running side by side (a webpack dev server plus the backend). Its standalone
build was macOS-only in practice; the Windows path was written but had never
been built or run.

| | v1 | v2 |
|---|---|---|
| To install | clone, venv, `npm install` | download a zip |
| Node.js needed | yes | no |
| Frontend build step | webpack | none |
| Download size | roughly 3× larger | ~47 MB |
| Windows build | never produced | built by CI |
| Proposed writes shown as | raw JSON | per-cell before/after diff |
| Replies rendered as | plain text | formatted markdown |

v1 is kept only so the old setup stays reachable. v2 replaces it entirely —
there is nothing v1 does that v2 doesn't.

## Running from source (development)

Only needed if you want to change the code. Requires Python 3.10+.

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Create a `.env` file with `ANTHROPIC_API_KEY=your_key_here`, then:

```bash
# The Add-in — serves the task pane and the API on https://localhost:8765
python server.py

# Or the CLI
python main.py
```

**There is no frontend build step and no Node.js.** `addin/src/taskpane/` is
plain browser JavaScript with no imports, so `server.py` serves it directly. Dev
and packaged runs serve identical files from an identical URL, which is why a
single `addin/manifest.xml` covers both.

The one exception is the task-pane test suite, which needs Node to run — it
is never needed to build or use the app, and CI runs it on a runner that
already has Node:

```bash
node tests/taskpane.test.js
```

It covers the markdown renderer (including that model output can't inject
markup), the spreadsheet column/cell-reference arithmetic behind the write
preview, and that the HTML, CSS, and JS still agree on IDs, classes, and
theme variables.

To sideload the manifest by hand while developing, run the packaged entry
point once (`python packaging/entry.py`), or place `addin/manifest.xml`
yourself:

- **macOS** — copy it to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
- **Windows** — add a string value under
  `HKCU\Software\Microsoft\Office\16.0\Wef\Developer` whose *name* is the
  manifest's `<Id>` GUID and whose *value* is the full path to the manifest.

Then: Excel → **Home → Add-ins → My Add-ins → Excel AI Agent**.

### Building a release

You don't need a Windows machine. Push a tag and GitHub Actions builds both
platforms and publishes them:

```bash
git tag v1.0.0 && git push origin v1.0.0
```

`.github/workflows/release.yml` runs `packaging/build.py` on hosted Windows
and macOS runners (PyInstaller can't cross-compile, which is the entire
reason CI does this), smoke-tests the result, and attaches the zips to a
GitHub Release. A push to **any** branch runs the tests and builds both
platforms without publishing anything, so you can confirm a build works
before you tag it — only a `v*` tag creates a Release.

To build locally instead — on the OS you're targeting:

```bash
pip install -r requirements-build.txt
python packaging/build.py
```

## Desktop popup UI (deprioritized)

An earlier attempt: press `Cmd/Ctrl+Shift+E` from anywhere for a floating
popup where you pick an Excel file and chat about it. Superseded by the
Excel Add-in above (works alongside you *in* Excel instead of a separate
window with manual file picking), but still functional — see `desktop/`.
This one does still need Node.js.

```bash
cd desktop
npm install
npm run dev
```

## Project status

Phase 1 (CLI) and Phase 2 (`excel_tools.py` — used only by the CLI's
file-based workflows, not the Add-in) complete. Two frontends exist on top
of the same Python backend: the Excel Add-in (`addin/`, primary — live
multi-sheet read/write tool-calling and file/image attachments) and the
earlier Electron popup (`desktop/`, deprioritized).

The Add-in ships as a standalone download for Windows and macOS built by CI
(`packaging/`, `.github/workflows/release.yml`), so anyone can run it without
a development environment. Decimal reconciliation and cycle templates are
still ahead. See the project plan for details.

**Not yet verified firsthand:** no packaged build has been run end-to-end on
a real Windows machine yet. CI proves the build succeeds and that the task
pane files are bundled; what it cannot exercise is the certificate trust
prompt and the registry sideload, which only run on a real Windows desktop. If something breaks, the likeliest spots are the certificate trust step
and the registry sideload — please open an issue with what you see.
