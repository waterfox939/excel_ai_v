# Excel AI Agent

An AI assistant (powered by Claude, Anthropic's API) that works alongside you
inside Excel — not a chatbot you upload files to. It lives in a task pane
next to your spreadsheet, sees whatever you've selected, and can read or
write data on **any sheet in the workbook**, live, while you're working.
Originally built around order records, shipping, UPS invoices, and payment
summaries, but general-purpose for any Excel workflow.

You can also just talk to it from a terminal (the original CLI), or from a
floating popup window (an earlier, now-secondary interface) — see below.

## How it runs on your device

Everything runs **locally on your own machine** — there's no cloud service,
no shared server, and nothing of yours is uploaded anywhere except the
messages you send to Anthropic's Claude API directly (using your own API
key, billed to your own account).

Concretely, two things run side by side on your computer:

1. **A small local backend** (Python) that talks to the Claude API on your
   behalf. It runs on `localhost` and is never reachable from outside your
   machine.
2. **A frontend** that you actually interact with — the Excel task pane,
   the CLI, or the popup window — which talks to that local backend over
   `localhost`.

Your spreadsheet data, your questions, and any files/images you attach are
sent only to Anthropic's API (to generate a response) and never touch any
server operated by this project.

## Supported devices

| Interface | Platform | Requirements |
|---|---|---|
| **Excel Add-in** (primary) | macOS (packaged standalone install); dev setup also works cross-platform | Excel desktop, Node.js (dev setup only) |
| **CLI** | macOS / Windows / Linux | Python 3 |
| **Desktop popup** (secondary) | macOS (built/tested here); Electron itself is cross-platform | Node.js |

The **standalone packaged install** (see "Standalone install" below — the
easiest way to run this without setting up a dev environment) currently
only has a build for **macOS**. Windows/Linux packaging hasn't been built
yet, though nothing about the underlying design is Mac-specific — see
Project status.

## Setup

1. Create a virtual environment and install dependencies:

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

2. Create a `.env` file in this directory with your API key:

   ```
   ANTHROPIC_API_KEY=your_key_here
   ```

3. Run the CLI:

   ```bash
   python main.py
   ```

## Excel Add-in (primary way to use this)

A task pane that runs *inside* Excel itself — no file picking, no separate
window. It reads whatever range you currently have selected in your open
workbook, can check and edit **any other sheet by name** (not just the
active one), and supports attaching images/PDFs/CSV/XLSX files to a
message. Built with Office.js.

Claude can call four tools, executed live against your open workbook via
`Excel.run` — never against a file on disk:

- `list_sheets`, `get_used_range`, `read_range` — run immediately, read-only.
- `write_range` — always shows an Apply/Reject card first. Nothing is
  written to the workbook until you click Apply.

Requires [Node.js](https://nodejs.org) (LTS) and Excel desktop installed.

```bash
# terminal 1 — Python backend
cd excel-ai-agent
source .venv/bin/activate
uvicorn server:app --reload --port 8765

# terminal 2 — Add-in dev server (HTTPS, proxies /api to the backend above)
cd excel-ai-agent/addin
npm install
npm run dev-server
```

First-time sideload on Mac (one-time setup so Excel knows the add-in exists):

1. Finder → `Cmd+Shift+G` → go to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (create the folder if it isn't there).
2. Copy `addin/manifest.xml` into it.
3. Open Excel → any workbook → **Home tab → Add-ins → My Add-ins** → select "Excel AI Agent".

The task pane opens on the right side of the window. Select some cells,
type a message — the current selection is sent along with your message
automatically.

### Standalone install (no dev setup, own API key) — Mac only for now

For running this on another machine without cloning the repo or running
two dev servers by hand. Everything stays local — there's no shared
backend, no hosting, and no API key of mine involved; each install uses
its own key.

**Build it once** (from a machine with this repo + Node + the Python venv):

```bash
source .venv/bin/activate
python3 packaging/build.py
```

This produces a folder at `packaging/dist/excel-ai-agent/` — copy that
whole folder to wherever you want to run it (including another Mac).

**Run it** (first time on a given machine):

```bash
cd packaging/dist/excel-ai-agent
./excel-ai-agent
```

On first run this will, in order: generate and trust a local HTTPS
certificate (no admin password needed — it only touches your personal
login keychain, not the system one), ask you to paste your own Anthropic
API key (input is hidden, saved to `~/.excel-ai-agent/config.json`, never
inside the app folder), and copy the manifest into Excel's `wef` folder
automatically. Leave the terminal window open — that's the local server;
closing it stops the Add-in from working. Open Excel afterward and the
Add-in will be there under **Home → Add-ins → My Add-ins**, same as the
dev setup above.

Re-running `./excel-ai-agent` later (e.g. after rebuilding a new version)
reuses the saved cert and API key — you won't be asked again.

## Desktop popup UI (deprioritized)

An earlier attempt: press `Cmd/Ctrl+Shift+E` from anywhere for a floating
popup where you pick an Excel file and chat about it. Superseded by the
Excel Add-in above (works alongside you *in* Excel instead of a separate
window with manual file picking), but still functional — see `desktop/`.

```bash
cd desktop
npm install
npm run dev
```

## Project status

Phase 1 (CLI) and Phase 2 (`excel_tools.py` — used only by the CLI's
file-based workflows, not the Add-in) complete. Two frontends exist on top
of the same Python backend: the Excel Add-in (`addin/`, primary — now with
live multi-sheet read/write tool-calling and file/image attachments) and
the earlier Electron popup (`desktop/`, deprioritized). The Add-in can also
be built into a standalone local install (`packaging/`, Mac only for now —
see above) so it can run on another machine without a dev setup. Decimal
reconciliation and cycle templates are still ahead. See the project plan
for details.
