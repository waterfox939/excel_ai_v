"""Local HTTP backend for the Excel Add-in task pane. Thin wrapper around agent.py's client."""
import base64
import csv
import io
import json
import sys
from pathlib import Path

import openpyxl
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from agent import client
from config import MODEL

PORT = 8765
# PyInstaller sets sys._MEIPASS to wherever it actually placed bundled
# `datas` — for --onedir that's an _internal/ folder next to the exe, not
# the exe's own directory (matches packaging/entry.py's base_dir()).
_BASE_DIR = Path(sys._MEIPASS) if getattr(sys, "frozen", False) else Path(__file__).parent
# Served straight from source — there is no frontend build step. taskpane.js
# is plain browser JS with no imports, so nothing needs bundling.
TASKPANE_DIR = _BASE_DIR / "addin" / "src" / "taskpane"
ASSETS_DIR = _BASE_DIR / "addin" / "assets"

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"])

MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024
SUPPORTED_IMAGE_TYPES = {"image/jpeg", "image/png", "image/gif", "image/webp"}
MAX_XLSX_ROWS = 500

EXCEL_TOOLS = [
    {
        "name": "list_sheets",
        "description": "List all worksheet names in the open workbook.",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_used_range",
        "description": "Get the used range's address and dimensions for a sheet.",
        "input_schema": {
            "type": "object",
            "properties": {"sheet_name": {"type": "string"}},
            "required": ["sheet_name"],
        },
    },
    {
        "name": "read_range",
        "description": "Read cell values from a sheet and range, e.g. sheet_name='Sheet2', address='A1:D10'.",
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string"},
                "address": {"type": "string"},
            },
            "required": ["sheet_name", "address"],
        },
    },
    {
        "name": "write_range",
        "description": (
            "Write values to a sheet and range. Requires the user to confirm "
            "before it takes effect."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string"},
                "address": {"type": "string"},
                "values": {"type": "array", "items": {"type": "array"}},
            },
            "required": ["sheet_name", "address", "values"],
        },
    },
]


class Selection(BaseModel):
    sheetName: str
    address: str
    values: list[list]


class Attachment(BaseModel):
    name: str
    media_type: str
    data: str  # base64, no data-URL prefix


class ChatRequest(BaseModel):
    messages: list[dict]
    file_path: str | None = None
    selection: Selection | None = None
    attachments: list[Attachment] | None = None


def _xlsx_to_csv(raw: bytes) -> str:
    """Render the first sheet of an .xlsx attachment as CSV text for Claude.

    Uses openpyxl directly rather than pandas — pandas was a ~50MB
    dependency pulled in for this single call, and it dominated the size of
    the packaged Windows/macOS download. read_only streams rather than
    building the whole sheet in memory; data_only yields the cached results
    of formulas instead of the formula source.
    """
    workbook = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    try:
        sheet = workbook.worksheets[0]
        out = io.StringIO()
        writer = csv.writer(out, lineterminator="\n")
        # Row 0 is the header, so MAX_XLSX_ROWS data rows means +1 total.
        for i, row in enumerate(sheet.iter_rows(values_only=True)):
            if i > MAX_XLSX_ROWS:
                break
            writer.writerow(["" if value is None else value for value in row])
        return out.getvalue()
    finally:
        workbook.close()


def _build_content_blocks(text: str, attachments: list[Attachment] | None) -> str | list[dict]:
    """Assemble a user message's content, folding in any attachments.

    Returns a plain string in the common no-attachment case, or a list of
    content blocks (text + image/document/inlined-text) when there are
    attachments. Unsupported or oversized files degrade to an inline note
    rather than erroring the whole request.
    """
    if not attachments:
        return text
    blocks: list[dict] = [{"type": "text", "text": text}]
    for att in attachments:
        try:
            raw = base64.b64decode(att.data)
        except Exception:
            blocks.append({"type": "text", "text": f"[Attachment {att.name}: could not decode, skipped]"})
            continue
        if len(raw) > MAX_ATTACHMENT_BYTES:
            blocks.append({"type": "text", "text": f"[Attachment {att.name} exceeds size limit, skipped]"})
            continue
        if att.media_type in SUPPORTED_IMAGE_TYPES:
            blocks.append(
                {"type": "image", "source": {"type": "base64", "media_type": att.media_type, "data": att.data}}
            )
        elif att.media_type == "application/pdf":
            blocks.append(
                {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": att.data}}
            )
        elif att.media_type == "text/csv" or att.name.lower().endswith(".csv"):
            blocks.append({"type": "text", "text": f"[Attached file: {att.name}]\n{raw.decode('utf-8', errors='replace')}"})
        elif att.name.lower().endswith(".xlsx"):
            try:
                csv_text = _xlsx_to_csv(raw)
                blocks.append({"type": "text", "text": f"[Attached file: {att.name}, sheet 1]\n{csv_text}"})
            except Exception as exc:
                blocks.append({"type": "text", "text": f"[Attachment {att.name}: could not parse xlsx ({exc}), skipped]"})
        else:
            blocks.append({"type": "text", "text": f"[Attachment {att.name}: unsupported type {att.media_type}, skipped]"})
    return blocks


def _with_selection_context(messages: list[dict], selection: Selection | None) -> list[dict]:
    """Prepend the live Excel selection to the latest user turn, if provided.

    Kept as a text prefix rather than a separate message so the conversation
    still alternates user/assistant as the API requires. Handles content
    being either a plain string or a list of content blocks (attachments).
    """
    if selection is None or not messages:
        return messages
    context_note = (
        f"[Current selection: {selection.sheetName}!{selection.address}, "
        f"values={selection.values}]\n\n"
    )
    last = messages[-1]
    content = last["content"]
    if isinstance(content, str):
        new_content = context_note + content
    else:
        new_content = [{"type": "text", "text": context_note}, *content]
    return [*messages[:-1], {**last, "content": new_content}]


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/api/chat")
def chat(request: ChatRequest) -> StreamingResponse:
    messages = list(request.messages)
    if messages and messages[-1]["role"] == "user":
        messages[-1] = {
            **messages[-1],
            "content": _build_content_blocks(messages[-1]["content"], request.attachments),
        }
    messages = _with_selection_context(messages, request.selection)

    def event_stream():
        with client.messages.stream(
            model=MODEL,
            max_tokens=1024,
            system=(
                "You are an assistant embedded in an Excel task pane, working "
                "alongside the user on their live, open workbook. Use the "
                "list_sheets/get_used_range/read_range tools to check the "
                "actual workbook when you need information from it — don't "
                "guess. If the user's message includes attached file "
                "contents, treat that attachment as the primary source for "
                "anything it directly answers; only fall back to reading the "
                "live workbook if the attachment doesn't cover the request. "
                "write_range requires the user to click Apply before it "
                "takes effect — mention that when proposing a write."
            ),
            messages=messages,
            tools=EXCEL_TOOLS,
        ) as stream:
            for event in stream:
                if event.type == "content_block_delta" and event.delta.type == "text_delta":
                    yield f'data: {json.dumps({"type": "text", "text": event.delta.text})}\n\n'
            final = stream.get_final_message()
            for block in final.content:
                if block.type == "tool_use":
                    yield f'data: {json.dumps({"type": "tool_use", "id": block.id, "name": block.name, "input": block.input})}\n\n'
        yield f'data: {json.dumps({"type": "done"})}\n\n'

    return StreamingResponse(event_stream(), media_type="text/event-stream")


# Serve the Add-in frontend. Must be registered AFTER all /api/* routes
# above — Starlette matches routes in registration order, and "/" below is a
# catch-all mount that would otherwise shadow them. Likewise /assets must
# precede "/", since the icons live outside the taskpane directory.
#
# This is the whole frontend "build": there is no bundler and no dev server.
# Dev and packaged runs serve the identical files from the identical URL
# (https://localhost:8765), which is why one manifest.xml covers both.
if ASSETS_DIR.is_dir():
    app.mount("/assets", StaticFiles(directory=str(ASSETS_DIR)), name="assets")
if TASKPANE_DIR.is_dir():
    app.mount("/", StaticFiles(directory=str(TASKPANE_DIR), html=True), name="taskpane")


if __name__ == "__main__":
    import uvicorn

    from cert_setup import ensure_cert

    cert = ensure_cert()
    ssl_kwargs = {"ssl_keyfile": str(cert.key_path), "ssl_certfile": str(cert.cert_path)} if cert else {}
    uvicorn.run(app, host="127.0.0.1", port=PORT, **ssl_kwargs)
