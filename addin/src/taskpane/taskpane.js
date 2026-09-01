/* global Office, Excel */

const selectionInfo = document.getElementById("selection-info");
const refreshSelectionBtn = document.getElementById("refresh-selection-btn");
const replyPane = document.getElementById("reply-pane");
const chatInput = document.getElementById("chat-input");
const attachBtn = document.getElementById("attach-btn");
const fileInput = document.getElementById("file-input");
const pendingAttachmentsEl = document.getElementById("pending-attachments");

let lastSelection = null;
// Full conversation history — the API is stateless, so we resend it every
// turn (same pattern as the Electron popup's renderer.js).
let conversation = [];
let pendingAttachments = []; // [{name, media_type, data}]

const MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024;
const EXT_MEDIA_TYPES = {
  csv: "text/csv",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pdf: "application/pdf",
};

Office.onReady(() => {
  chatInput.disabled = false;
  refreshSelection();
});

// ---------- Office.js helpers ----------

async function getSelectionContext() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();
    sheet.load("name");
    range.load("address, values, rowCount, columnCount");
    await context.sync();
    return { sheetName: sheet.name, address: range.address, values: range.values };
  });
}

async function refreshSelection() {
  try {
    lastSelection = await getSelectionContext();
    selectionInfo.textContent = `${lastSelection.sheetName}!${lastSelection.address} (${lastSelection.values.length}x${lastSelection.values[0]?.length ?? 0})`;
  } catch (err) {
    selectionInfo.textContent = `Could not read selection: ${err.message}`;
    lastSelection = null;
  }
}

refreshSelectionBtn.addEventListener("click", refreshSelection);

async function listSheetsTool() {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    return sheets.items.map((s) => s.name);
  });
}

async function getUsedRangeTool(sheetName) {
  return Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(sheetName).getUsedRange();
    range.load("address, rowCount, columnCount");
    await context.sync();
    return { address: range.address, rowCount: range.rowCount, columnCount: range.columnCount };
  });
}

async function readRangeTool(sheetName, address) {
  return Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(address);
    range.load("values");
    await context.sync();
    return range.values;
  });
}

async function writeRangeTool(sheetName, address, values) {
  return Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(address);
    range.values = values;
    await context.sync();
    return { applied: true };
  });
}

// ---------- Chat transcript rendering ----------

function appendTurn(role, text) {
  const div = document.createElement("div");
  div.className = `turn turn-${role}`;
  div.textContent = text;
  replyPane.appendChild(div);
  replyPane.scrollTop = replyPane.scrollHeight;
  return div;
}

function appendActivityLine(text) {
  const div = document.createElement("div");
  div.className = "activity-line";
  div.textContent = text;
  replyPane.appendChild(div);
  replyPane.scrollTop = replyPane.scrollHeight;
  return div;
}

function describeTool(name, input) {
  switch (name) {
    case "list_sheets":
      return "Listing sheets…";
    case "get_used_range":
      return `Checking extent of ${input.sheet_name}…`;
    case "read_range":
      return `Reading ${input.sheet_name}!${input.address}…`;
    default:
      return `Running ${name}…`;
  }
}

// Shows an Apply/Reject card for a write_range tool call. Resolves with the
// tool_result content once the user decides — nothing touches the workbook
// until Apply is clicked.
function showWriteConfirmCard(input) {
  return new Promise((resolve) => {
    const card = document.createElement("div");
    card.className = "write-confirm-card";
    const valuesPreview = JSON.stringify(input.values);
    card.innerHTML = `
      <div class="write-summary">Claude wants to write to <code>${input.sheet_name}!${input.address}</code>: <code>${valuesPreview}</code></div>
      <div class="write-actions">
        <button type="button" class="apply-btn">Apply</button>
        <button type="button" class="reject-btn">Reject</button>
      </div>
    `;
    replyPane.appendChild(card);
    replyPane.scrollTop = replyPane.scrollHeight;

    const applyBtn = card.querySelector(".apply-btn");
    const rejectBtn = card.querySelector(".reject-btn");
    const finish = (resultText) => {
      applyBtn.disabled = true;
      rejectBtn.disabled = true;
      const resultEl = document.createElement("div");
      resultEl.className = "write-result";
      resultEl.textContent = resultText;
      card.appendChild(resultEl);
      resolve(resultText);
    };

    applyBtn.addEventListener("click", async () => {
      try {
        await writeRangeTool(input.sheet_name, input.address, input.values);
        finish("Applied.");
      } catch (err) {
        finish(`Failed to apply: ${err.message}`);
      }
    });
    rejectBtn.addEventListener("click", () => finish("User declined this write."));
  });
}

async function executeExcelTool(toolUse) {
  const { name, input } = toolUse;
  if (name !== "write_range") {
    appendActivityLine(describeTool(name, input));
  }
  try {
    switch (name) {
      case "list_sheets":
        return JSON.stringify(await listSheetsTool());
      case "get_used_range":
        return JSON.stringify(await getUsedRangeTool(input.sheet_name));
      case "read_range":
        return JSON.stringify(await readRangeTool(input.sheet_name, input.address));
      case "write_range":
        return await showWriteConfirmCard(input);
      default:
        return `Unknown tool: ${name}`;
    }
  } catch (err) {
    return `Error: ${err.message}`;
  }
}

// ---------- Attachments ----------

function mediaTypeFor(file) {
  if (file.type) return file.type;
  const ext = file.name.split(".").pop().toLowerCase();
  return EXT_MEDIA_TYPES[ext] || "application/octet-stream";
}

function renderAttachmentChips() {
  pendingAttachmentsEl.innerHTML = "";
  pendingAttachments.forEach((att, i) => {
    const chip = document.createElement("div");
    chip.className = "attachment-chip";
    const nameSpan = document.createElement("span");
    nameSpan.textContent = att.name;
    const removeSpan = document.createElement("span");
    removeSpan.className = "remove-btn";
    removeSpan.textContent = "×";
    removeSpan.addEventListener("click", () => {
      pendingAttachments.splice(i, 1);
      renderAttachmentChips();
    });
    chip.appendChild(nameSpan);
    chip.appendChild(removeSpan);
    pendingAttachmentsEl.appendChild(chip);
  });
}

attachBtn.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", () => {
  for (const file of fileInput.files) {
    if (file.size > MAX_ATTACHMENT_BYTES) {
      appendActivityLine(`${file.name} is too large (max 10MB), skipped.`);
      continue;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const base64 = reader.result.split(",")[1];
      pendingAttachments.push({ name: file.name, media_type: mediaTypeFor(file), data: base64 });
      renderAttachmentChips();
    };
    reader.readAsDataURL(file);
  }
  fileInput.value = "";
});

// ---------- Chat loop ----------

async function runTurn(attachmentsForThisCall) {
  const response = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      messages: conversation,
      selection: lastSelection,
      attachments: attachmentsForThisCall || null,
    }),
  });

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    appendTurn("assistant", `[Backend error ${response.status}: ${body || response.statusText}]`);
    return;
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let assistantText = "";
  let replyEl = null;
  const toolUses = [];

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n\n");
      buffer = lines.pop();
      for (const line of lines) {
        if (!line.startsWith("data: ")) continue;
        const event = JSON.parse(line.slice("data: ".length));
        if (event.type === "text") {
          if (!replyEl) replyEl = appendTurn("assistant", "");
          assistantText += event.text;
          replyEl.textContent = assistantText;
          replyPane.scrollTop = replyPane.scrollHeight;
        } else if (event.type === "tool_use") {
          toolUses.push(event);
        }
      }
    }
  } catch (err) {
    appendTurn("assistant", `[Stream interrupted: ${err.message}]`);
    return;
  }

  if (toolUses.length === 0) {
    conversation.push({ role: "assistant", content: assistantText });
    return;
  }

  conversation.push({
    role: "assistant",
    content: [
      ...(assistantText ? [{ type: "text", text: assistantText }] : []),
      ...toolUses.map((t) => ({ type: "tool_use", id: t.id, name: t.name, input: t.input })),
    ],
  });

  const results = [];
  for (const t of toolUses) {
    const resultContent = await executeExcelTool(t);
    results.push({ type: "tool_result", tool_use_id: t.id, content: resultContent });
  }
  conversation.push({ role: "user", content: results });

  await runTurn(); // continue the loop — no attachments on tool-result follow-ups
}

async function sendChat(message) {
  const attachmentsForThisCall = pendingAttachments.length ? pendingAttachments : null;
  const attachmentNote = pendingAttachments.length
    ? `\n📎 ${pendingAttachments.map((a) => a.name).join(", ")}`
    : "";
  conversation.push({ role: "user", content: message });
  appendTurn("user", message + attachmentNote);
  pendingAttachments = [];
  renderAttachmentChips();

  await refreshSelection();
  await runTurn(attachmentsForThisCall);
}

chatInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter" && chatInput.value.trim()) {
    const message = chatInput.value.trim();
    chatInput.value = "";
    sendChat(message);
  }
});
