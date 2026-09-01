const recentFilesEl = document.getElementById('recent-files');
const browseBtn = document.getElementById('browse-btn');
const activeFileEl = document.getElementById('active-file');
const replyPane = document.getElementById('reply-pane');
const chatInput = document.getElementById('chat-input');

let activeFilePath = null;

function setActiveFile(filePath) {
  activeFilePath = filePath;
  activeFileEl.textContent = filePath ? filePath.split(/[\\/]/).pop() : 'No file selected';
  document.querySelectorAll('.recent-chip').forEach((chip) => {
    chip.classList.toggle('active', chip.dataset.path === filePath);
  });
}

async function renderRecentFiles() {
  const files = await window.api.getRecentFiles();
  recentFilesEl.innerHTML = '';
  for (const f of files) {
    const chip = document.createElement('div');
    chip.className = 'recent-chip';
    chip.textContent = f.name;
    chip.dataset.path = f.path;
    chip.addEventListener('click', () => setActiveFile(f.path));
    recentFilesEl.appendChild(chip);
  }
  if (files.length > 0 && !activeFilePath) {
    setActiveFile(files[0].path);
  }
}

browseBtn.addEventListener('click', async () => {
  const filePath = await window.api.pickFile();
  if (filePath) {
    await renderRecentFiles();
    setActiveFile(filePath);
  }
});

// Full conversation history — the API is stateless, so we resend it every
// turn. Client-owned rather than server-tracked: simplest way to support
// multiple independent chats without session/connection bookkeeping.
let conversation = [];

function appendTurn(role, text) {
  const div = document.createElement('div');
  div.className = `turn turn-${role}`;
  div.textContent = text;
  replyPane.appendChild(div);
  replyPane.scrollTop = replyPane.scrollHeight;
  return div;
}

async function sendChat(message) {
  conversation.push({ role: 'user', content: message });
  appendTurn('user', message);
  const replyEl = appendTurn('assistant', '');

  let response;
  try {
    response = await fetch(`http://127.0.0.1:${window.api.chatPort}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ messages: conversation, file_path: activeFilePath }),
    });
  } catch (err) {
    replyEl.textContent = `[Could not reach the backend: ${err.message}]`;
    conversation.pop();
    return;
  }

  if (!response.ok) {
    const body = await response.text().catch(() => '');
    replyEl.textContent = `[Backend error ${response.status}: ${body || response.statusText}]`;
    conversation.pop();
    return;
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let fullReply = '';
  let sawToolUse = false;

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });

      const lines = buffer.split('\n\n');
      buffer = lines.pop(); // last (possibly incomplete) chunk stays in the buffer
      for (const line of lines) {
        if (!line.startsWith('data: ')) continue;
        // Wire format is JSON-per-event (shared with the Excel Add-in), not
        // raw text — see server.py's /api/chat. This window has no way to
        // execute Excel tool calls (no Office.js context), so it can only
        // render text events and flag when a tool call was requested.
        const event = JSON.parse(line.slice('data: '.length));
        if (event.type === 'text') {
          fullReply += event.text;
          replyEl.textContent += event.text;
          replyPane.scrollTop = replyPane.scrollHeight;
        } else if (event.type === 'tool_use') {
          sawToolUse = true;
        }
      }
    }
  } catch (err) {
    replyEl.textContent += `\n[Stream interrupted: ${err.message}]`;
  }

  if (sawToolUse) {
    replyEl.textContent += '\n[Claude wanted to check/edit the Excel sheet, but this window can\'t act on it — try the Excel Add-in instead.]';
  }

  conversation.push({ role: 'assistant', content: fullReply });
}

chatInput.addEventListener('keydown', (event) => {
  if (event.key === 'Enter' && chatInput.value.trim()) {
    const message = chatInput.value.trim();
    chatInput.value = '';
    sendChat(message);
  } else if (event.key === 'Escape') {
    window.api.hidePopup();
  }
});

window.api.onShown(() => {
  chatInput.focus();
});

renderRecentFiles();
