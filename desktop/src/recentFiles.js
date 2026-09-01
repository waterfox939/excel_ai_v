// Tracks recently-picked Excel files and exposes the native "Browse..." dialog.
const path = require('path');
const { dialog } = require('electron');
const Store = require('electron-store');

const store = new Store({ name: 'recent-files' });
const MAX_RECENT = 8;

function getRecentFiles() {
  return store.get('files', []);
}

function addRecentFile(filePath) {
  const existing = getRecentFiles().filter((f) => f.path !== filePath);
  const updated = [
    { path: filePath, name: path.basename(filePath), lastOpened: Date.now() },
    ...existing,
  ].slice(0, MAX_RECENT);
  store.set('files', updated);
  return updated;
}

async function pickFile(parentWindow) {
  const result = await dialog.showOpenDialog(parentWindow, {
    properties: ['openFile'],
    filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  const filePath = result.filePaths[0];
  addRecentFile(filePath);
  return filePath;
}

module.exports = { getRecentFiles, addRecentFile, pickFile };
