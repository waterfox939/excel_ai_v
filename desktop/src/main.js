const path = require('path');
const { app, globalShortcut, ipcMain, Tray, Menu } = require('electron');

const pythonBackend = require('./pythonBackend');
const popup = require('./window');
const recentFiles = require('./recentFiles');

const HOTKEY = 'CommandOrControl+Shift+E';

let tray = null;

function registerHotkey(accelerator, handler) {
  globalShortcut.unregisterAll();
  globalShortcut.register(accelerator, handler);
}

function createTray() {
  tray = new Tray(path.join(__dirname, '..', 'renderer', 'tray-icon.png'));
  tray.setToolTip('Excel AI Agent');
  tray.setContextMenu(
    Menu.buildFromTemplate([{ label: 'Quit', click: () => app.quit() }])
  );
}

app.whenReady().then(async () => {
  pythonBackend.start();
  await pythonBackend.waitUntilHealthy();

  const win = popup.createPopupWindow();

  if (process.argv.includes('--dev')) {
    win.webContents.openDevTools({ mode: 'detach' });
  }

  ipcMain.handle('files:pick', () => recentFiles.pickFile(win));
  ipcMain.handle('files:getRecent', () => recentFiles.getRecentFiles());
  ipcMain.on('popup:hide', () => popup.hide());

  registerHotkey(HOTKEY, () => popup.toggle());
  createTray();
});

app.on('will-quit', () => {
  globalShortcut.unregisterAll();
  pythonBackend.stop();
});

// Popup app with no dock window — don't quit when the (hidden) window "closes".
app.on('window-all-closed', (event) => {
  event.preventDefault();
});
