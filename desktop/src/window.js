// Creates and controls the Spotlight-style popup window.
const path = require('path');
const { BrowserWindow, screen } = require('electron');

const WINDOW_WIDTH = 680;
const WINDOW_HEIGHT = 420;

let popupWindow = null;

function createPopupWindow() {
  popupWindow = new BrowserWindow({
    width: WINDOW_WIDTH,
    height: WINDOW_HEIGHT,
    frame: false,
    resizable: false,
    skipTaskbar: true,
    show: false,
    alwaysOnTop: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  if (process.platform === 'darwin') {
    popupWindow.setAlwaysOnTop(true, 'floating');
    popupWindow.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true });
  }

  popupWindow.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));

  // Never destroy the window — hide it instead, so toggling stays fast and
  // the backend connection stays warm.
  popupWindow.on('close', (event) => {
    event.preventDefault();
    popupWindow.hide();
  });

  popupWindow.on('blur', () => {
    popupWindow.hide();
  });

  return popupWindow;
}

function reposition() {
  const cursorPoint = screen.getCursorScreenPoint();
  const display = screen.getDisplayNearestPoint(cursorPoint);
  const { x, y, width, height } = display.workArea;
  const targetX = Math.round(x + (width - WINDOW_WIDTH) / 2);
  const targetY = Math.round(y + (height - WINDOW_HEIGHT) / 2);
  popupWindow.setPosition(targetX, targetY);
}

function show() {
  reposition();
  popupWindow.show();
  popupWindow.focus();
  popupWindow.webContents.send('popup:shown');
}

function hide() {
  popupWindow.hide();
}

function toggle() {
  if (popupWindow.isVisible()) {
    hide();
  } else {
    show();
  }
}

module.exports = { createPopupWindow, show, hide, toggle };
