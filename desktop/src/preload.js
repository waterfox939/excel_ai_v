const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  pickFile: () => ipcRenderer.invoke('files:pick'),
  getRecentFiles: () => ipcRenderer.invoke('files:getRecent'),
  hidePopup: () => ipcRenderer.send('popup:hide'),
  onShown: (callback) => ipcRenderer.on('popup:shown', callback),
  chatPort: 8765,
});
