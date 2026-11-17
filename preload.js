const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  openFile: () => ipcRenderer.invoke('open-file'),
  saveFile: (payload) => ipcRenderer.invoke('save-file', payload),
  confirm: (message) => ipcRenderer.invoke('confirm', message),
  setTitle: (title) => ipcRenderer.invoke('set-title', title),
  onMenu: (cb) => ipcRenderer.on('menu', (_evt, action) => cb(action)),
  onRequestClose: (cb) => ipcRenderer.on('request-close', () => cb()),
  closeNow: () => ipcRenderer.send('close-now'),
});
