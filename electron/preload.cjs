const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('triadeElectron', {
  startGoogleOAuth: (options) => ipcRenderer.invoke('google-oauth:start', options),
  refreshGoogleToken: (email, clientId) => ipcRenderer.invoke('google-oauth:refresh', email, clientId),
  unlinkGoogleAccount: (email) => ipcRenderer.invoke('google-oauth:unlink', email),
});