const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods that allow the renderer process to use
// the ipcRenderer without exposing the entire object
contextBridge.exposeInMainWorld(
  'api', {
    selectFiles: () => ipcRenderer.invoke('select-files'),
    compareFiles: (file1Path, file2Path) => ipcRenderer.invoke('compare-files', file1Path, file2Path)
  }
);