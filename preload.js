const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronExcel', {
  selectFile: () => ipcRenderer.invoke('excel:selectFile'),
  getSelectedFile: () => ipcRenderer.invoke('excel:getSelectedFile'),
  saveAlumnos: (alumnos) => ipcRenderer.invoke('excel:saveAlumnos', alumnos),
  openExternal: (url) => ipcRenderer.invoke('app:openExternal', url)
});
