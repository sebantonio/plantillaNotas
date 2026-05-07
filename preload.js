const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronExcel', {
  selectFile: () => ipcRenderer.invoke('excel:selectFile'),
  getSelectedFile: () => ipcRenderer.invoke('excel:getSelectedFile'),
  saveAlumnos: (alumnos) => ipcRenderer.invoke('excel:saveAlumnos', alumnos),
  getUnidades: () => ipcRenderer.invoke('excel:getUnidades'),
  saveUnidades: (unidades) => ipcRenderer.invoke('excel:saveUnidades', unidades),
  getRraaCriterios: () => ipcRenderer.invoke('excel:getRraaCriterios'),
  saveRraaCriterios: (rraa, criterios) => ipcRenderer.invoke('excel:saveRraaCriterios', { rraa, criterios }),
  openExternal: (url) => ipcRenderer.invoke('app:openExternal', url)
});
