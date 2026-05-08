const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronExcel', {
  selectFile: () => ipcRenderer.invoke('excel:selectFile'),
  getSelectedFile: () => ipcRenderer.invoke('excel:getSelectedFile'),
  saveAlumnos: (alumnos) => ipcRenderer.invoke('excel:saveAlumnos', alumnos),
  getUnidades: () => ipcRenderer.invoke('excel:getUnidades'),
  saveUnidades: (unidades) => ipcRenderer.invoke('excel:saveUnidades', unidades),
  getRraaCriterios: () => ipcRenderer.invoke('excel:getRraaCriterios'),
  saveRraaCriterios: (payloadOrRraa, criterios, ponderacionesUnidad = []) => {
    const payload = Array.isArray(payloadOrRraa)
      ? { rraa: payloadOrRraa, criterios, ponderacionesUnidad }
      : payloadOrRraa;
    return ipcRenderer.invoke('excel:saveRraaCriterios', payload);
  },
  getNotasActividad: (payload) => ipcRenderer.invoke('excel:getNotasActividad', payload),
  getNotasActividadesTipo: (payload) => ipcRenderer.invoke('excel:getNotasActividadesTipo', payload),
  saveNotasActividad: (payload) => ipcRenderer.invoke('excel:saveNotasActividad', payload),
  addActividad: (payload) => ipcRenderer.invoke('excel:addActividad', payload),
  getNotasEvaluacion: (payload) => ipcRenderer.invoke('excel:getNotasEvaluacion', payload),
  setSelectedFile: (filePath) => ipcRenderer.invoke('excel:setSelectedFile', filePath),
  openExternal: (url) => ipcRenderer.invoke('app:openExternal', url)
});
