(function () {
  if (window.electronExcel) {
    return;
  }

  const tauriCore = window.__TAURI__ && window.__TAURI__.core;

  if (!tauriCore || typeof tauriCore.invoke !== "function") {
    return;
  }

  const invoke = tauriCore.invoke;

  window.electronExcel = {
    selectFile: () => invoke("excel_select_file"),
    getSelectedFile: () => invoke("excel_get_selected_file"),
    saveAlumnos: (alumnos) => invoke("excel_save_alumnos", { alumnos }),
    getUnidades: () => invoke("excel_get_unidades"),
    saveUnidades: (unidades) => invoke("excel_save_unidades", { unidades }),
    getRraaCriterios: () => invoke("excel_get_rraa_criterios"),
    saveRraaCriterios: (payloadOrRraa, criterios, ponderacionesUnidad = []) => {
      const payload = Array.isArray(payloadOrRraa)
        ? { rraa: payloadOrRraa, criterios, ponderacionesUnidad }
        : payloadOrRraa;
      return invoke("excel_save_rraa_criterios", { payload });
    },
    getNotasActividad: (payload) => invoke("excel_get_notas_actividad", { payload }),
    saveNotasActividad: (payload) => invoke("excel_save_notas_actividad", { payload }),
    getNotasEvaluacion: (payload) => invoke("excel_get_notas_evaluacion", { payload }),
    getNotasEvaluacionAlumno: (payload) => invoke("excel_get_notas_evaluacion_alumno", { payload }),
    getAlumnosInformes: () => invoke("excel_get_alumnos_informes"),
    openExternal: (url) => invoke("app_open_external", { url })
  };
})();
