const path = require('path');
const fs = require('fs');
const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const XLSX = require('xlsx');
const JSZip = require('jszip');

let mainWindow;
let selectedExcelPath = null;
let allowWindowClose = false;
const defaultExcelName = 'plantilla313_dual - copia.xlsx';
const ACTIVITY_TYPES = [
  { key: 'practicas', label: 'Practicas', baseCol: 0 },
  { key: 'memorias', label: 'Memorias', baseCol: 112 },
  { key: 'otros', label: 'Otras actividades', baseCol: 223 },
  { key: 'controles', label: 'Control teorico/practico', baseCol: 334 }
];

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 850,
    minWidth: 900,
    minHeight: 650,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });

  mainWindow.on('close', async (event) => {
    if (allowWindowClose || mainWindow.webContents.isDestroyed()) {
      return;
    }

    event.preventDefault();

    try {
      const canClose = await mainWindow.webContents.executeJavaScript(`
        window.guardarCambiosAntesDeCerrar
          ? window.guardarCambiosAntesDeCerrar()
          : true
      `);

      if (canClose === false) {
        return;
      }
    } catch (error) {
      console.error('No se pudo ejecutar el guardado antes de cerrar:', error);
    }

    allowWindowClose = true;
    mainWindow.close();
  });

  mainWindow.loadFile(path.join(__dirname, 'index.html'));
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    allowWindowClose = false;
    createWindow();
  }
});

ipcMain.handle('excel:selectFile', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Selecciona la plantilla Excel',
    properties: ['openFile'],
    filters: [
      { name: 'Excel', extensions: ['xlsx', 'xlsm', 'xls'] }
    ]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  selectedExcelPath = result.filePaths[0];
  return loadAlumnosFromSelectedFile();
});

ipcMain.handle('excel:getSelectedFile', async () => {
  if (!selectedExcelPath) {
    selectedExcelPath = findDefaultExcelPath();
  }

  if (!selectedExcelPath) {
    return null;
  }

  return loadAlumnosFromSelectedFile();
});

ipcMain.handle('excel:saveAlumnos', async (_event, alumnos) => {
  if (!selectedExcelPath) {
    throw new Error('No hay ningun archivo Excel seleccionado.');
  }

  if (!Array.isArray(alumnos)) {
    throw new Error('La lista de alumnos no es valida.');
  }

  await saveAlumnosToFile(selectedExcelPath, alumnos);
  return loadAlumnosFromSelectedFile();
});

ipcMain.handle('excel:getUnidades', async () => {
  if (!selectedExcelPath) {
    selectedExcelPath = findDefaultExcelPath();
  }

  if (!selectedExcelPath) {
    return null;
  }

  return loadUnidadesFromSelectedFile();
});

ipcMain.handle('excel:saveUnidades', async (_event, unidades) => {
  if (!selectedExcelPath) {
    selectedExcelPath = findDefaultExcelPath();
  }

  if (!selectedExcelPath) {
    throw new Error('No hay ningun archivo Excel seleccionado.');
  }

  if (!Array.isArray(unidades)) {
    throw new Error('La lista de unidades no es valida.');
  }

  await saveUnidadesToFile(selectedExcelPath, unidades);
  return loadUnidadesFromSelectedFile();
});

ipcMain.handle('excel:getRraaCriterios', async () => {
  if (!selectedExcelPath) {
    selectedExcelPath = findDefaultExcelPath();
  }

  if (!selectedExcelPath) {
    return null;
  }

  return loadRraaCriteriosFromSelectedFile();
});

ipcMain.handle('excel:saveRraaCriterios', async (_event, payload) => {
  try {
    if (!selectedExcelPath) {
      selectedExcelPath = findDefaultExcelPath();
    }

    if (!selectedExcelPath) {
      throw new Error('No hay ningun archivo Excel seleccionado.');
    }

    if (!payload) {
      throw new Error('Los datos no tienen un formato valido.');
    }

    const rraa = Array.isArray(payload.rraa) ? payload.rraa : [];
    const criterios = Array.isArray(payload.criterios) ? payload.criterios : [];
    const ponderacionesUnidad = Array.isArray(payload.ponderacionesUnidad) ? payload.ponderacionesUnidad : [];

    await saveRraaCriteriosToFile(selectedExcelPath, rraa, criterios, ponderacionesUnidad);
    return loadRraaCriteriosFromSelectedFile();
  } catch (error) {
    throw new Error(`Error al guardar RRAA/CE: ${error.message}`);
  }
});

ipcMain.handle('excel:getNotasActividad', async (_event, payload = {}) => {
  if (!selectedExcelPath) {
    selectedExcelPath = findDefaultExcelPath();
  }

  if (!selectedExcelPath) {
    return null;
  }

  return loadNotasActividadFromSelectedFile(
    payload.unidad || 'U1',
    payload.tipo || 'practicas',
    Number(payload.actividad) || 1
  );
});

ipcMain.handle('excel:saveNotasActividad', async (_event, payload) => {
  if (!selectedExcelPath) {
    selectedExcelPath = findDefaultExcelPath();
  }

  if (!selectedExcelPath) {
    throw new Error('No hay ningun archivo Excel seleccionado.');
  }

  if (!payload || !payload.unidad || !payload.tipo || !Array.isArray(payload.notas)) {
    throw new Error('Los datos de notas no tienen un formato valido.');
  }

  await saveNotasActividadToFile(
    selectedExcelPath,
    payload.unidad,
    payload.tipo,
    Number(payload.actividad) || 1,
    payload.notas
  );
  return loadNotasActividadFromSelectedFile(
    payload.unidad,
    payload.tipo,
    Number(payload.actividad) || 1
  );
});

ipcMain.handle('excel:getNotasEvaluacion', async (_event, payload = {}) => {
  if (!selectedExcelPath) {
    selectedExcelPath = findDefaultExcelPath();
  }

  if (!selectedExcelPath) {
    return null;
  }

  return loadNotasEvaluacionFromSelectedFile(payload.evaluacion || '1');
});

ipcMain.handle('app:openExternal', async (_event, url) => {
  await shell.openExternal(url);
});

function findDefaultExcelPath() {
  const candidate = path.join(__dirname, defaultExcelName);
  return fs.existsSync(candidate) ? candidate : null;
}

function loadAlumnosFromSelectedFile() {
  const workbook = XLSX.readFile(selectedExcelPath, { cellDates: true });

  if (!workbook.SheetNames.includes('DATOS')) {
    throw new Error('El archivo no tiene la hoja "DATOS" esperada.');
  }

  const sheet = workbook.Sheets.DATOS;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const headerRowIdx = findHeaderRow(rows, 'ALUMNADO');

  if (headerRowIdx === -1) {
    throw new Error('No se encontro la seccion ALUMNADO en la hoja DATOS.');
  }

  const alumnosStart = headerRowIdx + 1;
  const alumnos = [];

  for (let i = alumnosStart; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const nombre = row[1];

    if (!nombre || String(nombre).trim() === '') {
      break;
    }

    alumnos.push({
      numero: row[0] || alumnos.length + 1,
      nombre: String(nombre),
      fechaNac: normaliseDateForUi(row[2])
    });
  }

  return {
    filePath: selectedExcelPath,
    fileName: path.basename(selectedExcelPath),
    alumnos
  };
}

function loadUnidadesFromSelectedFile() {
  const workbook = XLSX.readFile(selectedExcelPath, { cellDates: true });

  if (!workbook.SheetNames.includes('DATOS')) {
    throw new Error('El archivo no tiene la hoja "DATOS" esperada.');
  }

  const sheet = workbook.Sheets.DATOS;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const unidadesStart = findUnidadesStartRow(rows);

  if (unidadesStart === -1) {
    throw new Error('No se encontro la seccion UNIDADES en la hoja DATOS.');
  }

  const unidades = [];

  for (let idx = 0; idx < 16; idx += 1) {
    const row = rows[unidadesStart + idx] || [];
    const codigo = row[8] || `U${idx + 1}`;
    const nombre = row[9] || '';
    const evaluacion = row[10] || '';
    const horas = row[11] ?? '';

    if (codigo || nombre || evaluacion || horas !== '') {
      unidades.push({
        codigo: String(codigo || `U${idx + 1}`),
        nombre: String(nombre || ''),
        evaluacion: String(evaluacion || ''),
        horas: horas === null ? '' : String(horas)
      });
    }
  }

  return {
    filePath: selectedExcelPath,
    fileName: path.basename(selectedExcelPath),
    unidades
  };
}

function loadRraaCriteriosFromSelectedFile() {
  const workbook = XLSX.readFile(selectedExcelPath, { cellDates: true });

  if (!workbook.SheetNames.includes('DATOS') || !workbook.SheetNames.includes('PESOS')) {
    throw new Error('El archivo debe tener las hojas DATOS y PESOS.');
  }

  const datosRows = XLSX.utils.sheet_to_json(workbook.Sheets.DATOS, { header: 1, defval: null });
  const pesosRows = XLSX.utils.sheet_to_json(workbook.Sheets.PESOS, { header: 1, defval: null });
  const rraaStart = findRraaStartRow(datosRows);
  const criterionTexts = extractCriterionTexts(datosRows);

  if (rraaStart === -1) {
    throw new Error('No se encontro la seccion RRAA en la hoja DATOS.');
  }

  const rraa = [];

  for (let i = rraaStart; i < datosRows.length; i += 1) {
    const row = datosRows[i] || [];
    const numero = row[5];
    const descripcion = row[6];

    if ((!numero || String(numero).trim() === '') && (!descripcion || String(descripcion).trim() === '')) {
      break;
    }

    rraa.push({
      numero: numero || rraa.length + 1,
      descripcion: descripcion ? String(descripcion).trim() : ''
    });
  }

  const criterios = [];

  for (let colIdx = 0; colIdx < (pesosRows[3] || []).length; colIdx += 1) {
    const codigo = pesosRows[3] && pesosRows[3][colIdx];

    if (!isCriterionCode(codigo)) {
      continue;
    }

    const raNumero = Number(String(codigo).match(/^(\d+)/)[1]);
    const rraaItem = rraa.find((item) => Number(item.numero) === raNumero);

    criterios.push({
      numero: criterios.length + 1,
      codigo: String(codigo),
      nombre: String(codigo),
      originalCodigo: String(codigo),
      raNumero,
      raDescripcion: rraaItem ? rraaItem.descripcion : '',
      ponderacion: pesosRows[21] && pesosRows[21][colIdx] ? pesosRows[21][colIdx] : 0,
      ponderacionInstituto: pesosRows[21] && pesosRows[21][colIdx + 1] ? pesosRows[21][colIdx + 1] : 0,
      ponderacionEmpresa: pesosRows[21] && pesosRows[21][colIdx + 2] ? pesosRows[21][colIdx + 2] : 0,
      texto: criterionTexts[normalizeCriterionCode(codigo)] || '',
      colIdx
    });
  }

  const ponderacionesUnidad = [];

  for (let rowIdx = 5; rowIdx < 21; rowIdx += 1) {
    const row = pesosRows[rowIdx] || [];
    const nombre = row[0] && String(row[0]) !== '0' ? String(row[0]) : '';
    const ponderaciones = {};

    criterios.forEach((criterio) => {
      ponderaciones[criterio.colIdx] = {
        ponderacion: row[criterio.colIdx] || 0,
        ponderacionInstituto: row[criterio.colIdx + 1] || 0,
        ponderacionEmpresa: row[criterio.colIdx + 2] || 0
      };
    });

    ponderacionesUnidad.push({
      numero: rowIdx - 4,
      rowIdx,
      nombre,
      ponderaciones
    });
  }

  return {
    filePath: selectedExcelPath,
    fileName: path.basename(selectedExcelPath),
    rraa,
    criterios,
    ponderacionesUnidad
  };
}

function loadNotasActividadFromSelectedFile(unidad = 'U1', tipo = 'practicas', actividad = 1) {
  const workbook = XLSX.readFile(selectedExcelPath, { cellDates: true });
  const unidades = listUnitSheets(workbook);
  const selectedUnidad = unidades.find((item) => item.codigo === unidad)?.codigo || unidades[0]?.codigo || unidad;

  if (!workbook.SheetNames.includes(selectedUnidad)) {
    throw new Error(`El archivo no tiene la hoja "${selectedUnidad}".`);
  }

  const sheet = workbook.Sheets[selectedUnidad];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const selectedType = getActivityType(tipo);
  const blocks = findActivityBlocks(rows, selectedType.key);
  const selectedBlock = blocks.find((block) => Number(block.numero) === Number(actividad)) || blocks[0] || null;
  const notas = selectedBlock ? extractActivityNotes(rows, selectedBlock) : [];

  return {
    filePath: selectedExcelPath,
    fileName: path.basename(selectedExcelPath),
    unidad: selectedUnidad,
    tipo: selectedType.key,
    actividad: selectedBlock ? selectedBlock.numero : Number(actividad) || 1,
    unidades,
    tipos: ACTIVITY_TYPES.map((item) => ({
      key: item.key,
      label: item.label,
      actividades: findActivityBlocks(rows, item.key).map(formatActivityBlock)
    })),
    actividades: blocks.map(formatActivityBlock),
    notas,
    block: selectedBlock ? formatActivityBlock(selectedBlock) : null
  };
}

async function saveNotasActividadToFile(filePath, unidad, tipo, actividad, notas) {
  const workbook = XLSX.readFile(filePath, { cellDates: true });

  if (!workbook.SheetNames.includes(unidad)) {
    throw new Error(`El archivo no tiene la hoja "${unidad}".`);
  }

  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[unidad], { header: 1, defval: null });
  const selectedType = getActivityType(tipo);
  const block = findActivityBlocks(rows, selectedType.key).find((item) => Number(item.numero) === Number(actividad));

  if (!block) {
    throw new Error(`No se encontro la actividad ${actividad} de ${selectedType.label} en ${unidad}.`);
  }

  await editWorkbookSheetsXml(filePath, {
    [unidad]: (sheetXml) => {
      let xml = sheetXml;

      notas.forEach((item) => {
        const rowIdx = Number(item.rowIdx);

        if (!Number.isInteger(rowIdx) || rowIdx < block.firstStudentRow) {
          return;
        }

        const existingName = rows[rowIdx] && rows[rowIdx][block.nameCol];
        if (!existingName || String(existingName).trim() === '') {
          return;
        }

        xml = setXmlCell(xml, rowIdx, block.noteCol, normalizeGradeValue(item.nota), 'number');
      });

      return xml;
    }
  });
}

function loadNotasEvaluacionFromSelectedFile(evaluacion = '1') {
  const workbook = XLSX.readFile(selectedExcelPath, { cellDates: true, cellFormula: true });
  const sheetName = findEvaluationSheetName(workbook, evaluacion);

  if (!sheetName) {
    throw new Error(`No se encontro la hoja de la ${evaluacion} evaluacion.`);
  }

  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
  const title = normalizePlainText(sheetName) === 'FINAL'
    ? 'FINAL'
    : getCellDisplay(sheet, 2, 0) || sheetName;
  const layout = findEvaluationLayout(rows);
  const raColumns = findEvaluationRaColumns(sheet, rows, layout);
  const finalColumn = findEvaluationFinalColumn(sheet, rows, layout);
  const criteria = findEvaluationCriteriaColumns(sheet, rows, layout, raColumns, finalColumn);
  const alumnos = extractEvaluationStudents(sheet, rows, layout, raColumns, finalColumn, criteria);

  return {
    filePath: selectedExcelPath,
    fileName: path.basename(selectedExcelPath),
    sheetName,
    title,
    evaluacion: String(evaluacion),
    layout,
    raColumns,
    criteria,
    alumnos
  };
}

function findEvaluationLayout(rows) {
  for (let rowIdx = 0; rowIdx < rows.length; rowIdx += 1) {
    const row = rows[rowIdx] || [];
    const normalizedCells = row.map((value) => normalizePlainText(value));
    const notaCeCount = normalizedCells.filter((value) => value === 'NOTA CE').length;
    const nextRow = rows[rowIdx + 1] || [];
    const nextRowCriteriaCount = nextRow.filter((value) => isEvaluationCriterionCode(value)).length;

    if (notaCeCount > 0 && nextRowCriteriaCount > 0) {
      return {
        summaryRowIdx: rowIdx,
        codeRowIdx: rowIdx + 1,
        firstStudentRowIdx: rowIdx + 2
      };
    }
  }

  throw new Error('No se encontro la cabecera de notas de evaluacion.');
}

function findEvaluationSheetName(workbook, evaluacion) {
  const target = String(evaluacion || '1').trim();

  if (target === 'final') {
    return workbook.SheetNames.find((name) => normalizePlainText(name) === 'FINAL');
  }

  if (target === '2solo' || target === '3solo') {
    const evaluationNumber = target.charAt(0);
    return workbook.SheetNames.find((name) => {
      const normalized = normalizePlainText(name);
      return (
        normalized.includes(evaluationNumber) &&
        normalized.includes('EVA') &&
        normalized.includes('SOLO')
      );
    });
  }

  return workbook.SheetNames.find((name) => {
    const normalized = normalizePlainText(name);
    return (
      normalized.includes(`${target}`) &&
      normalized.includes('EVA') &&
      !normalized.includes('MAX') &&
      !normalized.includes('SOLO')
    );
  });
}

function findEvaluationRaColumns(sheet, rows, layout) {
  const summaryRow = rows[layout.summaryRowIdx] || [];
  const columns = [];

  summaryRow.forEach((value, colIdx) => {
    if (normalizePlainText(value) === 'NOTA CE') {
      const raNumber = findEvaluationRaNumberForBlock(rows, layout, colIdx);
      const label = raNumber ? `RRAA ${raNumber}` : `RRAA ${columns.length + 1}`;
      columns.push({
        colIdx,
        address: columnName(colIdx),
        label,
        numero: raNumber || columns.length + 1,
        peso: getCellDisplay(sheet, 13, colIdx) || ''
      });
    }
  });

  return columns;
}

function findEvaluationRaNumberForBlock(rows, layout, notaCeColIdx) {
  const codeRow = rows[layout.codeRowIdx] || [];

  for (let colIdx = notaCeColIdx + 1; colIdx < codeRow.length; colIdx += 1) {
    const value = codeRow[colIdx];
    const normalizedValue = normalizePlainText(value);

    if (normalizedValue === 'NOTA CE' || normalizedValue === 'NOTA FINAL') {
      break;
    }

    const match = String(value || '').trim().match(/^(\d+)/);
    if (match) {
      return Number(match[1]);
    }
  }

  return null;
}

function findEvaluationFinalColumn(sheet, rows, layout) {
  const summaryRow = rows[layout.summaryRowIdx] || [];
  let colIdx = summaryRow.findIndex((value) => normalizePlainText(value) === 'NOTA FINAL');

  if (colIdx === -1) {
    const titleRow = rows[2] || [];
    colIdx = titleRow.findIndex((value) => normalizePlainText(value) === 'RESUMEN');
  }

  if (colIdx === -1) {
    const codeRow = rows[layout.codeRowIdx] || [];
    for (let idx = codeRow.length - 1; idx >= 0; idx -= 1) {
      if (isEvaluationCriterionCode(codeRow[idx]) || normalizePlainText(codeRow[idx]) === 'REC') {
        colIdx = idx + 1;
        break;
      }
    }
  }

  if (colIdx === -1) {
    throw new Error('No se encontro la columna NOTA FINAL en la evaluacion.');
  }

  return {
    colIdx,
    address: columnName(colIdx),
    label: 'Nota final'
  };
}

function findEvaluationCriteriaColumns(sheet, rows, layout, raColumns, finalColumn) {
  const codeRow = rows[layout.codeRowIdx] || [];
  const criteria = [];

  raColumns.forEach((ra, idx) => {
    const nextRa = raColumns[idx + 1];
    const startCol = ra.colIdx + 1;
    const endCol = nextRa ? nextRa.colIdx - 1 : finalColumn.colIdx - 1;

    for (let colIdx = startCol; colIdx <= endCol; colIdx += 1) {
      const code = codeRow[colIdx];

      if (!isEvaluationCriterionCode(code)) {
        continue;
      }

      criteria.push({
        colIdx,
        address: columnName(colIdx),
        raColIdx: ra.colIdx,
        raLabel: ra.label,
        codigo: String(code).trim(),
        peso: getCellDisplay(sheet, 12, colIdx) || ''
      });
    }
  });

  return criteria;
}

function extractEvaluationStudents(sheet, rows, layout, raColumns, finalColumn, criteria) {
  const alumnos = [];

  for (let rowIdx = layout.firstStudentRowIdx; rowIdx < rows.length; rowIdx += 1) {
    const nombre = getCellDisplay(sheet, rowIdx, 0);

    if (!nombre || String(nombre).trim() === '') {
      continue;
    }

    if (normalizePlainText(nombre).includes('MEDIA') || normalizePlainText(nombre).includes('PONDERACION')) {
      continue;
    }

    alumnos.push({
      rowIdx,
      numero: alumnos.length + 1,
      nombre: String(nombre).trim(),
      final: getEvaluationNumber(sheet, rowIdx, finalColumn.colIdx),
      finalDisplay: getCellDisplay(sheet, rowIdx, finalColumn.colIdx),
      rraa: raColumns.map((ra) => ({
        colIdx: ra.colIdx,
        label: ra.label,
        nota: getEvaluationNumber(sheet, rowIdx, ra.colIdx),
        display: getCellDisplay(sheet, rowIdx, ra.colIdx)
      })),
      criterios: criteria.map((criterion) => ({
        colIdx: criterion.colIdx,
        raColIdx: criterion.raColIdx,
        raLabel: criterion.raLabel,
        codigo: criterion.codigo,
        nota: getEvaluationNumber(sheet, rowIdx, criterion.colIdx),
        display: getCellDisplay(sheet, rowIdx, criterion.colIdx)
      }))
    });
  }

  return alumnos;
}

function isEvaluationCriterionCode(value) {
  return Boolean(value && /^\d+\.?[a-z]\)?$/i.test(String(value).trim()));
}

function getCellDisplay(sheet, rowIdx, colIdx) {
  const address = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
  const cell = sheet[address];

  if (!cell || cell.v === null || cell.v === undefined || cell.v === '') {
    return '';
  }

  return cell.w !== undefined ? String(cell.w) : String(cell.v);
}

function getEvaluationNumber(sheet, rowIdx, colIdx) {
  const address = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
  const cell = sheet[address];
  const value = cell && cell.v !== undefined ? cell.v : null;
  const numeric = Number(value);
  return Number.isNaN(numeric) ? null : numeric;
}

async function saveAlumnosToFile(filePath, alumnos) {
  const workbook = XLSX.readFile(filePath, { cellDates: true });

  if (!workbook.SheetNames.includes('DATOS')) {
    throw new Error('El archivo no tiene la hoja "DATOS" esperada.');
  }

  const sheet = workbook.Sheets.DATOS;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const headerRowIdx = findHeaderRow(rows, 'ALUMNADO');

  if (headerRowIdx === -1) {
    throw new Error('No se encontro la seccion ALUMNADO en la hoja DATOS.');
  }

  const alumnosStart = headerRowIdx + 1;
  const existingCount = countExistingAlumnos(rows, alumnosStart);
  const rowsToClear = Math.max(existingCount, alumnos.length);

  await editDatosSheetXml(filePath, (sheetXml) => {
    let xml = sheetXml;

    for (let idx = 0; idx < rowsToClear; idx += 1) {
      const rowIdx = alumnosStart + idx;
      xml = setXmlCell(xml, rowIdx, 0, null);
      xml = setXmlCell(xml, rowIdx, 1, null);
      xml = setXmlCell(xml, rowIdx, 2, null);
    }

    alumnos.forEach((alumno, idx) => {
      const rowIdx = alumnosStart + idx;
      xml = setXmlCell(xml, rowIdx, 0, Number(alumno.numero) || idx + 1, 'number');
      xml = setXmlCell(xml, rowIdx, 1, String(alumno.nombre || '').trim(), 'text');
      xml = setXmlCell(xml, rowIdx, 2, alumno.fechaNac || null, 'date');
    });

    return xml;
  });
}

async function saveUnidadesToFile(filePath, unidades) {
  const workbook = XLSX.readFile(filePath, { cellDates: true });

  if (!workbook.SheetNames.includes('DATOS')) {
    throw new Error('El archivo no tiene la hoja "DATOS" esperada.');
  }

  const sheet = workbook.Sheets.DATOS;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const unidadesStart = findUnidadesStartRow(rows);

  if (unidadesStart === -1) {
    throw new Error('No se encontro la seccion UNIDADES en la hoja DATOS.');
  }

  const normalized = normalizeUnidades(unidades).slice(0, 16);

  await editDatosSheetXml(filePath, (sheetXml) => {
    let xml = sheetXml;

    for (let idx = 0; idx < 16; idx += 1) {
      const unidad = normalized[idx] || {
        codigo: `U${idx + 1}`,
        nombre: '',
        evaluacion: '',
        horas: ''
      };

      xml = setXmlCell(xml, unidadesStart + idx, 8, unidad.codigo || `U${idx + 1}`, 'text');
      xml = setXmlCell(xml, unidadesStart + idx, 9, unidad.nombre || null, 'text');
      xml = setXmlCell(xml, unidadesStart + idx, 10, unidad.evaluacion || null, 'text');
      xml = setXmlCell(xml, unidadesStart + idx, 11, unidad.horas || null, 'number');
    }

    xml = syncEvaluationUnitBlocksXml(xml, rows, normalized);
    return xml;
  });
}

async function saveRraaCriteriosToFile(filePath, rraa, criterios, ponderacionesUnidad) {
  const workbook = XLSX.readFile(filePath, { cellDates: true });

  if (!workbook.SheetNames.includes('DATOS') || !workbook.SheetNames.includes('PESOS')) {
    throw new Error('El archivo debe tener las hojas DATOS y PESOS.');
  }

  const datosRows = XLSX.utils.sheet_to_json(workbook.Sheets.DATOS, { header: 1, defval: null });
  const rraaStart = findRraaStartRow(datosRows);

  if (rraaStart === -1) {
    throw new Error('No se encontro la seccion RRAA en la hoja DATOS.');
  }

  const normalizedRraa = rraa.map((item, idx) => ({
    numero: Number(item.numero) || idx + 1,
    descripcion: String(item.descripcion || '').trim()
  }));

  const normalizedCriterios = criterios.map((item) => ({
    codigo: normalizeCriterionCodeForPesos(item.codigo || item.nombre),
    originalCodigo: normalizeCriterionCodeForPesos(item.originalCodigo || item.codigo || item.nombre),
    raNumero: Number(item.raNumero) || criterionRaNumber(item.codigo || item.nombre),
    texto: String(item.texto || '').trim(),
    ponderacion: parseDecimal(item.ponderacion),
    ponderacionInstituto: parseDecimal(item.ponderacionInstituto),
    ponderacionEmpresa: parseDecimal(item.ponderacionEmpresa),
    colIdx: Number.isInteger(item.colIdx) ? item.colIdx : null
  }));

  const normalizedUnitWeights = normalizeUnitWeights(ponderacionesUnidad);

  await editWorkbookSheetsXml(filePath, {
    DATOS: (sheetXml) => {
      let xml = sheetXml;

      // Actualizar solo RRAA (no tocar criterios en DATOS)
      normalizedRraa.forEach((item, idx) => {
        const rowIdx = rraaStart + idx;
        xml = setXmlCell(xml, rowIdx, 5, item.numero, 'number');
        xml = setXmlCell(xml, rowIdx, 6, item.descripcion, 'text');
      });

      // Actualizar texto de criterios SOLO si encontramos su fila específica
      normalizedCriterios.forEach((criterio) => {
        const rowIdx = findCriterionTextRow(datosRows, criterio.originalCodigo) !== -1
          ? findCriterionTextRow(datosRows, criterio.originalCodigo)
          : findCriterionTextRow(datosRows, criterio.codigo);
        if (rowIdx !== -1) {
          xml = setXmlCell(xml, rowIdx, 21, stripCriterionClosingParen(criterio.codigo), 'text');
          xml = setXmlCell(xml, rowIdx, 22, criterio.texto || null, 'text');
        }
      });

      return xml;
    },
    PESOS: (sheetXml) => {
      let xml = sheetXml;

      // Actualizar códigos de criterios en fila 3
      normalizedCriterios.forEach((criterio) => {
        const colIdx = criterio.colIdx ?? findCriterionColumnForSave(criterio, normalizedCriterios);

        if (colIdx === -1) {
          return;
        }

        xml = setXmlCell(xml, 3, colIdx, criterio.codigo, 'text');
      });

      // Actualizar ponderaciones por unidad (filas 5-20)
      normalizedUnitWeights.forEach((unidad) => {
        Object.entries(unidad.ponderaciones).forEach(([colKey, values]) => {
          const colIdx = Number(colKey);
          if (colIdx >= 0) {
            xml = setXmlCell(xml, unidad.rowIdx, colIdx, parseDecimal(values.ponderacion), 'number');
            xml = setXmlCell(xml, unidad.rowIdx, colIdx + 1, parseDecimal(values.ponderacionInstituto), 'number');
            xml = setXmlCell(xml, unidad.rowIdx, colIdx + 2, parseDecimal(values.ponderacionEmpresa), 'number');
          }
        });
      });

      return xml;
    }
  });
}

function listUnitSheets(workbook) {
  const unitsFromDatos = {};

  if (workbook.Sheets.DATOS) {
    const datosRows = XLSX.utils.sheet_to_json(workbook.Sheets.DATOS, { header: 1, defval: null });
    const unidadesStart = findUnidadesStartRow(datosRows);

    if (unidadesStart !== -1) {
      for (let idx = 0; idx < 16; idx += 1) {
        const row = datosRows[unidadesStart + idx] || [];
        const codigo = String(row[8] || `U${idx + 1}`).trim();
        const nombre = String(row[9] || '').trim();

        if (codigo) {
          unitsFromDatos[codigo.toUpperCase()] = nombre;
        }
      }
    }
  }

  return workbook.SheetNames
    .filter((name) => /^U\d+$/i.test(name))
    .sort((a, b) => Number(a.replace(/\D/g, '')) - Number(b.replace(/\D/g, '')))
    .map((codigo) => {
      const nombre = unitsFromDatos[codigo.toUpperCase()] || '';
      return {
        codigo,
        nombre,
        label: nombre ? `${codigo} - ${nombre}` : codigo
      };
    });
}

function getActivityType(tipo) {
  return ACTIVITY_TYPES.find((item) => item.key === tipo) || ACTIVITY_TYPES[0];
}

function findActivityBlocks(rows, tipoKey) {
  const activityType = getActivityType(tipoKey);
  const blocks = [];

  for (let rowIdx = 0; rowIdx < rows.length - 4; rowIdx += 1) {
    const numberCell = rows[rowIdx + 1] && rows[rowIdx + 1][activityType.baseCol + 1];
    const activityNumber = Number(numberCell);

    if (numberCell === null || numberCell === undefined || numberCell === '' || Number.isNaN(activityNumber)) {
      continue;
    }

    const headerRow = rows[rowIdx + 3] || [];
    const nameCol = findActivityHeaderCol(headerRow, activityType.baseCol, 'NOMBRE Y APELLIDOS');
    const noteCol = findActivityHeaderCol(headerRow, activityType.baseCol, 'NOTA FINAL');

    if (nameCol === -1 || noteCol === -1) {
      continue;
    }

    blocks.push({
      tipo: activityType.key,
      tipoLabel: activityType.label,
      numero: activityNumber,
      titleRow: rowIdx,
      numberRow: rowIdx + 1,
      headerRow: rowIdx + 3,
      firstStudentRow: rowIdx + 4,
      nameCol,
      noteCol
    });
  }

  return blocks;
}

function findActivityHeaderCol(row, startCol, expectedText) {
  const normalizedExpected = normalizePlainText(expectedText);

  for (let colIdx = startCol; colIdx < startCol + 8; colIdx += 1) {
    if (normalizePlainText(row[colIdx]).includes(normalizedExpected)) {
      return colIdx;
    }
  }

  return -1;
}

function extractActivityNotes(rows, block) {
  const notes = [];

  for (let rowIdx = block.firstStudentRow; rowIdx < rows.length; rowIdx += 1) {
    const row = rows[rowIdx] || [];
    const nombre = row[block.nameCol];

    if (!nombre || String(nombre).trim() === '' || String(nombre).trim() === '0') {
      break;
    }

    notes.push({
      numero: notes.length + 1,
      rowIdx,
      nombre: String(nombre).trim(),
      nota: formatValueForUi(row[block.noteCol])
    });
  }

  return notes;
}

function formatActivityBlock(block) {
  return {
    numero: block.numero,
    label: `${block.tipoLabel} ${block.numero}`,
    firstStudentRow: block.firstStudentRow,
    noteCol: block.noteCol
  };
}

function normalizeGradeValue(value) {
  if (value === null || value === undefined || String(value).trim() === '') {
    return null;
  }

  const numeric = Number(String(value).replace(',', '.'));
  return Number.isNaN(numeric) ? null : numeric;
}

function formatValueForUi(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).replace('.', ',');
}

function normalizePlainText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim();
}

function findHeaderRow(rows, text) {
  return rows.findIndex((row) => {
    const cell = row && row[1];
    return cell && String(cell).toUpperCase().includes(text);
  });
}

function findUnidadesStartRow(rows) {
  const headerRowIdx = rows.findIndex((row) => {
    const unidadHeader = row && row[9];
    return (
      unidadHeader &&
      row[8] &&
      String(unidadHeader).toUpperCase().includes('UNIDADES')
    );
  });

  return headerRowIdx === -1 ? -1 : headerRowIdx + 1;
}

function findRraaStartRow(rows) {
  const headerRowIdx = rows.findIndex((row) => {
    const rraaHeader = row && row[6];
    return rraaHeader && String(rraaHeader).toUpperCase().includes('RRAA');
  });

  return headerRowIdx === -1 ? -1 : headerRowIdx + 1;
}

function findCriteriosStartRow(rows) {
  // Busca la fila donde empieza la sección de criterios (donde está el primer código en columna 21)
  // Generalmente está después de la sección de unidades
  let startRow = 0;

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const codigo = row[21];

    // El primer código de criterio tiene el formato "1.a)"
    if (codigo && String(codigo).match(/^\d+\.[a-z]\)/i)) {
      startRow = i;
      break;
    }
  }

  return startRow > 0 ? startRow : -1;
}

function normalizeUnitWeights(ponderacionesUnidad) {
  return ponderacionesUnidad
    .filter((unidad) => Number.isInteger(unidad.rowIdx) || !Number.isNaN(Number(unidad.rowIdx)))
    .map((unidad) => {
      const ponderaciones = {};

      Object.entries(unidad.ponderaciones || {}).forEach(([colKey, values]) => {
        const colIdx = Number(colKey);

        if (Number.isNaN(colIdx)) {
          return;
        }

        const instituto = parseDecimal(values.ponderacionInstituto);
        const empresa = parseDecimal(values.ponderacionEmpresa);
        ponderaciones[colIdx] = {
          ponderacionInstituto: instituto,
          ponderacionEmpresa: empresa,
          ponderacion: Number((instituto + empresa).toFixed(6))
        };
      });

      return {
        rowIdx: Number(unidad.rowIdx),
        ponderaciones
      };
    });
}

function isCriterionCode(value) {
  return Boolean(value && /^\d+\.[a-z]\)/i.test(String(value).trim()));
}

function normalizeCriterionCodeForPesos(value) {
  const text = String(value || '').trim();

  if (!text) {
    return '';
  }

  return text.endsWith(')') ? text : `${text})`;
}

function stripCriterionClosingParen(value) {
  return String(value || '').trim().replace(/\)$/u, '');
}

function normalizeCriterionCode(value) {
  return stripCriterionClosingParen(value).toLowerCase();
}

function parseDecimal(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  const number = Number(String(value).replace(',', '.'));
  return Number.isNaN(number) ? 0 : number;
}

function extractCriterionTexts(rows) {
  const texts = {};

  rows.forEach((row) => {
    const code = row && row[21];
    const text = row && row[22];

    if (code && text && isCriterionCode(`${code})`)) {
      texts[normalizeCriterionCode(code)] = String(text).trim();
    }
  });

  return texts;
}

function findCriterionTextRow(rows, code) {
  const normalized = normalizeCriterionCode(code);

  return rows.findIndex((row) => {
    const rowCode = row && row[21];
    return rowCode && normalizeCriterionCode(rowCode) === normalized;
  });
}

function criterionRaNumber(value) {
  const match = String(value || '').match(/^(\d+)/);
  return match ? Number(match[1]) : 0;
}

function findCriterionColumnForSave(criterion, criterios) {
  if (!criterion.raNumero) {
    return -1;
  }

  const usedColumns = new Set(
    criterios
      .map((item) => item.colIdx)
      .filter((colIdx) => Number.isInteger(colIdx))
  );
  const start = 2 + (criterion.raNumero - 1) * 40;

  for (let slot = 0; slot < 13; slot += 1) {
    const candidate = start + slot * 3;

    if (!usedColumns.has(candidate)) {
      criterion.colIdx = candidate;
      usedColumns.add(candidate);
      return candidate;
    }
  }

  return -1;
}

function normalizeUnidades(unidades) {
  return unidades
    .map((unidad, idx) => ({
      codigo: String(unidad.codigo || `U${idx + 1}`).trim(),
      nombre: String(unidad.nombre || '').trim(),
      evaluacion: String(unidad.evaluacion || '').trim(),
      horas: String(unidad.horas || '').trim()
    }))
    .filter((unidad) => unidad.codigo || unidad.nombre || unidad.evaluacion || unidad.horas);
}

function syncEvaluationUnitBlocks(sheet, rows, unidades) {
  const evaluations = ['1ª', '2ª', '3ª'];

  evaluations.forEach((evaluation) => {
    const startRow = findEvaluationBlockStartRow(rows, evaluation);

    if (startRow === -1) {
      return;
    }

    const unitsForEvaluation = unidades.filter((unidad) => unidad.evaluacion === evaluation);

    for (let idx = 0; idx < 16; idx += 1) {
      const unidad = unitsForEvaluation[idx];
      setCell(sheet, startRow + idx, 9, `U${idx + 1}`);
      setCell(sheet, startRow + idx, 10, unidad ? unidad.nombre : '');
      setCell(sheet, startRow + idx, 11, evaluation);
    }
  });
}

function findEvaluationBlockStartRow(rows, evaluation) {
  const normalizedEvaluation = evaluation.replace('ª', '');
  const headerRowIdx = rows.findIndex((row) => {
    const header = row && row[10];
    return (
      header &&
      String(header).toUpperCase().includes('UNIDADES') &&
      String(header).includes(normalizedEvaluation)
    );
  });

  return headerRowIdx === -1 ? -1 : headerRowIdx + 1;
}

function setCell(sheet, rowIdx, colIdx, value) {
  const address = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });

  if (value === null || value === undefined || value === '') {
    delete sheet[address];
    return;
  }

  const numericValue = Number(value);
  if (colIdx === 11 && value !== '' && !Number.isNaN(numericValue)) {
    sheet[address] = { t: 'n', v: numericValue };
    return;
  }

  sheet[address] = { t: 's', v: String(value) };
}

async function editDatosSheetXml(filePath, editSheetXml) {
  return editWorkbookSheetsXml(filePath, { DATOS: editSheetXml });
}

async function editWorkbookSheetsXml(filePath, sheetEdits) {
  const input = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(input);
  const entries = Object.entries(sheetEdits);

  for (const [sheetName, editSheetXml] of entries) {
    const sheetPath = await findWorksheetPath(zip, sheetName);
    const sheetFile = zip.file(sheetPath);

    if (!sheetFile) {
      throw new Error(`No se encontro el XML de la hoja ${sheetName}.`);
    }

    const originalXml = await sheetFile.async('string');
    const updatedXml = editSheetXml(originalXml);
    zip.file(sheetPath, updatedXml);
  }

  await removeCalcChain(zip);
  await forceWorkbookRecalculation(zip);

  const output = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });

  fs.writeFileSync(filePath, output);
}

async function findWorksheetPath(zip, sheetName) {
  const workbookXml = await readZipText(zip, 'xl/workbook.xml');
  const relsXml = await readZipText(zip, 'xl/_rels/workbook.xml.rels');
  const sheetRegex = /<sheet\b[^>]*>/g;
  let sheetMatch;

  while ((sheetMatch = sheetRegex.exec(workbookXml)) !== null) {
    const sheetTag = sheetMatch[0];

    if (getXmlAttribute(sheetTag, 'name') !== sheetName) {
      continue;
    }

    const relId = getXmlAttribute(sheetTag, 'r:id');
    const relRegex = /<Relationship\b[^>]*>/g;
    let relMatch;

    while ((relMatch = relRegex.exec(relsXml)) !== null) {
      const relTag = relMatch[0];

      if (getXmlAttribute(relTag, 'Id') === relId) {
        const target = getXmlAttribute(relTag, 'Target');
        return target.startsWith('xl/') ? target : `xl/${target}`;
      }
    }
  }

  throw new Error(`No se encontro la hoja "${sheetName}" dentro del libro.`);
}

async function readZipText(zip, fileName) {
  const file = zip.file(fileName);

  if (!file) {
    throw new Error(`No se encontro ${fileName} en el Excel.`);
  }

  return file.async('string');
}

async function removeCalcChain(zip) {
  zip.remove('xl/calcChain.xml');

  const relsPath = 'xl/_rels/workbook.xml.rels';
  const relsFile = zip.file(relsPath);
  if (relsFile) {
    const relsXml = await relsFile.async('string');
    const cleanedRelsXml = relsXml.replace(
      /<Relationship\b[^>]*Type="[^"]*\/calcChain"[^>]*(?:\/>|><\/Relationship>)/g,
      ''
    );
    zip.file(relsPath, cleanedRelsXml);
  }

  const contentTypesPath = '[Content_Types].xml';
  const contentTypesFile = zip.file(contentTypesPath);
  if (contentTypesFile) {
    const contentTypesXml = await contentTypesFile.async('string');
    const cleanedContentTypesXml = contentTypesXml.replace(
      /<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*(?:\/>|><\/Override>)/g,
      ''
    );
    zip.file(contentTypesPath, cleanedContentTypesXml);
  }
}

async function forceWorkbookRecalculation(zip) {
  const workbookPath = 'xl/workbook.xml';
  const workbookFile = zip.file(workbookPath);

  if (!workbookFile) {
    return;
  }

  let workbookXml = await workbookFile.async('string');

  if (/<calcPr\b[^>]*\/>/.test(workbookXml)) {
    workbookXml = workbookXml.replace(/<calcPr\b[^>]*\/>/, (tag) => {
      let next = tag;
      next = setXmlAttribute(next, 'calcMode', 'auto');
      next = setXmlAttribute(next, 'fullCalcOnLoad', '1');
      next = setXmlAttribute(next, 'forceFullCalc', '1');
      return next;
    });
  } else if (/<calcPr\b[^>]*>[\s\S]*?<\/calcPr>/.test(workbookXml)) {
    workbookXml = workbookXml.replace(/<calcPr\b[^>]*>/, (tag) => {
      let next = tag;
      next = setXmlAttribute(next, 'calcMode', 'auto');
      next = setXmlAttribute(next, 'fullCalcOnLoad', '1');
      next = setXmlAttribute(next, 'forceFullCalc', '1');
      return next;
    });
  } else {
    workbookXml = workbookXml.replace(
      '</workbook>',
      '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>'
    );
  }

  zip.file(workbookPath, workbookXml);
}

function setXmlAttribute(tag, name, value) {
  const regex = new RegExp(`\\b${name}="[^"]*"`);

  if (regex.test(tag)) {
    return tag.replace(regex, `${name}="${escapeXml(value)}"`);
  }

  if (tag.endsWith('/>')) {
    return tag.replace('/>', ` ${name}="${escapeXml(value)}"/>`);
  }

  return tag.replace('>', ` ${name}="${escapeXml(value)}">`);
}

function getXmlAttribute(tag, name) {
  const escapedName = name.replace(':', '\\:');
  const regex = new RegExp(`${escapedName}="([^"]*)"`);
  const match = tag.match(regex);
  return match ? unescapeXml(match[1]) : null;
}

function syncEvaluationUnitBlocksXml(sheetXml, rows, unidades) {
  let xml = sheetXml;
  const evaluations = ['1ª', '2ª', '3ª'];

  evaluations.forEach((evaluation) => {
    const startRow = findEvaluationBlockStartRow(rows, evaluation);

    if (startRow === -1) {
      return;
    }

    const unitsForEvaluation = unidades.filter((unidad) =>
      normalizeEvaluationLabel(unidad.evaluacion) === evaluation
    );

    for (let idx = 0; idx < 16; idx += 1) {
      const unidad = unitsForEvaluation[idx];
      xml = setXmlCell(xml, startRow + idx, 9, `U${idx + 1}`, 'text');
      xml = setXmlCell(xml, startRow + idx, 10, unidad ? unidad.nombre : null, 'text');
      xml = setXmlCell(xml, startRow + idx, 11, evaluation, 'text');
    }
  });

  return xml;
}

function setXmlCell(sheetXml, rowIdx, colIdx, value, valueType = 'text') {
  const rowNumber = rowIdx + 1;
  const cellRef = `${columnName(colIdx)}${rowNumber}`;
  const rowRegex = new RegExp(`<row\\b[^>]*\\br="${rowNumber}"[^>]*>[\\s\\S]*?<\\/row>`);
  const rowMatch = sheetXml.match(rowRegex);
  let xml = sheetXml;

  if (!rowMatch) {
    xml = insertXmlRow(xml, rowNumber);
    return setXmlCell(xml, rowIdx, colIdx, value, valueType);
  }

  const originalRow = rowMatch[0];
  const cellRegex = new RegExp(`<c\\b[^>]*\\br="${escapeRegex(cellRef)}"[^>]*(?:>[\\s\\S]*?<\\/c>|\\s*\\/>)`);

  if (value === null || value === undefined || value === '') {
    const clearedRow = originalRow.replace(cellRegex, '');
    return xml.replace(originalRow, clearedRow);
  }

  const existingCell = originalRow.match(cellRegex);
  const styleId = existingCell ? getXmlAttribute(existingCell[0], 's') : null;
  const newCell = buildXmlCell(cellRef, value, valueType, styleId);
  const updatedRow = cellRegex.test(originalRow)
    ? originalRow.replace(cellRegex, newCell)
    : insertXmlCellInRow(originalRow, newCell, colIdx);

  return xml.replace(originalRow, updatedRow);
}

function insertXmlRow(sheetXml, rowNumber) {
  const newRow = `<row r="${rowNumber}"></row>`;
  const sheetDataClose = '</sheetData>';

  if (!sheetXml.includes(sheetDataClose)) {
    throw new Error('La hoja DATOS no tiene una estructura sheetData valida.');
  }

  return sheetXml.replace(sheetDataClose, `${newRow}${sheetDataClose}`);
}

function insertXmlCellInRow(rowXml, cellXml, colIdx) {
  const cells = [...rowXml.matchAll(/<c\b[^>]*\br="([A-Z]+)\d+"[^>]*(?:>[\s\S]*?<\/c>|\s*\/>)/g)];

  for (const cell of cells) {
    if (columnIndex(cell[1]) > colIdx) {
      return rowXml.replace(cell[0], `${cellXml}${cell[0]}`);
    }
  }

  return rowXml.replace('</row>', `${cellXml}</row>`);
}

function buildXmlCell(cellRef, value, valueType, styleId = null) {
  const styleAttribute = styleId ? ` s="${escapeXml(styleId)}"` : '';

  if (valueType === 'number') {
    const numberValue = Number(value);

    if (!Number.isNaN(numberValue)) {
      return `<c r="${cellRef}"${styleAttribute}><v>${numberValue}</v></c>`;
    }
  }

  if (valueType === 'date') {
    const serial = excelSerialFromDate(value);

    if (serial !== null) {
      return `<c r="${cellRef}"${styleAttribute}><v>${serial}</v></c>`;
    }
  }

  const text = String(value);
  const spaceAttribute = text.trim() !== text ? ' xml:space="preserve"' : '';
  return `<c r="${cellRef}"${styleAttribute} t="inlineStr"><is><t${spaceAttribute}>${escapeXml(text)}</t></is></c>`;
}

function excelSerialFromDate(value) {
  if (!value) {
    return null;
  }

  const date = value instanceof Date ? value : new Date(`${value}T00:00:00Z`);

  if (Number.isNaN(date.getTime())) {
    return null;
  }

  return Math.floor(date.getTime() / 86400000) + 25569;
}

function columnName(colIdx) {
  let name = '';
  let value = colIdx + 1;

  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }

  return name;
}

function columnIndex(name) {
  return name.split('').reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function unescapeXml(value) {
  return String(value)
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, '>')
    .replace(/&lt;/g, '<')
    .replace(/&amp;/g, '&');
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeEvaluationLabel(value) {
  const text = String(value || '').trim();

  if (text.includes('1')) {
    return '1ª';
  }

  if (text.includes('2')) {
    return '2ª';
  }

  if (text.includes('3')) {
    return '3ª';
  }

  return text;
}

function countExistingAlumnos(rows, alumnosStart) {
  let count = 0;

  for (let i = alumnosStart; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const nombre = row[1];

    if (!nombre || String(nombre).trim() === '') {
      break;
    }

    count += 1;
  }

  return count;
}

function countExistingRraa(rows, rraaStart) {
  let count = 0;

  for (let i = rraaStart; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const numero = row[5];
    const descripcion = row[6];

    if ((!numero || String(numero).trim() === '') && (!descripcion || String(descripcion).trim() === '')) {
      break;
    }

    count += 1;
  }

  return count;
}

function normaliseDateForUi(value) {
  if (!value) {
    return '';
  }

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === 'number') {
    const date = new Date((value - 25569) * 86400 * 1000);
    return date.toISOString().slice(0, 10);
  }

  return String(value);
}
