const path = require('path');
const fs = require('fs');
const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const XLSX = require('xlsx');
const JSZip = require('jszip');

let mainWindow;
let selectedExcelPath = null;
const defaultExcelName = 'plantilla313_dual - copia.xlsx';

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
  if (!selectedExcelPath) {
    selectedExcelPath = findDefaultExcelPath();
  }

  if (!selectedExcelPath) {
    throw new Error('No hay ningun archivo Excel seleccionado.');
  }

  if (!payload || !Array.isArray(payload.rraa) || !Array.isArray(payload.criterios)) {
    throw new Error('Los RRAA y criterios no tienen un formato valido.');
  }

  await saveRraaCriteriosToFile(selectedExcelPath, payload.rraa, payload.criterios);
  return loadRraaCriteriosFromSelectedFile();
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

  if (rraaStart === -1) {
    throw new Error('No se encontro la seccion RRAA en la hoja DATOS.');
  }

  const rraa = [];

  for (let i = rraaStart; i < datosRows.length; i += 1) {
    const row = datosRows[i] || [];
    const descripcion = row[6];

    if (!descripcion || String(descripcion).trim() === '') {
      break;
    }

    rraa.push({
      numero: row[5] || rraa.length + 1,
      descripcion: String(descripcion)
    });
  }

  const criterios = [];

  for (let i = 5; i < 21; i += 1) {
    const row = pesosRows[i] || [];
    const nombre = row[0];

    if (nombre && String(nombre).trim() !== '') {
      criterios.push({
        numero: criterios.length + 1,
        nombre: String(nombre),
        ponderacion: row[2] || 0
      });
    }
  }

  return {
    filePath: selectedExcelPath,
    fileName: path.basename(selectedExcelPath),
    rraa,
    criterios
  };
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

async function saveRraaCriteriosToFile(filePath, rraa, criterios) {
  const workbook = XLSX.readFile(filePath, { cellDates: true });

  if (!workbook.SheetNames.includes('DATOS') || !workbook.SheetNames.includes('PESOS')) {
    throw new Error('El archivo debe tener las hojas DATOS y PESOS.');
  }

  const datosRows = XLSX.utils.sheet_to_json(workbook.Sheets.DATOS, { header: 1, defval: null });
  const rraaStart = findRraaStartRow(datosRows);

  if (rraaStart === -1) {
    throw new Error('No se encontro la seccion RRAA en la hoja DATOS.');
  }

  const existingRraaCount = countExistingRraa(datosRows, rraaStart);
  const rraaRowsToClear = Math.max(existingRraaCount, rraa.length);
  const normalizedRraa = rraa.map((item, idx) => ({
    numero: Number(item.numero) || idx + 1,
    descripcion: String(item.descripcion || '').trim()
  }));
  const normalizedCriterios = criterios.slice(0, 16).map((item) => ({
    nombre: String(item.nombre || '').trim(),
    ponderacion: Number(item.ponderacion) || 0
  }));

  await editWorkbookSheetsXml(filePath, {
    DATOS: (sheetXml) => {
      let xml = sheetXml;

      for (let idx = 0; idx < rraaRowsToClear; idx += 1) {
        const rowIdx = rraaStart + idx;
        xml = setXmlCell(xml, rowIdx, 5, null);
        xml = setXmlCell(xml, rowIdx, 6, null);
      }

      normalizedRraa.forEach((item, idx) => {
        const rowIdx = rraaStart + idx;
        xml = setXmlCell(xml, rowIdx, 5, item.numero, 'number');
        xml = setXmlCell(xml, rowIdx, 6, item.descripcion, 'text');
      });

      return xml;
    },
    PESOS: (sheetXml) => {
      let xml = sheetXml;

      for (let idx = 0; idx < 16; idx += 1) {
        const rowIdx = 5 + idx;
        const criterio = normalizedCriterios[idx];
        xml = setXmlCell(xml, rowIdx, 0, criterio ? criterio.nombre : null, 'text');
        xml = setXmlCell(xml, rowIdx, 2, criterio ? criterio.ponderacion : null, 'number');
      }

      return xml;
    }
  });
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
    const descripcion = row[6];

    if (!descripcion || String(descripcion).trim() === '') {
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
