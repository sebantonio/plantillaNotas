const path = require('path');
const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const XLSX = require('xlsx');

let mainWindow;
let selectedExcelPath = null;

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

  saveAlumnosToFile(selectedExcelPath, alumnos);
  return loadAlumnosFromSelectedFile();
});

ipcMain.handle('app:openExternal', async (_event, url) => {
  await shell.openExternal(url);
});

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

function saveAlumnosToFile(filePath, alumnos) {
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

  for (let idx = 0; idx < rowsToClear; idx += 1) {
    const rowIdx = alumnosStart + idx;
    rows[rowIdx] = rows[rowIdx] || [];
    rows[rowIdx][0] = null;
    rows[rowIdx][1] = null;
    rows[rowIdx][2] = null;
  }

  alumnos.forEach((alumno, idx) => {
    const rowIdx = alumnosStart + idx;
    rows[rowIdx] = rows[rowIdx] || [];
    rows[rowIdx][0] = Number(alumno.numero) || idx + 1;
    rows[rowIdx][1] = String(alumno.nombre || '').trim();
    rows[rowIdx][2] = alumno.fechaNac || null;
  });

  const newSheet = XLSX.utils.aoa_to_sheet(rows);
  Object.keys(sheet).forEach((key) => {
    if (key.startsWith('!')) {
      newSheet[key] = sheet[key];
    }
  });

  workbook.Sheets.DATOS = newSheet;
  XLSX.writeFile(workbook, filePath);
}

function findHeaderRow(rows, text) {
  return rows.findIndex((row) => {
    const cell = row && row[1];
    return cell && String(cell).toUpperCase().includes(text);
  });
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
