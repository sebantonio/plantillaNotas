const fs = require('fs');
const path = require('path');
const { commands, setSelectedExcelPath } = require('./main');

const stateFile = path.join(__dirname, '.tauri-selected-excel.json');

function readInput() {
  const raw = fs.readFileSync(0, 'utf8').trim();
  return raw ? JSON.parse(raw) : {};
}

function readSelectedPath() {
  try {
    const state = JSON.parse(fs.readFileSync(stateFile, 'utf8'));
    return state && state.filePath ? state.filePath : null;
  } catch (_error) {
    return null;
  }
}

function writeSelectedPath(filePath) {
  fs.writeFileSync(stateFile, JSON.stringify({ filePath }, null, 2));
}

async function main() {
  const request = readInput();
  const command = request.command;
  const payload = request.payload || {};
  const selectedPath = readSelectedPath();

  if (selectedPath) {
    setSelectedExcelPath(selectedPath);
  }

  if (!commands[command]) {
    throw new Error(`Comando no soportado: ${command}`);
  }

  const data = command === 'selectFile'
    ? await commands.selectFile(payload.filePath)
    : await commands[command](payload);

  if ((command === 'selectFile' || command === 'setSelectedFile') && payload.filePath) {
    writeSelectedPath(payload.filePath);
  }

  process.stdout.write(JSON.stringify({ ok: true, data }));
}

main().catch((error) => {
  process.stdout.write(JSON.stringify({
    ok: false,
    error: error.message || String(error)
  }));
  process.exitCode = 1;
});
