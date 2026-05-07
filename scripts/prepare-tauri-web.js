const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const outDir = path.join(root, 'tauri-web');
const files = [
  'index.html',
  'gestor-alumnos.html',
  'gestor-rraa-criterios.html',
  'gestor-unidades.html',
  'gestor-notas.html',
  'visor-notas.html',
  'informes.html',
  'app-bridge.js'
];

fs.rmSync(outDir, { recursive: true, force: true });
fs.mkdirSync(outDir, { recursive: true });

files.forEach((file) => {
  fs.copyFileSync(path.join(root, file), path.join(outDir, file));
});

console.log(`Preparados ${files.length} archivos para Tauri en ${outDir}`);
