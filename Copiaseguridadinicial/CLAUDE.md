# Gestor de Notas FP - Guía del Codebase

**Proyecto**: Gestor de Notas FP | **Versión**: 2.0.0 | **Stack**: Tauri v2.11.0 + Node.js + XLSX | **Estado**: Producción

## Estructura

```
plantillaNotas/
├── HTML: index.html, gestor-alumnos.html, gestor-rraa-criterios.html, 
│         gestor-unidades.html, gestor-notas.html, visor-notas.html, informes.html
├── Backend: main.js (Node IPC), app-bridge.js, tauri-node-backend.js
├── Scripts: prepare-tauri-web.js
├── src-tauri/: main.rs, Cargo.toml, tauri.conf.json
└── Documentacion/: README.md, TAURI_MIGRACION.md, OPTIMIZACIONES_RENDIMIENTO.md
```

## Stack

- **Frontend**: HTML5 + CSS + Vanilla JS (sin framework)
- **Desktop**: Tauri v2 (reemplazó Electron)
- **Backend**: Node.js con IPC
- **Excel**: XLSX + JSZip
- **Persistencia**: JSON local

## Comandos

```powershell
npm start                           # Dev Electron
node scripts/prepare-tauri-web.js && npm run tauri:dev    # Dev Tauri
npm run tauri:build                # Build EXE (src-tauri/target/release/bundle/nsis/)
```

## Datos Excel

| Hoja | Descripción |
|------|-------------|
| DATOS | Cols 0-2: alumno (Nº, Nombre, Nac); Cols 5-6: RRAA; Cols 8-11: Unidades; Cols 21-22: Criterios |
| PESOS | Fila 3: códigos criterios; Filas 5-20: ponderaciones (16 máx) |
| U1-U16 | Notas de actividades |
| Evaluaciones | 1ª EVA, 2ª EVA, 3ª EVA, FINAL |

## Optimizaciones (Mejora 6x)

| Problema | Solución | Resultado |
|----------|----------|-----------|
| Carga informes ~30s | Lazy load evaluaciones + Promise.all | ~5s (1ª eval) |
| Cambio alumno ~5s | Event delegation + DocumentFragment | Instantáneo |
| Renders notas ~500ms | requestAnimationFrame + WeakMap cache | ~50ms |

**Regla crítica**: No filtrar `sheets: []` en funciones que exploran `SheetNames`.

## IPC Handlers

**main.js / tauri-node-backend.js**:
- `commandSelectFile()`, `commandSetSelectedFile(filePath)`
- `commandLoadAlumnos()`, `commandLoadUnidades()`, `commandLoadRraa()`, `commandLoadNotasActividad()`, `commandLoadNotasEvaluacion()`
- `commandSaveAlumnos(data)`, `commandSaveUnidades(data)`, `commandSaveRraa(data)`, `commandSaveNotas(data)`

**Tauri Rust**: `excel_select_file`, `excel_set_selected_file`, `excel_get_selected_file`, `excel_invoke_command`

## Cambios Recientes (2026-05-08)

- **Menú recientes**: verificación de existencia, botones ✕ para eliminar, UI en hover
- **Fix borramiento**: validación XML + normalización cell refs (previene corrupción)
- **Agregar actividades**: nueva funcionalidad con campo nombre
- **Emoji UI**: reemplazo 📝 👁️ ➕ 📊 en pantalla inicio

## Checklist

- [x] Tauri migrado
- [x] Features funcionales
- [x] Optimizaciones aplicadas
- [x] Archivos recientes
- [x] Fix crítico (corrupción)
- [ ] Build EXE final

## Próximos Pasos

1. Commit de cambios (si aplica)
2. `npm run tauri:build`
3. Testear EXE

**Responsable**: Sebantonio | **Última actualización**: 2026-05-08
