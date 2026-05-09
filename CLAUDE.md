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

## Estructura Hojas U1-U16 (crítico)

```
Fila 1 (idx 0): "U1", "1ª" — título
Fila 2 (idx 1): vacía / fusionada
Fila 3 (idx 2): "NOTA RA" en col F(5), T(19), AH(33)... — cabeceras RA
Fila 4 (idx 3): porcentajes CE (3%, 7%, 4%...)
Fila 5 (idx 4): códigos CE (1.a), 1.b)... 2.a), 2.b)...) — entre cols RA consecutivas
Fila 6+ (idx 5+): datos alumnos (col D=nombre, E=nota unidad, F=RA1, T=RA2...)
```

- Detección "NOTA RA": buscar en filas 0-3 celdas que empiecen por "NOTA RA" o "RA \d"
- Detección CE: buscar `is_criterion_code` en fila idx=4 (`first_row`), entre columnas RA consecutivas
- `first_row = 4` — el loop empieza aquí pero salta fila 5 (col D vacía) y lee desde fila 6

## Páginas HTML

| Archivo | Función |
|---------|---------|
| index.html | Inicio — menú principal |
| gestor-notas.html | Introducir notas actividades + CE por alumno |
| gestor-alumnos.html | Gestión de alumnos |
| gestor-rraa-criterios.html | Gestión de RA y criterios |
| gestor-unidades.html | Gestión de unidades |
| visor-notas.html | RRAA y CCEE (evaluaciones) |
| visor-actividades.html | Ver notas por actividad + panel RA |
| visor-unidades.html | Ver notas por unidad + desplegable RA/CE por alumno |
| informes.html | Informes finales |

## Cambios Recientes (2026-05-08/09)

- **Menú recientes**: verificación de existencia, botones ✕ para eliminar, UI en hover
- **Fix borramiento**: validación XML + normalización cell refs (previene corrupción)
- **Agregar actividades**: nueva funcionalidad con campo nombre
- **Emoji UI**: reemplazo 📝 👁️ ➕ 📊 en pantalla inicio
- **CE por alumno (gestor-notas)**: desplegable per-alumno con selector RA + inputs CE editables; flag `_ceModified` para no escribir todos; Rust guarda ceNotas por alumno en `excel_save_notas_actividad`
- **Panel RA (visor-actividades)**: botón ▼ RA por columna actividad → panel con media CE por RA
- **Panel RA/CE (visor-unidades)**: botón ▼ RA por alumno → fila expandible con notas RA; botón ▼ CE por RA → notas CE de ese RA. Rust en `load_notas_unidad` detecta cols RA ("NOTA RA") y CE (fila idx=4)
- **Restyling gestor-notas**: eliminado gradiente morado, ahora coincide con paleta indigo del resto (#6366f1, fondo #f3f6fb)

## IPC Handlers

**app-bridge.js → Rust**:
- `excel_select_file`, `excel_set_selected_file`, `excel_get_selected_file`
- `excel_get_alumnos`, `excel_save_alumnos`
- `excel_get_unidades`, `excel_save_unidades`
- `excel_get_rraa_criterios`, `excel_save_rraa_criterios`
- `excel_get_notas_actividad`, `excel_save_notas_actividad` (incluye ceNotas por alumno)
- `excel_save_ce_notas` (mismo CE para todos — uso batch)
- `excel_add_actividad`
- `excel_get_notas_actividades_tipo`
- `excel_get_notas_evaluacion`, `excel_get_notas_evaluacion_alumno`
- `excel_get_notas_unidad` (devuelve raColumnas + ceNotas por RA por alumno)
- `excel_get_alumnos_informes`

## Checklist

- [x] Tauri migrado
- [x] Features funcionales
- [x] Optimizaciones aplicadas
- [x] Archivos recientes
- [x] Fix crítico (corrupción)
- [x] CE por alumno en gestor-notas
- [x] Paneles RA/CE en visor-actividades y visor-unidades
- [x] Restyling gestor-notas
- [ ] Build EXE final

## Próximos Pasos

1. `npm run tauri:build`
2. Testear EXE

**Responsable**: Sebantonio | **Última actualización**: 2026-05-09
