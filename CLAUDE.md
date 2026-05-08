# Plantilla Notas - Guía del Codebase

## Resumen Ejecutivo

**Proyecto**: Gestor de Notas FP (Formación Profesional)  
**Versión**: 2.0.0  
**Stack**: Tauri v2.11.0 + Node.js + XLSX  
**Estado**: Producción  
**Última actualización**: 2026-05-08

## Estructura del Proyecto

```
plantillaNotas/
├── index.html                    # Panel principal + archivos recientes
├── gestor-alumnos.html          # Gestión de alumnos
├── gestor-rraa-criterios.html   # RRAA y criterios de evaluación
├── gestor-unidades.html         # Unidades de aprendizaje
├── gestor-notas.html            # Entrada de notas (tabla optimizada)
├── visor-notas.html             # Vista de notas por alumno
├── informes.html                # Reportes (lazy loading por evaluación)
│
├── main.js                       # Backend Node.js (IPC handlers)
├── app-bridge.js                # Bridge Tauri-Frontend
├── tauri-node-backend.js        # Backend Tauri (persistencia)
├── preload.js                   # Preload Electron (si aplica)
│
├── scripts/
│   └── prepare-tauri-web.js     # Copia HTML a tauri-web/
│
├── src-tauri/                   # Código Rust
│   ├── src/main.rs
│   ├── Cargo.toml
│   └── tauri.conf.json
│
└── Documentacion/
    ├── README.md
    ├── TAURI_MIGRACION.md
    ├── OPTIMIZACIONES_RENDIMIENTO.md
    └── arreglar_cuello.md
```

## Stack Técnico

- **Frontend**: HTML5 + CSS + Vanilla JavaScript (sin framework)
- **Desktop**: Tauri v2 (reemplazó Electron)
- **Backend**: Node.js con IPC (Tauri invoke / Electron ipcRenderer)
- **Excel I/O**: XLSX + JSZip
- **Persistencia**: Archivos JSON locales (.tauri-selected-excel.json, localStorage)

## Comandos de Desarrollo

```powershell
# Dev local (Electron, sin Rust)
npm start

# Dev con Tauri (requiere Rust)
node scripts/prepare-tauri-web.js
npm run tauri:dev

# Build EXE final
node scripts/prepare-tauri-web.js
npm run tauri:build
# Output: src-tauri/target/release/bundle/nsis/
```

## Arquitectura de Datos

### Archivo Excel Esperado

**Hoja DATOS**
- Cols 0-2: Nº, Nombre, Fecha Nac (alumnos)
- Cols 5-6: Nº RRAA, Descripción RRAA
- Cols 8-11: Código Unit, Nombre, Eval, Horas
- Cols 21-22: Código Criterio, Texto Criterio

**Hoja PESOS**
- Fila 3: Códigos de criterios
- Filas 5-20: Ponderaciones por unidad (16 unidades máx)

**Hojas de Unidad** (U1–U16)
- Contienen notas de actividades

**Hojas de Evaluación**
- "1ª EVA", "2ª EVA", "3ª EVA", "FINAL"

## Optimizaciones de Rendimiento

### Problema Original
- Carga de informes: ~30s (6 lecturas Excel secuenciales)
- Cambio de alumno: ~5s (5 renders sin batching)

### Soluciones Aplicadas

1. **Filtrado de hojas XLSX** (commit 26939c6)
   - Solo lee hojas necesarias: `XLSX.readFile(path, { sheets: ['DATOS', 'PESOS'] })`
   - ⚠️ No aplica a funciones que exploran SheetNames (rompe discovery)

2. **Event delegation en gestor-notas.html** (commit 26939c6)
   - 100 listeners → 1 listener delegado en `notasBody`
   - DocumentFragment para 100 reflows → 1 reflow

3. **Batching + Memoization en informes.html** (commit b6fc2bb)
   - `requestAnimationFrame` agrupa 5 renders en 1 frame
   - `WeakMap` cachea stats para no recalcular

4. **Promise.all + Lazy Loading** (commits 4088e4a, 2026-05-08)
   - Carga 4 evaluaciones en paralelo → 1 evaluación al arrancar
   - Carga bajo demanda cuando el usuario cambia dropdown

### Resultado
| Métrica | Original | Optimizado |
|---------|----------|-----------|
| Carga informes | ~30s | ~5s (1ª eval) |
| Cambio alumno | ~5s | instantáneo |
| Renderizado notas | ~500ms | ~50ms |

## Bug Conocido / Lección Aprendida

**No filtrar sheets en funciones que exploran SheetNames.**

Funciones como `loadNotasActividadFromSelectedFile()` llaman a `listUnitSheets(workbook)` que itera `workbook.SheetNames`. Si filtras sheets, SheetNames queda vacío y la búsqueda de unidades falla.

**Regla**: Solo filtrar `sheets: []` cuando conoces exactamente qué hoja necesitas ANTES de leer.

## IPC Handlers Principales

### main.js (Node.js)
```javascript
commandSelectFile()                    // Abre diálogo de archivo
commandSetSelectedFile(filePath)       // Establece archivo activo sin diálogo
commandLoadAlumnos()                   // Lee alumnos de DATOS
commandLoadUnidades()                  // Lee unidades
commandLoadRraa()                      // Lee RRAA + criterios
commandLoadNotasActividad()            // Lee notas de unidades
commandLoadNotasEvaluacion()          // Lee notas de evaluaciones
commandSaveAlumnos(data)              // Guarda alumnos a Excel
commandSaveUnidades(data)             // Guarda unidades
commandSaveRraa(data)                 // Guarda RRAA
commandSaveNotas(data)                // Guarda notas
```

### tauri-node-backend.js
Similar a main.js pero para Tauri. Persiste ruta en `.tauri-selected-excel.json`.

### Tauri Rust (src-tauri/src/main.rs)
```rust
excel_select_file           // Dialog + persist
excel_set_selected_file     // Set sin dialog
excel_get_selected_file     // Lee archivo activo
excel_invoke_command        // Router a tauri-node-backend.js
```

## Cambios Recientes (2026-05-08 - 2026-05-09 - En Compilación)

### Commits
- **f80c4d5**: feat mejorar menu recientes con verificación y eliminación (+162 -6)
  - Verificación automática de existencia de archivos
  - Botones ✕ para eliminar items individuales
  - Botón 🗑️ para limpiar historial completo
  - UI inteligente: botones en hover
  - Sincronización automática Electron/Tauri
- **3f9e8b6**: fix remover limpieza de celdas que causa corrupción Excel
  - Soluciona error de "Registros quitados"
  - Remueve intento de eliminar celdas (creaba XML inválido)
  - Enfoque más simple y seguro
- **bceb864**: ui mejorar pantalla de inicio con iconos emoji (+45 -16)
  - Reemplazar letras por emojis: 📝 👁️ ➕ 📊 (acciones)
  - Emojis en pasos: 👥 ⚖️ 📚 ✅
  - Animaciones hover mejoradas
  - Backgrounds de color suave por tipo
- **4712967**: fix copiar headers completos de memorias al añadir actividad (+7 -3) [REVERTIDO]
  - Soluciona error de reparación Excel
  - Headers de memorias/otros/controles se copian completamente
  - Diferencia entre filas de cabecera vs estudiantes
- **8aa340f**: reparado error que borraba hoja (+71 -6 main.js)
  - Fix crítico: validación XML + normalización cell refs en copia de actividades
  - Previene corrupción de archivos al guardar
- **9271e96**: incluida opcion añadir actividad
- **69ba757**: creado añadir actividad
- **2c80aae**: añadidonombre actividad
- **6b0954d**: optimizao informes

### Nueva Funcionalidad
- Opción de agregar actividades al sistema
- Campo nombre para actividades
- Dos fixes críticos que garantizan integridad de datos:
  1. Validación XML antes de guardar
  2. Copia completa de headers de otros tipos de actividad

## Checklist de Estado

- [x] Migración a Tauri completada
- [x] Build funcional
- [x] Todas las features funcionan
- [x] Optimizaciones de rendimiento aplicadas
- [x] Archivos recientes en index.html
- [x] Fix crítico de borramiento (8aa340f)
- [ ] Commit final de cambios del 2026-05-08
- [ ] Build EXE final con todas las optimizaciones

## Próximos Pasos

1. Hacer commit de los cambios (si aplica)
2. Compilar EXE: `npm run tauri:build`
3. Testear el ejecutable completo
4. Validar que las nuevas opciones de actividades funcionan correctamente

## Referencias

- Documentación de optimizaciones: `Documentacion/OPTIMIZACIONES_RENDIMIENTO.md`
- Plan de mejoras: `Documentacion/arreglar_cuello.md`
- Notas de migración Tauri: `Documentacion/TAURI_MIGRACION.md`

---

**Última actualización**: 2026-05-08  
**Responsable**: Sebantonio
