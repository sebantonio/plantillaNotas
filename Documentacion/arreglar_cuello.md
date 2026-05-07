# Plan para resolver el cuello de botella de informes.html

**Fecha**: 2026-05-07  
**Problema**: `informes.html` tarda 5-7 segundos al cargar (antes eran 30s, ya se mejoró con Promise.all y filtrado de hojas)  
**Estado actual**: commit `b03ea75` — optimizaciones aplicadas pero sin caché de estado entre llamadas IPC

---

## El problema raíz (confirmado por investigación)

**No es SheetJS el lento. Es la arquitectura del sidecar.**

Cada vez que el frontend llama a `invoke("excel_get_notas_evaluacion")`, Tauri ejecuta:
```
spawn proceso Node → cold start (~300ms) → XLSX.readFile() completo → devuelve datos → proceso muere
```

Esto ocurre 4 veces en paralelo (4 evaluaciones). Aunque van en paralelo, cada una relanza Node y relee el Excel entero. No hay memoria entre llamadas.

---

## Las 3 soluciones (ordenadas por facilidad)

---

### OPCIÓN A — Opciones de parseo SheetJS ⚡ (30 min, riesgo cero)

Sin tocar la arquitectura, añadir estas opciones al `XLSX.readFile()` en `main.js`:

```js
// ANTES (actual)
XLSX.readFile(path, { cellDates: true })

// DESPUÉS
XLSX.readFile(path, {
  cellDates: true,
  dense: true,       // arrays en lugar de objetos — más eficiente en V8 moderno
  cellFormula: false, // no parsear fórmulas (no las necesitamos)
  cellHTML: false,    // no convertir a HTML rico (no lo usamos)
  cellNF: false,      // no parsear formatos numéricos
})
```

**Por qué funciona:**
- `dense: true` — la documentación SheetJS dice que "V8 evolucionó para que array-of-arrays sea más eficiente que objetos grandes". Node.js ≥ 18 usa V8 moderno.
- `cellFormula: false` + `cellHTML: false` — el Excel tiene fórmulas y formato. SheetJS los parsea por defecto aunque no los usemos.
- `cellNF: false` — evita parsear los formatos numéricos de celda.

**Funciones a modificar en main.js** (las que NO filtran hojas, porque sí necesitan leer todo):
- `loadNotasActividadFromSelectedFile` → línea ~574
- `loadNotasEvaluacionFromSelectedFile` → línea ~620

**Estimación**: -30% a -70% en tiempo de parseo según la complejidad del Excel.

**Riesgo**: Cero. Si algo falla, quitar las opciones y volver al estado actual.

---

### OPCIÓN B — Sidecar Node persistente 🔧 (1-2 días, riesgo medio)

**Idea:** En lugar de relanzar Node por cada llamada, arrancar un servidor HTTP local al inicio de la app. El workbook se parsea una vez y queda en memoria.

**Cómo funciona:**
```
App arranca → Node servidor HTTP en puerto aleatorio (ej: 48392)
Frontend → invoke("excel_get_sheet", {sheet: "1ª EVA"})
Tauri Rust → HTTP GET http://127.0.0.1:48392/sheet/1ª EVA
Servidor Node → workbook ya en memoria → responde en <10ms
```

**Implementación posible con `tauri-plugin-js`:**
- GitHub: `HuakunShen/tauri-plugin-js`
- Plugin no oficial pero activo (2025). Gestiona el ciclo de vida del proceso JS desde Rust.
- Usa `kkrpc` para RPC tipado sobre stdio — sin overhead de red.
- El proceso JS es long-lived por diseño.

**Implementación alternativa con Express:**
1. Compilar un servidor Express con `pkg` o `bun build --compile` → binario standalone
2. Tauri lo arranca como sidecar al inicio y guarda el PID en `AppState`
3. Cuando la app cierra, Tauri mata el proceso
4. El servidor Node mantiene el workbook en un `Map` en memoria

**Cache en el servidor Node:**
```js
let _wb = null;
let _wbPath = null;

app.get('/sheet/:name', (req, res) => {
  const filePath = req.query.path;
  if (_wbPath !== filePath) {
    _wb = XLSX.readFile(filePath, { dense: true, cellHTML: false, cellFormula: false });
    _wbPath = filePath;
  }
  const sheet = _wb.Sheets[req.params.name];
  res.json(XLSX.utils.sheet_to_json(sheet, { header: 1 }));
});
```

**Primera carga**: 5-7s igual (hay que leer el Excel)  
**Segunda carga del mismo Excel**: ~0ms (ya en memoria)  
**Cambio de alumno en informes**: instantáneo

**Tamaño estimado del proceso Node**: ~50-100MB RAM.

---

### OPCIÓN C — Calamine en Rust 🦀 (3-5 días, requiere Rust)

**Idea:** Eliminar Node completamente para la lectura de Excel. Usar la librería Rust `calamine` directamente en el backend Tauri. Los datos quedan en `State<Mutex<>>` entre llamadas IPC.

**Por qué es la opción óptima a largo plazo:**
- Sin proceso Node separado → sin cold start
- Cache nativa entre llamadas (el State de Tauri persiste toda la sesión)
- Calamine soporta lazy loading de hojas concretas en XLSX
- Benchmarks: ~1.1 millones de celdas/segundo, ~7x más rápido que C#/Python equivalentes
- El proceso Rust ya está activo — no hay que arrancar nada más

**Implementación:**
```toml
# Cargo.toml
[dependencies]
calamine = "0.34"
serde = { version = "1", features = ["derive"] }
```

```rust
use calamine::{open_workbook, Xlsx, Reader};
use std::collections::HashMap;
use std::sync::Mutex;

struct ExcelCache {
    sheets: HashMap<String, Vec<Vec<serde_json::Value>>>,
    path: String,
}

#[tauri::command]
fn get_sheet_data(
    path: String,
    sheet: String,
    state: tauri::State<Mutex<ExcelCache>>,
) -> Result<Vec<Vec<serde_json::Value>>, String> {
    let mut cache = state.lock().unwrap();

    if cache.path != path {
        cache.sheets.clear();
        cache.path = path.clone();
    }

    if let Some(data) = cache.sheets.get(&sheet) {
        return Ok(data.clone()); // hit de caché
    }

    let mut workbook: Xlsx<_> = open_workbook(&path)
        .map_err(|e| e.to_string())?;

    let range = workbook.worksheet_range(&sheet)
        .map_err(|e| e.to_string())?;

    let data: Vec<Vec<serde_json::Value>> = range.rows()
        .map(|row| row.iter().map(|cell| /* convertir celda a JSON */ ).collect())
        .collect();

    cache.sheets.insert(sheet, data.clone());
    Ok(data)
}
```

**Limitación crítica:** Calamine es **solo lectura**. Para guardar datos en el Excel, seguiría siendo necesario SheetJS/Node. Habría que mantener el sidecar solo para escritura.

---

## Comparativa rápida

| Opción | 1ª carga | 2ª carga | Complejidad | Riesgo |
|--------|----------|----------|-------------|--------|
| A — Opciones SheetJS | ~3-5s (-40%) | ~3-5s | Muy baja | Nulo |
| B — Sidecar persistente | ~5-7s (igual) | ~0ms | Media | Bajo |
| C — Calamine Rust | <1s estimado | ~0ms | Alta | Bajo |

---

## Recomendación para mañana

**Paso 1 (15 min):** Aplicar Opción A en `main.js`. Medir si hay mejora apreciable.

**Paso 2 (según resultado):** Si con la Opción A el tiempo baja a ~2-3s, puede ser suficiente. Si sigue siendo lento, valorar la Opción B.

**La Opción B (sidecar persistente)** es la que resuelve el problema de raíz sin reescribir en Rust. El punto de entrada más claro es arrancar un servidor Express simple desde Tauri.

---

## Archivos relevantes

| Archivo | Líneas clave |
|---------|-------------|
| `main.js` | ~574 `loadNotasActividadFromSelectedFile` |
| `main.js` | ~620 `loadNotasEvaluacionFromSelectedFile` |
| `src-tauri/src/main.rs` | Todo el archivo — aquí iría calamine si se implementa |
| `tauri-node-backend.js` | Aquí iría la lógica del servidor persistente |

---

## Fuentes consultadas

- [SheetJS Parse Options — documentación oficial](https://docs.sheetjs.com/docs/api/parse-options/)
- [GitHub SheetJS Issue #409 — bookSheets reduce parse time 60%](https://github.com/SheetJS/sheetjs/issues/409)
- [Tauri v2 — Node.js as sidecar](https://v2.tauri.app/learn/sidecar-nodejs/)
- [Tauri v2 — State Management](https://v2.tauri.app/develop/state-management/)
- [tauri-plugin-js — GitHub HuakunShen](https://github.com/HuakunShen/tauri-plugin-js)
- [calamine — GitHub tafia/calamine](https://github.com/tafia/calamine)
- [PkgPulse benchmark SheetJS vs ExcelJS 2026](https://www.pkgpulse.com/blog/sheetjs-vs-exceljs-vs-node-xlsx-excel-files-node-2026)
- [Dev.to — Adding Node.js server to Tauri as sidecar](https://dev.to/zaid_sunasra/adding-nodejs-server-to-tauri-app-as-a-sidecar-509j)
