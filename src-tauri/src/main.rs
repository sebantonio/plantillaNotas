#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use once_cell::sync::Lazy;
use regex::Regex;
use serde_json::{json, Value};
use std::collections::HashMap;
use std::io::{Read as IoRead, Write};
use std::path::Path;
use std::sync::Mutex;

// ---------------------------------------------------------------------------
// Estado global: ruta del Excel seleccionado
// ---------------------------------------------------------------------------

static SELECTED_PATH: Lazy<Mutex<Option<String>>> = Lazy::new(|| Mutex::new(None));

fn get_selected_path() -> Option<String> {
    SELECTED_PATH.lock().unwrap().clone()
}

fn set_selected_path(p: Option<String>) {
    *SELECTED_PATH.lock().unwrap() = p;
}

fn find_default_excel_path() -> Option<String> {
    if let Ok(exe) = std::env::current_exe() {
        if let Some(dir) = exe.parent() {
            let candidate = dir.join("plantilla313_dual - copia.xlsx");
            if candidate.exists() {
                return Some(candidate.to_string_lossy().to_string());
            }
        }
    }
    None
}

fn require_selected_path() -> Result<String, String> {
    match get_selected_path().or_else(find_default_excel_path) {
        Some(p) => Ok(p),
        None => Err("No hay ningun archivo Excel seleccionado.".to_string()),
    }
}

// ---------------------------------------------------------------------------
// Utilidades de columnas Excel
// ---------------------------------------------------------------------------

fn col_name(idx: usize) -> String {
    let mut name = String::new();
    let mut v = idx + 1;
    while v > 0 {
        let r = (v - 1) % 26;
        name.insert(0, (b'A' + r as u8) as char);
        v = (v - 1) / 26;
    }
    name
}

fn col_index(name: &str) -> usize {
    name.chars()
        .fold(0usize, |acc, c| acc * 26 + (c as usize - 'A' as usize + 1))
        - 1
}

fn normalize_plain(s: &str) -> String {
    s.chars()
        .map(|c| match c {
            'á' | 'à' | 'â' | 'ä' | 'Á' | 'À' | 'Â' | 'Ä' => 'A',
            'é' | 'è' | 'ê' | 'ë' | 'É' | 'È' | 'Ê' | 'Ë' => 'E',
            'í' | 'ì' | 'î' | 'ï' | 'Í' | 'Ì' | 'Î' | 'Ï' => 'I',
            'ó' | 'ò' | 'ô' | 'ö' | 'Ó' | 'Ò' | 'Ô' | 'Ö' => 'O',
            'ú' | 'ù' | 'û' | 'ü' | 'Ú' | 'Ù' | 'Û' | 'Ü' => 'U',
            'ñ' | 'Ñ' => 'N',
            other => other.to_ascii_uppercase(),
        })
        .collect::<String>()
        .trim()
        .to_string()
}

fn parse_decimal(v: &Value) -> f64 {
    match v {
        Value::Number(n) => n.as_f64().unwrap_or(0.0),
        Value::String(s) => s.replace(',', ".").parse::<f64>().unwrap_or(0.0),
        _ => 0.0,
    }
}

// ---------------------------------------------------------------------------
// Leer XLSX con calamine
// ---------------------------------------------------------------------------

use calamine::{open_workbook_auto, Data, Reader};

fn cell_data_to_value(cell: &Data) -> Value {
    match cell {
        Data::Empty => Value::Null,
        Data::String(s) => Value::String(s.clone()),
        Data::Float(f) => json!(f),
        Data::Int(i) => json!(i),
        Data::Bool(b) => Value::Bool(*b),
        Data::DateTime(f) => json!(f.as_f64()),
        Data::Error(_) => Value::Null,
        _ => Value::Null,
    }
}

fn read_sheet_rows(path: &str, sheet: &str) -> Result<Vec<Vec<Value>>, String> {
    let mut wb = open_workbook_auto(path).map_err(|e| e.to_string())?;
    let range = wb
        .worksheet_range(sheet)
        .map_err(|e| format!("Hoja '{sheet}' no encontrada: {e}"))?;

    let (row_start, col_start) = range.start().map(|(r, c)| (r as usize, c as usize)).unwrap_or((0, 0));
    let (row_end, _col_end) = range.end().map(|(r, c)| (r as usize, c as usize)).unwrap_or((0, 0));

    // Build a dense array where result[row_idx] corresponds to Excel row (row_idx + 1).
    // Calamine's range.rows() skips rows that have no cells in the XML, so we must
    // use get_value(r, c) to preserve the Excel row-to-index correspondence.
    let total_rows = if row_end >= row_start { row_end - row_start + 1 } else { 0 };
    let total_cols = range.width();

    // result[i] must correspond to Excel row (i+1), so pad with row_start empty rows first.
    let mut result: Vec<Vec<Value>> = vec![vec![]; row_start];
    for r in 0..total_rows {
        let abs_row = (row_start + r) as u32;
        let row: Vec<Value> = (0..total_cols)
            .map(|c| {
                let abs_col = (col_start + c) as u32;
                range.get_value((abs_row, abs_col))
                    .map(cell_data_to_value)
                    .unwrap_or(Value::Null)
            })
            .collect();
        // Prepend col_start empty cells so column indices match Excel column indices
        let mut full_row = vec![Value::Null; col_start];
        full_row.extend(row);
        result.push(full_row);
    }
    Ok(result)
}

fn sheet_names(path: &str) -> Result<Vec<String>, String> {
    let wb = open_workbook_auto(path).map_err(|e| e.to_string())?;
    Ok(wb.sheet_names().to_vec())
}

fn cell_val_str(v: &Value) -> String {
    match v {
        Value::String(s) => s.clone(),
        Value::Number(n) => {
            if let Some(i) = n.as_i64() { i.to_string() } else { n.as_f64().unwrap_or(0.0).to_string() }
        }
        Value::Bool(b) => b.to_string(),
        _ => String::new(),
    }
}

fn cell_str(rows: &[Vec<Value>], row: usize, col: usize) -> String {
    rows.get(row)
        .and_then(|r| r.get(col))
        .map(|v| cell_val_str(v).trim().to_string())
        .unwrap_or_default()
}

fn cell_f64(rows: &[Vec<Value>], row: usize, col: usize) -> Option<f64> {
    rows.get(row).and_then(|r| r.get(col)).and_then(|v| match v {
        Value::Number(n) => n.as_f64(),
        Value::String(s) => s.replace(',', ".").parse().ok(),
        _ => None,
    })
}

fn normalise_date_for_ui(rows: &[Vec<Value>], row: usize, col: usize) -> String {
    let v = rows.get(row).and_then(|r| r.get(col)).cloned().unwrap_or(Value::Null);
    match &v {
        Value::String(s) if !s.is_empty() => s.clone(),
        Value::Number(n) => {
            if let Some(serial) = n.as_f64() {
                let days = serial as i64 - 25569;
                let secs = days * 86400;
                chrono::DateTime::from_timestamp(secs, 0)
                    .map(|dt: chrono::DateTime<chrono::Utc>| dt.format("%Y-%m-%d").to_string())
                    .unwrap_or_default()
            } else {
                String::new()
            }
        }
        _ => String::new(),
    }
}

// ---------------------------------------------------------------------------
// Buscar filas de secciones en hoja DATOS
// ---------------------------------------------------------------------------

fn find_header_row(rows: &[Vec<Value>], text: &str) -> Option<usize> {
    let target = text.to_uppercase();
    rows.iter().position(|row| {
        row.get(1)
            .map(|v| cell_val_str(v).to_uppercase().contains(&target))
            .unwrap_or(false)
    })
}

fn find_unidades_start(rows: &[Vec<Value>]) -> Option<usize> {
    let idx = rows.iter().position(|row| {
        let col8 = row.get(8).map(|v| !matches!(v, Value::Null)).unwrap_or(false);
        let col9 = row.get(9).map(|v| cell_val_str(v).to_uppercase().contains("UNIDADES")).unwrap_or(false);
        col8 && col9
    })?;
    Some(idx + 1)
}

fn find_rraa_start(rows: &[Vec<Value>]) -> Option<usize> {
    let idx = rows.iter().position(|row| {
        row.get(6).map(|v| cell_val_str(v).to_uppercase().contains("RRAA")).unwrap_or(false)
    })?;
    Some(idx + 1)
}

// ---------------------------------------------------------------------------
// Comandos: selectFile / getSelectedFile / setSelectedFile
// ---------------------------------------------------------------------------

#[tauri::command]
fn excel_select_file() -> Result<Value, String> {
    let file = rfd::FileDialog::new()
        .set_title("Selecciona la plantilla Excel")
        .add_filter("Excel", &["xlsx", "xlsm", "xls"])
        .pick_file();
    match file {
        Some(path) => {
            let p = path.to_string_lossy().to_string();
            set_selected_path(Some(p.clone()));
            load_alumnos(&p)
        }
        None => Ok(Value::Null),
    }
}

#[tauri::command]
fn excel_get_selected_file() -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    match get_selected_path() {
        Some(p) => load_alumnos(&p),
        None => Ok(Value::Null),
    }
}

#[tauri::command]
fn excel_set_selected_file(file_path: String) -> Result<Value, String> {
    if file_path.is_empty() { return Err("No se especificó ningún archivo.".to_string()); }
    if !Path::new(&file_path).exists() { return Err(format!("El archivo no existe: {file_path}")); }
    let name = Path::new(&file_path).file_name().unwrap_or_default().to_string_lossy().to_string();
    set_selected_path(Some(file_path.clone()));
    Ok(json!({ "filePath": file_path, "fileName": name }))
}

#[tauri::command]
fn excel_verify_file_exists(file_path: String) -> Result<bool, String> {
    Ok(Path::new(&file_path).exists())
}

// ---------------------------------------------------------------------------
// Cargar alumnos
// ---------------------------------------------------------------------------

fn load_alumnos(path: &str) -> Result<Value, String> {
    let rows = read_sheet_rows(path, "DATOS")?;
    let header_row = find_header_row(&rows, "ALUMNADO")
        .ok_or("No se encontro la seccion ALUMNADO en la hoja DATOS.")?;
    let mut alumnos = Vec::new();
    for i in (header_row + 1)..rows.len() {
        let nombre = cell_str(&rows, i, 1);
        if nombre.is_empty() { break; }
        let numero = cell_f64(&rows, i, 0).map(|f| json!(f as i64)).unwrap_or(json!(alumnos.len() + 1));
        alumnos.push(json!({ "numero": numero, "nombre": nombre, "fechaNac": normalise_date_for_ui(&rows, i, 2) }));
    }
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({ "filePath": path, "fileName": file_name, "alumnos": alumnos }))
}

#[tauri::command]
fn excel_save_alumnos(alumnos: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let arr = alumnos.as_array().ok_or("La lista de alumnos no es valida.")?.clone();
    save_alumnos_to_file(&path, &arr)?;
    load_alumnos(&path)
}

// ---------------------------------------------------------------------------
// Cargar unidades
// ---------------------------------------------------------------------------

fn load_unidades(path: &str) -> Result<Value, String> {
    let rows = read_sheet_rows(path, "DATOS")?;
    let start = find_unidades_start(&rows).ok_or("No se encontro la seccion UNIDADES.")?;
    let mut unidades = Vec::new();
    for idx in 0..16 {
        let ri = start + idx;
        let codigo = { let s = cell_str(&rows, ri, 8); if s.is_empty() { format!("U{}", idx + 1) } else { s } };
        let nombre = cell_str(&rows, ri, 9);
        let evaluacion = cell_str(&rows, ri, 10);
        let horas = cell_str(&rows, ri, 11);
        unidades.push(json!({ "codigo": codigo, "nombre": nombre, "evaluacion": evaluacion, "horas": horas }));
    }
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({ "filePath": path, "fileName": file_name, "unidades": unidades }))
}

#[tauri::command]
fn excel_get_unidades() -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    match get_selected_path() { Some(p) => load_unidades(&p), None => Ok(Value::Null) }
}

#[tauri::command]
fn excel_save_unidades(unidades: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let arr = unidades.as_array().ok_or("La lista de unidades no es valida.")?.clone();
    save_unidades_to_file(&path, &arr)?;
    load_unidades(&path)
}

// ---------------------------------------------------------------------------
// Cargar RRAA y criterios
// ---------------------------------------------------------------------------

fn is_criterion_code(s: &str) -> bool {
    Regex::new(r"^\d+\.[a-z]\)$").unwrap().is_match(&s.trim().to_lowercase())
}

fn normalize_criterion_code(s: &str) -> String {
    s.trim().trim_end_matches(')').to_lowercase().to_string()
}

fn load_rraa_criterios(path: &str) -> Result<Value, String> {
    let (rraa, criterios, ponderaciones_unidad) = extract_rraa_criterios_data(path)
        .ok_or("No se encontro la seccion RRAA o las hojas DATOS/PESOS.")?;
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({ "filePath": path, "fileName": file_name, "rraa": rraa, "criterios": criterios, "ponderacionesUnidad": ponderaciones_unidad }))
}

#[tauri::command]
fn excel_get_rraa_criterios() -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    match get_selected_path() { Some(p) => load_rraa_criterios(&p), None => Ok(Value::Null) }
}

#[tauri::command]
fn excel_save_rraa_criterios(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let rraa = payload["rraa"].as_array().cloned().unwrap_or_default();
    let criterios = payload["criterios"].as_array().cloned().unwrap_or_default();
    let pond_unidad = payload["ponderacionesUnidad"].as_array().cloned().unwrap_or_default();
    save_rraa_criterios_to_file(&path, &rraa, &criterios, &pond_unidad)?;
    load_rraa_criterios(&path)
}

// Extrae datos de RRAA/criterios sin metadata de fichero; devuelve (rraa, criterios, ponderaciones_unidad)
fn extract_rraa_criterios_data(path: &str) -> Option<(Vec<Value>, Vec<Value>, Vec<Value>)> {
    let datos_rows = read_sheet_rows(path, "DATOS").ok()?;
    let pesos_rows = read_sheet_rows(path, "PESOS").ok()?;
    let rraa_start = find_rraa_start(&datos_rows)?;

    let mut rraa: Vec<Value> = Vec::new();
    for i in rraa_start..datos_rows.len() {
        let numero = cell_str(&datos_rows, i, 5);
        let desc = cell_str(&datos_rows, i, 6);
        if numero.is_empty() && desc.is_empty() { break; }
        rraa.push(json!({
            "numero": if numero.is_empty() { json!(rraa.len() + 1) } else { numero.parse::<i64>().map(|n| json!(n)).unwrap_or(json!(numero)) },
            "descripcion": desc
        }));
    }

    let mut criterion_texts: HashMap<String, String> = HashMap::new();
    for row in &datos_rows {
        let code = row.get(21).map(cell_val_str).unwrap_or_default();
        let text = row.get(22).map(cell_val_str).unwrap_or_default();
        let with_paren = if code.ends_with(')') { code.clone() } else { format!("{code})") };
        if is_criterion_code(&with_paren) && !text.is_empty() {
            criterion_texts.insert(normalize_criterion_code(&code), text);
        }
    }

    let mut criterios: Vec<Value> = Vec::new();
    let pesos_row3 = pesos_rows.get(3).cloned().unwrap_or_default();
    for (col_idx, cell) in pesos_row3.iter().enumerate() {
        let codigo_raw = cell_val_str(cell);
        let with_paren = if codigo_raw.ends_with(')') { codigo_raw.clone() } else { format!("{codigo_raw})") };
        if !is_criterion_code(&with_paren) { continue; }
        let ra_num: i64 = codigo_raw.chars().take_while(|c| c.is_ascii_digit()).collect::<String>().parse().unwrap_or(0);
        let ra_desc = rraa.iter().find(|r: &&Value| r["numero"].as_i64() == Some(ra_num))
            .and_then(|r| r["descripcion"].as_str()).unwrap_or("").to_string();
        let ponder = cell_f64(&pesos_rows, 21, col_idx).unwrap_or(0.0);
        let ponder_inst = cell_f64(&pesos_rows, 21, col_idx + 1).unwrap_or(0.0);
        let ponder_emp = cell_f64(&pesos_rows, 21, col_idx + 2).unwrap_or(0.0);
        let texto = criterion_texts.get(&normalize_criterion_code(&codigo_raw)).cloned().unwrap_or_default();
        let codigo_full = if codigo_raw.ends_with(')') { codigo_raw.clone() } else { format!("{codigo_raw})") };
        criterios.push(json!({
            "numero": criterios.len() + 1, "codigo": codigo_full, "nombre": codigo_full,
            "originalCodigo": codigo_full, "raNumero": ra_num, "raDescripcion": ra_desc,
            "ponderacion": ponder, "ponderacionInstituto": ponder_inst, "ponderacionEmpresa": ponder_emp,
            "texto": texto, "colIdx": col_idx
        }));
    }

    let mut ponderaciones_unidad: Vec<Value> = Vec::new();
    for row_idx in 5..=20usize {
        let nombre = { let s = cell_str(&pesos_rows, row_idx, 0); if s == "0" { String::new() } else { s } };
        let mut ponderaciones = serde_json::Map::new();
        for c in &criterios {
            let ci = c["colIdx"].as_u64().unwrap_or(0) as usize;
            ponderaciones.insert(ci.to_string(), json!({
                "ponderacion": cell_f64(&pesos_rows, row_idx, ci).unwrap_or(0.0),
                "ponderacionInstituto": cell_f64(&pesos_rows, row_idx, ci + 1).unwrap_or(0.0),
                "ponderacionEmpresa": cell_f64(&pesos_rows, row_idx, ci + 2).unwrap_or(0.0)
            }));
        }
        ponderaciones_unidad.push(json!({ "numero": row_idx - 4, "rowIdx": row_idx, "nombre": nombre, "ponderaciones": ponderaciones }));
    }

    Some((rraa, criterios, ponderaciones_unidad))
}

// ---------------------------------------------------------------------------
// Tipos de actividad y bloques
// ---------------------------------------------------------------------------

struct ActivityType { key: &'static str, label: &'static str, base_col: usize }

const ACTIVITY_TYPES: &[ActivityType] = &[
    ActivityType { key: "practicas",  label: "Practicas",                base_col: 0   },
    ActivityType { key: "memorias",   label: "Memorias",                 base_col: 112 },
    ActivityType { key: "otros",      label: "Otras actividades",        base_col: 223 },
    ActivityType { key: "controles",  label: "Control teorico/practico", base_col: 334 },
];

struct ActivityFixedCols { start: usize, end: usize }

fn activity_fixed_cols(key: &str) -> ActivityFixedCols {
    match key {
        "practicas"  => ActivityFixedCols { start: 0,   end: 109 },
        "memorias"   => ActivityFixedCols { start: 112, end: 220 },
        "otros"      => ActivityFixedCols { start: 223, end: 331 },
        "controles"  => ActivityFixedCols { start: 334, end: 442 },
        _            => ActivityFixedCols { start: 0,   end: 109 },
    }
}

fn get_activity_type(key: &str) -> &'static ActivityType {
    ACTIVITY_TYPES.iter().find(|t| t.key == key).unwrap_or(&ACTIVITY_TYPES[0])
}

#[derive(Clone, Debug)]
#[allow(dead_code)]
struct ActivityBlock {
    tipo: String, tipo_label: String, numero: i64,
    title_row: usize, number_row: usize, included_row: usize,
    header_row: usize, first_student_row: usize,
    name_col: usize, note_col: usize, number_col: usize,
    name_value_col: usize, included_col: usize,
    nombre: String, incluida: bool,
    ce_cols: Vec<(String, usize)>,
}

fn find_activity_header_col(row: &[Value], start_col: usize, text: &str) -> Option<usize> {
    let target = normalize_plain(text);
    (start_col..(start_col + 8)).find(|&ci| {
        row.get(ci).map(|v| normalize_plain(&cell_val_str(v)).contains(&target)).unwrap_or(false)
    })
}

fn resolve_name_value_col(row: &[Value], start_col: usize) -> usize {
    find_activity_header_col(row, start_col, "NOMBRE").map(|c| c + 1).unwrap_or(start_col + 3)
}

fn find_activity_blocks(rows: &[Vec<Value>], tipo_key: &str) -> Vec<ActivityBlock> {
    let at = get_activity_type(tipo_key);
    let mut blocks = Vec::new();
    for row_idx in 0..rows.len().saturating_sub(4) {
        let number_cell = rows.get(row_idx + 1).and_then(|r| r.get(at.base_col + 1)).cloned().unwrap_or(Value::Null);
        let activity_number: i64 = match &number_cell {
            Value::Number(n) => match n.as_i64().or_else(|| n.as_f64().map(|f| f as i64)) { Some(v) => v, None => continue },
            Value::String(s) => match s.trim().parse::<f64>() { Ok(f) if f > 0.0 => f as i64, _ => continue },
            _ => continue,
        };
        let header_row_data = rows.get(row_idx + 3).cloned().unwrap_or_default();
        let name_col = match find_activity_header_col(&header_row_data, at.base_col, "NOMBRE Y APELLIDOS") { Some(c) => c, None => continue };
        let note_col = match find_activity_header_col(&header_row_data, at.base_col, "NOTA FINAL") { Some(c) => c, None => continue };
        let row1 = rows.get(row_idx + 1).cloned().unwrap_or_default();
        let name_value_col = resolve_name_value_col(&row1, at.base_col);
        let nombre = cell_val_str(row1.get(name_value_col).unwrap_or(&Value::Null)).trim().to_string();
        // CE codes are in row_idx+2 (the INCLUIDO/FECHA row), not the header row
        let ce_code_row = rows.get(row_idx + 2).cloned().unwrap_or_default();
        let included_val = ce_code_row.get(at.base_col + 1).map(cell_val_str).unwrap_or_default();
        let fixed = activity_fixed_cols(at.key);
        let mut ce_cols = Vec::new();
        for ci in (note_col + 1)..=fixed.end {
            if let Some(v) = ce_code_row.get(ci) {
                let s = cell_val_str(v);
                let strim = s.trim().to_lowercase();
                if strim.is_empty() { continue; }
                let code = if strim.ends_with(')') { strim.clone() } else { format!("{strim})") };
                if is_criterion_code(&code) { ce_cols.push((code, ci)); }
            }
        }
        blocks.push(ActivityBlock {
            tipo: at.key.to_string(), tipo_label: at.label.to_string(), numero: activity_number,
            title_row: row_idx, number_row: row_idx + 1, included_row: row_idx + 2,
            header_row: row_idx + 3, first_student_row: row_idx + 4,
            name_col, note_col, number_col: at.base_col + 1, name_value_col,
            included_col: at.base_col + 1, nombre, incluida: normalize_plain(&included_val) == "X",
            ce_cols,
        });
    }
    blocks
}

fn format_activity_block(b: &ActivityBlock) -> Value {
    let label = if b.nombre.is_empty() { format!("{} {}", b.tipo_label, b.numero) } else { format!("{} {} - {}", b.tipo_label, b.numero, b.nombre) };
    json!({ "numero": b.numero, "nombre": b.nombre, "incluida": b.incluida, "label": label, "firstStudentRow": b.first_student_row, "noteCol": b.note_col })
}

fn extract_activity_notes(rows: &[Vec<Value>], block: &ActivityBlock, max_alumnos: Option<usize>) -> Vec<Value> {
    let mut notes = Vec::new();
    for row_idx in block.first_student_row..rows.len() {
        let nombre = cell_str(rows, row_idx, block.name_col);
        if nombre.is_empty() || nombre == "0" { break; }
        let nota_raw = rows.get(row_idx).and_then(|r| r.get(block.note_col)).cloned().unwrap_or(Value::Null);
        let nota_str = match &nota_raw {
            Value::Number(n) => n.as_f64().map(|f| f.to_string().replace('.', ",")),
            Value::String(s) => Some(s.replace('.', ",")),
            _ => None,
        }.unwrap_or_default();
        let mut ce_notas = serde_json::Map::new();
        for (code, ci) in &block.ce_cols {
            let val = rows.get(row_idx).and_then(|r| r.get(*ci)).cloned().unwrap_or(Value::Null);
            let nota_ce = match &val {
                Value::Number(n) => n.as_f64().map(|f| f.to_string().replace('.', ",")),
                Value::String(s) if !s.trim().is_empty() => Some(s.replace('.', ",")),
                _ => None,
            };
            if let Some(s) = nota_ce { ce_notas.insert(code.clone(), json!(s)); }
        }
        notes.push(json!({ "numero": notes.len() + 1, "rowIdx": row_idx, "nombre": nombre, "nota": nota_str, "ceNotas": Value::Object(ce_notas) }));
        if let Some(max) = max_alumnos { if notes.len() >= max { break; } }
    }
    notes
}

fn list_unit_sheets(path: &str) -> Result<Vec<Value>, String> {
    let names = sheet_names(path)?;
    let datos_rows = read_sheet_rows(path, "DATOS").unwrap_or_default();
    let mut units_from_datos: HashMap<String, String> = HashMap::new();
    if let Some(start) = find_unidades_start(&datos_rows) {
        for idx in 0..16 {
            let code = { let s = cell_str(&datos_rows, start + idx, 8); if s.is_empty() { format!("U{}", idx + 1) } else { s } };
            let nombre = cell_str(&datos_rows, start + idx, 9);
            units_from_datos.insert(code.to_uppercase(), nombre);
        }
    }
    let re = Regex::new(r"(?i)^U\d+$").unwrap();
    let mut unit_names: Vec<String> = names.iter().filter(|n| re.is_match(n)).cloned().collect();
    unit_names.sort_by_key(|n| n.chars().filter(|c| c.is_ascii_digit()).collect::<String>().parse::<i64>().unwrap_or(0));
    Ok(unit_names.iter().map(|codigo| {
        let nombre = units_from_datos.get(&codigo.to_uppercase()).cloned().unwrap_or_default();
        let label = if nombre.is_empty() { codigo.clone() } else { format!("{} - {}", codigo, nombre) };
        json!({ "codigo": codigo, "nombre": nombre, "label": label })
    }).collect())
}

// ---------------------------------------------------------------------------
// Notas actividad
// ---------------------------------------------------------------------------

fn load_notas_actividad(path: &str, unidad: &str, tipo: &str, actividad: i64, max_alumnos: Option<usize>) -> Result<Value, String> {
    let rows = read_sheet_rows(path, unidad).map_err(|_| format!("El archivo no tiene la hoja \"{unidad}\"."))?;
    let at = get_activity_type(tipo);
    let blocks = find_activity_blocks(&rows, at.key);
    let selected_block = blocks.iter().find(|b| b.numero == actividad).or_else(|| blocks.first());
    let notas = selected_block.map(|b| extract_activity_notes(&rows, b, max_alumnos)).unwrap_or_default();
    let tipos: Vec<Value> = ACTIVITY_TYPES.iter().map(|t| {
        let tblocks: Vec<Value> = find_activity_blocks(&rows, t.key).iter().map(format_activity_block).collect();
        let incluidas = tblocks.iter().filter(|b| b["incluida"].as_bool().unwrap_or(false)).count();
        let total = tblocks.len();
        json!({ "key": t.key, "label": t.label, "actividades": tblocks, "incluidas": incluidas, "total": total })
    }).collect();
    let unidades = list_unit_sheets(path)?;
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();

    // Cargar RRAA y criterios filtrados para la unidad actual
    let (rraa, todas_criterios, ponderaciones_unidad, criterios_unidad) =
        if let Some((rraa, criterios, pu)) = extract_rraa_criterios_data(path) {
            let unidad_idx = unidades.iter().position(|u| u["codigo"].as_str() == Some(unidad));
            let criterios_filtrados = if let Some(idx) = unidad_idx {
                if let Some(pu_entry) = pu.get(idx) {
                    criterios.iter().filter_map(|c| {
                        let ci_key = c["colIdx"].as_u64().unwrap_or(0).to_string();
                        let pesos = pu_entry["ponderaciones"].get(&ci_key)?;
                        let p = pesos["ponderacion"].as_f64().unwrap_or(0.0);
                        let pi = pesos["ponderacionInstituto"].as_f64().unwrap_or(0.0);
                        let pe = pesos["ponderacionEmpresa"].as_f64().unwrap_or(0.0);
                        if p == 0.0 && pi == 0.0 && pe == 0.0 { return None; }
                        let mut c2 = c.clone();
                        if let Value::Object(ref mut obj) = c2 { obj.insert("ponderacionUnidad".to_string(), pesos.clone()); }
                        Some(c2)
                    }).collect()
                } else { vec![] }
            } else { vec![] };
            (rraa, criterios, pu, criterios_filtrados)
        } else {
            (vec![], vec![], vec![], vec![])
        };

    Ok(json!({
        "filePath": path, "fileName": file_name, "unidad": unidad, "tipo": at.key,
        "actividad": selected_block.map(|b| b.numero).unwrap_or(actividad),
        "unidades": unidades, "tipos": tipos,
        "actividades": blocks.iter().map(format_activity_block).collect::<Vec<_>>(),
        "notas": notas,
        "block": selected_block.map(format_activity_block),
        "rraa": rraa,
        "criterios": criterios_unidad,
        "todasCriterios": todas_criterios,
        "ponderacionesUnidad": ponderaciones_unidad
    }))
}

#[tauri::command]
fn excel_get_notas_actividad(payload: Value) -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    let unidad = payload["unidad"].as_str().unwrap_or("U1").to_string();
    let tipo = payload["tipo"].as_str().unwrap_or("practicas").to_string();
    let actividad = payload["actividad"].as_i64().unwrap_or(1);
    let max_alumnos = payload["maxAlumnos"].as_u64().map(|n| n as usize);
    load_notas_actividad(&path, &unidad, &tipo, actividad, max_alumnos)
}

fn load_notas_actividades_tipo(path: &str, unidad: &str, tipo: &str) -> Result<Value, String> {
    let rows = read_sheet_rows(path, unidad).map_err(|_| format!("El archivo no tiene la hoja \"{unidad}\"."))?;
    let at = get_activity_type(tipo);
    let blocks = find_activity_blocks(&rows, at.key);
    let activities: Vec<Value> = blocks.iter().map(|b| {
        let notas = extract_activity_notes(&rows, b, None);
        let mut v = format_activity_block(b);
        v["tipo"] = json!(at.key); v["tipoLabel"] = json!(at.label); v["notas"] = json!(notas); v
    }).collect();
    let tipos: Vec<Value> = ACTIVITY_TYPES.iter().map(|t| {
        let tblocks: Vec<Value> = find_activity_blocks(&rows, t.key).iter().map(format_activity_block).collect();
        let incluidas = tblocks.iter().filter(|b| b["incluida"].as_bool().unwrap_or(false)).count();
        let total = tblocks.len();
        json!({ "key": t.key, "label": t.label, "actividades": tblocks, "incluidas": incluidas, "total": total })
    }).collect();
    let unidades = list_unit_sheets(path)?;
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({ "filePath": path, "fileName": file_name, "unidad": unidad, "unidades": unidades, "tipo": at.key, "tipos": tipos, "activities": activities }))
}

#[tauri::command]
fn excel_get_notas_actividades_tipo(payload: Value) -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    let unidad = payload["unidad"].as_str().unwrap_or("U1").to_string();
    let tipo = payload["tipo"].as_str().unwrap_or("practicas").to_string();
    load_notas_actividades_tipo(&path, &unidad, &tipo)
}

// ---------------------------------------------------------------------------
// Notas evaluación
// ---------------------------------------------------------------------------

fn find_evaluation_sheet_name(names: &[String], evaluacion: &str) -> Option<String> {
    let target = evaluacion.trim();
    if target == "final" { return names.iter().find(|n| normalize_plain(n) == "FINAL").cloned(); }
    if target == "2solo" || target == "3solo" {
        let num = &target[..1];
        return names.iter().find(|n| { let norm = normalize_plain(n); norm.contains(num) && norm.contains("EVA") && norm.contains("SOLO") }).cloned();
    }
    names.iter().find(|n| { let norm = normalize_plain(n); norm.contains(&target.to_uppercase()) && norm.contains("EVA") && !norm.contains("MAX") && !norm.contains("SOLO") }).cloned()
}

fn is_eval_criterion_code(s: &str) -> bool {
    Regex::new(r"^\d+\.?[a-z]\)?$").unwrap().is_match(s.trim())
}

fn load_notas_evaluacion(path: &str, evaluacion: &str) -> Result<Value, String> {
    let names = sheet_names(path)?;
    let sheet_name = find_evaluation_sheet_name(&names, evaluacion)
        .ok_or_else(|| format!("No se encontro la hoja de la {evaluacion} evaluacion."))?;
    let rows = read_sheet_rows(path, &sheet_name)?;

    // Buscar layout
    let (summary_row_idx, code_row_idx, first_student_row_idx) = {
        let mut found = None;
        for row_idx in 0..rows.len() {
            let nota_ce_count = rows[row_idx].iter().filter(|v| normalize_plain(&cell_val_str(v)) == "NOTA CE").count();
            if nota_ce_count == 0 { continue; }
            let next = rows.get(row_idx + 1).cloned().unwrap_or_default();
            if next.iter().filter(|v| is_eval_criterion_code(&cell_val_str(v))).count() > 0 {
                found = Some((row_idx, row_idx + 1, row_idx + 2));
                break;
            }
        }
        found.ok_or("No se encontro la cabecera de notas de evaluacion.")?
    };

    let summary_row = &rows[summary_row_idx];
    let mut ra_columns: Vec<Value> = Vec::new();
    for (ci, cell) in summary_row.iter().enumerate() {
        if normalize_plain(&cell_val_str(cell)) != "NOTA CE" { continue; }
        let code_row = rows.get(code_row_idx).cloned().unwrap_or_default();
        let ra_num = code_row[ci+1..].iter().find_map(|v| {
            let s = cell_val_str(v);
            let norm = normalize_plain(&s);
            if norm == "NOTA CE" || norm == "NOTA FINAL" { return None; }
            s.trim().chars().take_while(|c| c.is_ascii_digit()).collect::<String>().parse::<i64>().ok()
        }).unwrap_or(ra_columns.len() as i64 + 1);
        let peso = cell_str(&rows, 13, ci);
        ra_columns.push(json!({ "colIdx": ci, "address": col_name(ci), "label": format!("RRAA {ra_num}"), "numero": ra_num, "peso": peso }));
    }

    let final_col = summary_row.iter().position(|v| normalize_plain(&cell_val_str(v)) == "NOTA FINAL")
        .or_else(|| {
            let code_row = rows.get(code_row_idx).cloned().unwrap_or_default();
            code_row.iter().enumerate().rev().find_map(|(ci, v)| {
                let s = cell_val_str(v);
                if is_eval_criterion_code(&s) || normalize_plain(&s) == "REC" { Some(ci + 1) } else { None }
            })
        })
        .ok_or("No se encontro la columna NOTA FINAL.")?;

    let code_row = rows.get(code_row_idx).cloned().unwrap_or_default();
    let mut criteria: Vec<Value> = Vec::new();
    for (ra_idx, ra) in ra_columns.iter().enumerate() {
        let ra_ci = ra["colIdx"].as_u64().unwrap() as usize;
        let next_ra_ci = ra_columns.get(ra_idx + 1).and_then(|r| r["colIdx"].as_u64()).map(|n| n as usize).unwrap_or(final_col);
        for ci in (ra_ci + 1)..next_ra_ci {
            let code = cell_val_str(code_row.get(ci).unwrap_or(&Value::Null));
            if !is_eval_criterion_code(&code) { continue; }
            criteria.push(json!({ "colIdx": ci, "address": col_name(ci), "raColIdx": ra_ci, "raLabel": ra["label"], "codigo": code.trim(), "peso": cell_str(&rows, 12, ci) }));
        }
    }

    let mut alumnos: Vec<Value> = Vec::new();
    for row_idx in first_student_row_idx..rows.len() {
        let nombre = cell_str(&rows, row_idx, 0);
        if nombre.is_empty() { continue; }
        let norm = normalize_plain(&nombre);
        if norm.contains("MEDIA") || norm.contains("PONDERACION") { continue; }
        let rraa_vals: Vec<Value> = ra_columns.iter().map(|ra| {
            let ci = ra["colIdx"].as_u64().unwrap() as usize;
            json!({ "colIdx": ci, "label": ra["label"], "nota": cell_f64(&rows, row_idx, ci), "display": cell_str(&rows, row_idx, ci) })
        }).collect();
        let crit_vals: Vec<Value> = criteria.iter().map(|c| {
            let ci = c["colIdx"].as_u64().unwrap() as usize;
            json!({ "colIdx": ci, "raColIdx": c["raColIdx"], "raLabel": c["raLabel"], "codigo": c["codigo"], "nota": cell_f64(&rows, row_idx, ci), "display": cell_str(&rows, row_idx, ci) })
        }).collect();
        alumnos.push(json!({ "rowIdx": row_idx, "numero": alumnos.len() + 1, "nombre": nombre, "final": cell_f64(&rows, row_idx, final_col), "finalDisplay": cell_str(&rows, row_idx, final_col), "rraa": rraa_vals, "criterios": crit_vals }));
    }

    // Filtrar RRAA y criterios sin ningún dato de alumno (columnas vacías = inactivos)
    let active_crit_cols: std::collections::HashSet<usize> = criteria.iter().filter(|c| {
        let ci = c["colIdx"].as_u64().unwrap_or(0) as usize;
        alumnos.iter().any(|a| {
            a["criterios"].as_array().map(|arr| arr.iter().any(|v| {
                v["colIdx"].as_u64().map(|x| x as usize) == Some(ci)
                    && v["nota"].is_number()
                    && v["nota"].as_f64().unwrap_or(0.0) != 0.0
            })).unwrap_or(false)
        })
    }).map(|c| c["colIdx"].as_u64().unwrap_or(0) as usize).collect();

    let active_ra_cols: std::collections::HashSet<usize> = ra_columns.iter().filter(|ra| {
        let ci = ra["colIdx"].as_u64().unwrap_or(0) as usize;
        alumnos.iter().any(|a| {
            a["rraa"].as_array().map(|arr| arr.iter().any(|v| {
                v["colIdx"].as_u64().map(|x| x as usize) == Some(ci)
                    && v["nota"].is_number()
                    && v["nota"].as_f64().unwrap_or(0.0) != 0.0
            })).unwrap_or(false)
        })
    }).map(|ra| ra["colIdx"].as_u64().unwrap_or(0) as usize).collect();

    let criteria_filtered: Vec<Value> = criteria.into_iter().filter(|c| {
        active_crit_cols.contains(&(c["colIdx"].as_u64().unwrap_or(0) as usize))
    }).collect();
    let ra_columns_filtered: Vec<Value> = ra_columns.into_iter().filter(|ra| {
        active_ra_cols.contains(&(ra["colIdx"].as_u64().unwrap_or(0) as usize))
    }).collect();
    let alumnos_filtered: Vec<Value> = alumnos.into_iter().map(|mut a| {
        if let Some(arr) = a["criterios"].as_array_mut() {
            *arr = arr.iter().filter(|v| {
                let ci = v["colIdx"].as_u64().unwrap_or(0) as usize;
                active_crit_cols.contains(&ci)
            }).cloned().collect();
        }
        if let Some(arr) = a["rraa"].as_array_mut() {
            *arr = arr.iter().filter(|v| {
                let ci = v["colIdx"].as_u64().unwrap_or(0) as usize;
                active_ra_cols.contains(&ci)
            }).cloned().collect();
        }
        a
    }).collect();

    let title = if normalize_plain(&sheet_name) == "FINAL" { "FINAL".to_string() } else { cell_str(&rows, 2, 0) };
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({
        "filePath": path, "fileName": file_name, "sheetName": sheet_name, "title": title, "evaluacion": evaluacion,
        "layout": { "summaryRowIdx": summary_row_idx, "codeRowIdx": code_row_idx, "firstStudentRowIdx": first_student_row_idx },
        "raColumns": ra_columns_filtered, "criteria": criteria_filtered, "alumnos": alumnos_filtered
    }))
}

#[tauri::command]
fn excel_get_notas_evaluacion(payload: Value) -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    load_notas_evaluacion(&path, payload["evaluacion"].as_str().unwrap_or("1"))
}

#[tauri::command]
fn excel_get_notas_evaluacion_alumno(payload: Value) -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    let evaluacion = payload["evaluacion"].as_str().unwrap_or("1").to_string();
    let alumno = payload["alumno"].as_str().unwrap_or("").to_string();
    let mut data = load_notas_evaluacion(&path, &evaluacion)?;
    if let Some(arr) = data["alumnos"].as_array_mut() {
        let filtered: Vec<Value> = arr.iter().filter(|a| a["nombre"].as_str().unwrap_or("") == alumno).cloned().collect();
        data["alumnos"] = json!(filtered);
    }
    Ok(data)
}

// Lee la nota final de la unidad (col E = índice 4).
// Usa el primer bloque de prácticas para localizar name_col y first_student_row,
// garantizando que leemos exactamente las mismas filas de alumnos que el resto del sistema.
fn load_notas_unidad(path: &str, unidad: &str) -> Result<Value, String> {
    let rows = read_sheet_rows(path, unidad)
        .map_err(|_| format!("No se encontró la hoja \"{unidad}\"."))?;
    let unidades = list_unit_sheets(path)?;

    let nota_col: usize = 4; // columna E = nota final de la unidad

    // La nota final de la unidad está fija en E5 (índice fila 4, col 4).
    // Los nombres de alumnos están en la columna con "NOMBRE Y APELLIDOS" de la fila 3 (índice 3).
    // Si no se encuentra, buscamos la columna con "NOMBRE" en las primeras 5 filas.
    // Nombres en col D (índice 3), notas en col E (índice 4), desde fila 5 (índice 4).
    // Ignorar filas sin nombre en D. Parar al llegar a una fila vacía tras haber recogido alumnos.
    let name_col: usize = 3; // columna D
    let first_row: usize = 4; // fila 5 en Excel

    let mut alumnos: Vec<Value> = Vec::new();
    let mut consecutive_empty = 0;
    for ri in first_row..rows.len() {
        let nombre = cell_str(&rows, ri, name_col);
        if nombre.is_empty() {
            consecutive_empty += 1;
            if consecutive_empty >= 5 && !alumnos.is_empty() { break; }
            continue;
        }
        consecutive_empty = 0;
        let norm = normalize_plain(&nombre);
        // Saltar filas que son etiquetas del Excel, no alumnos
        if norm.contains("MEDIA") || norm.contains("PONDERACION")
            || norm.contains("EVALUACION") || norm.contains("EVA")
            || norm.contains("PESO") || norm.contains("DATO")
            || norm.contains("SELECCIONA") || norm.contains("FINAL")
            || norm.starts_with("U") && norm.len() <= 3
        { continue; }
        let nota = cell_f64(&rows, ri, nota_col);
        let display = cell_str(&rows, ri, nota_col);
        alumnos.push(json!({ "nombre": nombre, "nota": nota, "display": display }));
    }

    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({ "filePath": path, "fileName": file_name, "unidad": unidad, "unidades": unidades, "alumnos": alumnos }))
}

#[tauri::command]
fn excel_get_notas_unidad(payload: Value) -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    let unidad = payload["unidad"].as_str().unwrap_or("U1").to_string();
    load_notas_unidad(&path, &unidad)
}

#[tauri::command]
fn excel_get_alumnos_informes() -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    let result = load_alumnos(&path)?;
    Ok(json!(result["alumnos"].as_array().unwrap_or(&vec![]).iter().map(|a| a["nombre"].clone()).collect::<Vec<_>>()))
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;").replace('\'', "&apos;")
}

fn unescape_xml(s: &str) -> String {
    s.replace("&apos;", "'").replace("&quot;", "\"").replace("&gt;", ">").replace("&lt;", "<").replace("&amp;", "&")
}

fn get_xml_attr(tag: &str, name: &str) -> Option<String> {
    let pattern = format!(r#"{}="([^"]*)""#, regex::escape(name).replace("\\:", ":").replace("\\r", "r"));
    // Use raw attribute search for namespaced attrs like r:id
    let search = format!("{}=\"", name);
    if let Some(pos) = tag.find(&search) {
        let after = &tag[pos + search.len()..];
        if let Some(end) = after.find('"') {
            return Some(unescape_xml(&after[..end]));
        }
    }
    let _ = pattern;
    None
}

fn set_xml_attr(tag: &str, name: &str, value: &str) -> String {
    let search = format!("{}=\"", name);
    if let Some(pos) = tag.find(&search) {
        let end = tag[pos + search.len()..].find('"').map(|e| pos + search.len() + e + 1).unwrap_or(tag.len());
        let replacement = format!("{}=\"{}\"", name, escape_xml(value));
        format!("{}{}{}", &tag[..pos], replacement, &tag[end..])
    } else if tag.ends_with("/>") {
        format!("{} {}=\"{}\"/>", &tag[..tag.len()-2], name, escape_xml(value))
    } else {
        let pos = tag.find('>').unwrap_or(tag.len());
        format!("{} {}=\"{}\"{}", &tag[..pos], name, escape_xml(value), &tag[pos..])
    }
}

fn get_xml_row(sheet_xml: &str, row_number: usize) -> Option<String> {
    let pattern = format!(r#"<row\b[^>/]*\br="{row_number}"[^>/]*>[\s\S]*?</row>"#);
    Regex::new(&pattern).ok()?.find(sheet_xml).map(|m| m.as_str().to_string())
}

fn insert_xml_row_at(sheet_xml: &str, row_number: usize) -> Result<String, String> {
    let new_row = format!("<row r=\"{row_number}\"></row>");
    if sheet_xml.contains("</sheetData>") {
        Ok(sheet_xml.replace("</sheetData>", &format!("{new_row}</sheetData>")))
    } else {
        Err("La hoja no tiene una estructura sheetData valida.".to_string())
    }
}

fn col_index_from_ref(cell_ref: &str) -> usize {
    let col_part: String = cell_ref.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    col_index(&col_part)
}

fn insert_xml_cell_in_row(row_xml: &str, cell_xml: &str, col_idx: usize) -> String {
    let cell_re = Regex::new(r#"<c\b[^>/]*(?:/>|>[\s\S]*?</c>)"#).unwrap();
    for m in cell_re.find_iter(row_xml) {
        if let Some(r) = get_xml_attr(m.as_str(), "r") {
            if col_index_from_ref(&r) > col_idx {
                return row_xml.replacen(m.as_str(), &format!("{cell_xml}{}", m.as_str()), 1);
            }
        }
    }
    row_xml.replace("</row>", &format!("{cell_xml}</row>"))
}

fn excel_serial_from_date(value: &str) -> Option<i64> {
    let parts: Vec<&str> = value.split('-').collect();
    if parts.len() == 3 {
        let (y, m, d) = (parts[0].parse::<i32>().ok()?, parts[1].parse::<u32>().ok()?, parts[2].parse::<u32>().ok()?);
        use chrono::NaiveDate;
        let date = NaiveDate::from_ymd_opt(y, m, d)?;
        let epoch = NaiveDate::from_ymd_opt(1970, 1, 1)?;
        Some((date - epoch).num_days() + 25569)
    } else {
        None
    }
}

fn build_xml_cell(cell_ref: &str, value: &Value, value_type: &str, style_id: Option<&str>) -> String {
    let style_attr = style_id.map(|s| format!(" s=\"{}\"", escape_xml(s))).unwrap_or_default();
    match value_type {
        "number" => {
            let num = match value {
                Value::Number(n) => n.as_f64().unwrap_or(0.0),
                Value::String(s) => s.replace(',', ".").parse().unwrap_or(0.0),
                _ => return format!("<c r=\"{cell_ref}\"{style_attr}/>"),
            };
            format!("<c r=\"{cell_ref}\"{style_attr}><v>{num}</v></c>")
        }
        "date" => {
            let s = match value { Value::String(s) => s.clone(), _ => return format!("<c r=\"{cell_ref}\"{style_attr}/>") };
            match excel_serial_from_date(&s) {
                Some(serial) => format!("<c r=\"{cell_ref}\"{style_attr}><v>{serial}</v></c>"),
                None => format!("<c r=\"{cell_ref}\"{style_attr}/>"),
            }
        }
        _ => {
            let text = match value {
                Value::String(s) => s.clone(),
                Value::Number(n) => { if let Some(i) = n.as_i64() { i.to_string() } else { n.as_f64().unwrap_or(0.0).to_string() } }
                Value::Bool(b) => b.to_string(),
                _ => return format!("<c r=\"{cell_ref}\"{style_attr}/>"),
            };
            let space_attr = if text.trim() != text { " xml:space=\"preserve\"" } else { "" };
            format!("<c r=\"{cell_ref}\"{style_attr} t=\"inlineStr\"><is><t{space_attr}>{}</t></is></c>", escape_xml(&text))
        }
    }
}

fn set_xml_cell(sheet_xml: &str, row_idx: usize, col_idx: usize, value: Option<&Value>, value_type: &str) -> Result<String, String> {
    let row_number = row_idx + 1;
    let cell_ref = format!("{}{}", col_name(col_idx), row_number);
    let row_pattern = format!(r#"<row\b[^>/]*\br="{row_number}"[^>/]*>[\s\S]*?</row>"#);
    let row_re = Regex::new(&row_pattern).unwrap();

    let mut xml = sheet_xml.to_string();
    if !row_re.is_match(&xml) { xml = insert_xml_row_at(&xml, row_number)?; }

    let original_row = row_re.find(&xml).map(|m| m.as_str().to_string()).unwrap();
    let cell_pattern = format!(r#"<c\b[^>/]*\br="{}"[^>/]*(?:>[\s\S]*?</c>|\s*/?>)"#, regex::escape(&cell_ref));
    let cell_re = Regex::new(&cell_pattern).unwrap();

    let is_empty_val = matches!(value, None) || matches!(value, Some(Value::Null)) || value.map(|v| v.as_str() == Some("")).unwrap_or(false);
    if is_empty_val {
        let cleared = cell_re.replace(&original_row, "").to_string();
        return Ok(xml.replacen(&original_row, &cleared, 1));
    }

    let val = value.unwrap();
    let style_id = cell_re.find(&original_row).and_then(|m| get_xml_attr(m.as_str(), "s"));
    let new_cell = build_xml_cell(&cell_ref, val, value_type, style_id.as_deref());
    let updated_row = if cell_re.is_match(&original_row) {
        cell_re.replace(&original_row, new_cell.as_str()).to_string()
    } else {
        insert_xml_cell_in_row(&original_row, &new_cell, col_idx)
    };
    Ok(xml.replacen(&original_row, &updated_row, 1))
}

// ---------------------------------------------------------------------------
// ZIP: editar hojas y guardar
// ---------------------------------------------------------------------------

fn find_worksheet_path_in_zip(workbook_xml: &str, rels_xml: &str, sheet_name: &str) -> Result<String, String> {
    let sheet_re = Regex::new(r#"<sheet\b[^>]*>"#).unwrap();
    for m in sheet_re.find_iter(workbook_xml) {
        let tag = m.as_str();
        if get_xml_attr(tag, "name").as_deref() != Some(sheet_name) { continue; }
        let rel_id = get_xml_attr(tag, "r:id").ok_or("No r:id en sheet tag")?;
        let rel_re = Regex::new(r#"<Relationship\b[^>]*>"#).unwrap();
        for rm in rel_re.find_iter(rels_xml) {
            if get_xml_attr(rm.as_str(), "Id").as_deref() == Some(&rel_id) {
                let target = get_xml_attr(rm.as_str(), "Target").ok_or("No Target")?;
                return Ok(if target.starts_with("xl/") { target } else { format!("xl/{target}") });
            }
        }
    }
    Err(format!("No se encontro la hoja \"{sheet_name}\" dentro del libro."))
}

fn assert_worksheet_xml_safe(sheet_xml: &str, sheet_name: &str) -> Result<(), String> {
    let row_re = Regex::new(r#"<row\b[^>/]*(?:/>|>[\s\S]*?</row>)"#).unwrap();
    for m in row_re.find_iter(sheet_xml) {
        let row_xml = m.as_str();
        let row_num = get_xml_attr(row_xml, "r").unwrap_or_default();
        let open2 = Regex::new(r"<c\b").unwrap().find_iter(row_xml).count();
        let closed = row_xml.matches("</c>").count();
        let self_closed = Regex::new(r"<c\b[^>/]*/>").unwrap().find_iter(row_xml).count();
        if open2 != closed + self_closed {
            return Err(format!("El XML generado para {sheet_name}, fila {row_num}, no es seguro."));
        }
    }
    Ok(())
}

fn force_workbook_recalc(xml: &str) -> String {
    let apply = |tag: &str| -> String {
        let t = set_xml_attr(tag, "calcMode", "auto");
        let t = set_xml_attr(&t, "fullCalcOnLoad", "1");
        set_xml_attr(&t, "forceFullCalc", "1")
    };
    let self_re = Regex::new(r"<calcPr\b[^>]*/\s*>").unwrap();
    let open_re = Regex::new(r"<calcPr\b[^>]*>").unwrap();
    if self_re.is_match(xml) {
        self_re.replace(xml, |caps: &regex::Captures| apply(&caps[0])).to_string()
    } else if open_re.is_match(xml) {
        open_re.replace(xml, |caps: &regex::Captures| apply(&caps[0])).to_string()
    } else {
        xml.replace("</workbook>", "<calcPr calcMode=\"auto\" fullCalcOnLoad=\"1\" forceFullCalc=\"1\"/></workbook>")
    }
}

fn edit_workbook_sheets_xml(path: &str, sheet_edits: Vec<(&str, Box<dyn Fn(&str) -> Result<String, String>>)>) -> Result<(), String> {
    let input = std::fs::read(path).map_err(|e| e.to_string())?;
    let cursor = std::io::Cursor::new(input);
    let mut zip = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let mut files: HashMap<String, Vec<u8>> = HashMap::new();
    for i in 0..zip.len() {
        let mut f = zip.by_index(i).map_err(|e| e.to_string())?;
        let name = f.name().to_string();
        let mut buf = Vec::new();
        f.read_to_end(&mut buf).map_err(|e| e.to_string())?;
        files.insert(name, buf);
    }

    let workbook_xml = String::from_utf8_lossy(files.get("xl/workbook.xml").ok_or("No xl/workbook.xml")?).to_string();
    let rels_xml = String::from_utf8_lossy(files.get("xl/_rels/workbook.xml.rels").ok_or("No workbook.xml.rels")?).to_string();

    for (sheet_name, edit_fn) in &sheet_edits {
        let sheet_path = find_worksheet_path_in_zip(&workbook_xml, &rels_xml, sheet_name)?;
        let original = String::from_utf8_lossy(files.get(&sheet_path).ok_or_else(|| format!("No se encontro {sheet_path}"))?).to_string();
        let updated = edit_fn(&original)?;
        assert_worksheet_xml_safe(&updated, sheet_name)?;
        files.insert(sheet_path, updated.into_bytes());
    }

    files.remove("xl/calcChain.xml");
    if let Some(b) = files.get_mut("xl/_rels/workbook.xml.rels") {
        let s = String::from_utf8_lossy(b).to_string();
        *b = Regex::new(r#"<Relationship\b[^>]*Type="[^"]*/calcChain"[^>]*(?:/>|></Relationship>)"#).unwrap().replace_all(&s, "").into_owned().into_bytes();
    }
    if let Some(b) = files.get_mut("[Content_Types].xml") {
        let s = String::from_utf8_lossy(b).to_string();
        *b = Regex::new(r#"<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*(?:/>|></Override>)"#).unwrap().replace_all(&s, "").into_owned().into_bytes();
    }
    if let Some(b) = files.get_mut("xl/workbook.xml") {
        let s = String::from_utf8_lossy(b).to_string();
        *b = force_workbook_recalc(&s).into_bytes();
    }

    let names_ordered: Vec<String> = {
        let input2 = std::fs::read(path).map_err(|e| e.to_string())?;
        let mut z2 = zip::ZipArchive::new(std::io::Cursor::new(input2)).map_err(|e| e.to_string())?;
        (0..z2.len()).filter_map(|i| z2.by_index_raw(i).ok().map(|f| f.name().to_string())).collect()
    };
    let mut final_order: Vec<String> = names_ordered.into_iter().filter(|n| n != "xl/calcChain.xml").collect();
    for k in files.keys() {
        if !final_order.contains(k) { final_order.push(k.clone()); }
    }

    let out_buf = Vec::new();
    let mut zw = zip::ZipWriter::new(std::io::Cursor::new(out_buf));
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .compression_level(Some(6));
    for name in &final_order {
        if let Some(data) = files.get(name) {
            zw.start_file(name, options).map_err(|e| e.to_string())?;
            zw.write_all(data).map_err(|e| e.to_string())?;
        }
    }
    let out_cursor = zw.finish().map_err(|e| e.to_string())?;
    std::fs::write(path, out_cursor.into_inner()).map_err(|e| e.to_string())?;
    Ok(())
}

// ---------------------------------------------------------------------------
// save_alumnos_to_file
// ---------------------------------------------------------------------------

fn save_alumnos_to_file(path: &str, alumnos: &[Value]) -> Result<(), String> {
    let rows = read_sheet_rows(path, "DATOS")?;
    let header_row = find_header_row(&rows, "ALUMNADO").ok_or("No se encontro la seccion ALUMNADO.")?;
    let alumnos_start = header_row + 1;
    let existing_count = (alumnos_start..rows.len()).take_while(|&i| !cell_str(&rows, i, 1).is_empty()).count();
    let rows_to_clear = existing_count.max(alumnos.len());
    let alumnos_owned = alumnos.to_vec();

    edit_workbook_sheets_xml(path, vec![("DATOS", Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        for idx in 0..rows_to_clear {
            let ri = alumnos_start + idx;
            s = set_xml_cell(&s, ri, 0, None, "text")?;
            s = set_xml_cell(&s, ri, 1, None, "text")?;
            s = set_xml_cell(&s, ri, 2, None, "text")?;
        }
        for (idx, alumno) in alumnos_owned.iter().enumerate() {
            let ri = alumnos_start + idx;
            s = set_xml_cell(&s, ri, 0, Some(&alumno["numero"]), "number")?;
            s = set_xml_cell(&s, ri, 1, Some(&alumno["nombre"]), "text")?;
            let fecha = &alumno["fechaNac"];
            if !fecha.is_null() && fecha.as_str().map(|v| !v.is_empty()).unwrap_or(false) {
                s = set_xml_cell(&s, ri, 2, Some(fecha), "date")?;
            }
        }
        Ok(s)
    }))])
}

// ---------------------------------------------------------------------------
// save_unidades_to_file
// ---------------------------------------------------------------------------

fn find_eval_block_start(rows: &[Vec<Value>], eval_num: &str) -> Option<usize> {
    let idx = rows.iter().position(|row| {
        row.get(10).map(|v| {
            let s = cell_val_str(v).to_uppercase();
            s.contains("UNIDADES") && s.contains(eval_num)
        }).unwrap_or(false)
    })?;
    Some(idx + 1)
}

fn sync_eval_unit_blocks_xml(sheet_xml: &str, rows: &[Vec<Value>], unidades: &[Value]) -> Result<String, String> {
    let mut xml = sheet_xml.to_string();
    for (eval_str, eval_label) in [("1", "1ª"), ("2", "2ª"), ("3", "3ª")] {
        let start = match find_eval_block_start(rows, eval_str) { Some(s) => s, None => continue };
        let units_for_eval: Vec<&Value> = unidades.iter().filter(|u| u["evaluacion"].as_str().unwrap_or("").contains(eval_str)).collect();
        for idx in 0..16 {
            let ri = start + idx;
            let u_label = json!(format!("U{}", idx + 1));
            xml = set_xml_cell(&xml, ri, 9, Some(&u_label), "text")?;
            let nombre = units_for_eval.get(idx).and_then(|u| u["nombre"].as_str()).unwrap_or("");
            let nombre_val = json!(nombre);
            xml = set_xml_cell(&xml, ri, 10, if nombre.is_empty() { None } else { Some(&nombre_val) }, "text")?;
            let eval_val = json!(eval_label);
            xml = set_xml_cell(&xml, ri, 11, Some(&eval_val), "text")?;
        }
    }
    Ok(xml)
}

fn save_unidades_to_file(path: &str, unidades: &[Value]) -> Result<(), String> {
    let rows = read_sheet_rows(path, "DATOS")?;
    let start = find_unidades_start(&rows).ok_or("No se encontro la seccion UNIDADES.")?;
    let unidades_owned = unidades.to_vec();
    let rows_owned = rows.clone();

    edit_workbook_sheets_xml(path, vec![("DATOS", Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        for idx in 0..16 {
            let u = unidades_owned.get(idx);
            let ri = start + idx;
            let codigo = u.and_then(|v| v["codigo"].as_str()).filter(|c| !c.is_empty()).map(|c| c.to_string()).unwrap_or_else(|| format!("U{}", idx + 1));
            s = set_xml_cell(&s, ri, 8, Some(&json!(codigo)), "text")?;
            let nombre = u.and_then(|v| v["nombre"].as_str()).unwrap_or("");
            let nombre_val = json!(nombre);
            s = set_xml_cell(&s, ri, 9, if nombre.is_empty() { None } else { Some(&nombre_val) }, "text")?;
            let eval = u.and_then(|v| v["evaluacion"].as_str()).unwrap_or("");
            let eval_val = json!(eval);
            s = set_xml_cell(&s, ri, 10, if eval.is_empty() { None } else { Some(&eval_val) }, "text")?;
            let horas_str = u.and_then(|v| v["horas"].as_str()).unwrap_or("");
            if horas_str.is_empty() {
                s = set_xml_cell(&s, ri, 11, None, "number")?;
            } else {
                let h: f64 = horas_str.parse().unwrap_or(0.0);
                s = set_xml_cell(&s, ri, 11, Some(&json!(h)), "number")?;
            }
        }
        s = sync_eval_unit_blocks_xml(&s, &rows_owned, &unidades_owned)?;
        Ok(s)
    }))])
}

// ---------------------------------------------------------------------------
// save_rraa_criterios_to_file
// ---------------------------------------------------------------------------

fn find_criterion_text_row(rows: &[Vec<Value>], code: &str) -> Option<usize> {
    let normalized = normalize_criterion_code(code);
    rows.iter().position(|row| normalize_criterion_code(&row.get(21).map(cell_val_str).unwrap_or_default()) == normalized)
}

fn save_rraa_criterios_to_file(path: &str, rraa: &[Value], criterios: &[Value], pond_unidad: &[Value]) -> Result<(), String> {
    let datos_rows = read_sheet_rows(path, "DATOS")?;
    let rraa_start = find_rraa_start(&datos_rows).ok_or("No se encontro la seccion RRAA.")?;
    let rraa_owned = rraa.to_vec(); let criterios_owned = criterios.to_vec();
    let pond_owned = pond_unidad.to_vec(); let datos_owned = datos_rows.clone();

    edit_workbook_sheets_xml(path, vec![
        ("DATOS", Box::new({
            let rraa2 = rraa_owned.clone(); let criterios2 = criterios_owned.clone(); let datos2 = datos_owned.clone();
            move |xml: &str| {
                let mut s = xml.to_string();
                for (idx, item) in rraa2.iter().enumerate() {
                    s = set_xml_cell(&s, rraa_start + idx, 5, Some(&item["numero"]), "number")?;
                    s = set_xml_cell(&s, rraa_start + idx, 6, Some(&item["descripcion"]), "text")?;
                }
                for criterio in &criterios2 {
                    let codigo = criterio["codigo"].as_str().unwrap_or("").to_string();
                    let original = criterio["originalCodigo"].as_str().unwrap_or(&codigo).to_string();
                    if let Some(ri) = find_criterion_text_row(&datos2, &original).or_else(|| find_criterion_text_row(&datos2, &codigo)) {
                        let stripped = codigo.trim_end_matches(')').to_string();
                        s = set_xml_cell(&s, ri, 21, Some(&json!(stripped)), "text")?;
                        s = set_xml_cell(&s, ri, 22, Some(&criterio["texto"]), "text")?;
                    }
                }
                Ok(s)
            }
        })),
        ("PESOS", Box::new({
            let criterios3 = criterios_owned.clone(); let pond3 = pond_owned.clone();
            move |xml: &str| {
                let mut s = xml.to_string();
                for criterio in &criterios3 {
                    if let Some(ci) = criterio["colIdx"].as_u64().map(|n| n as usize) {
                        s = set_xml_cell(&s, 3, ci, Some(&criterio["codigo"]), "text")?;
                    }
                }
                for unidad in &pond3 {
                    if let Some(ri) = unidad["rowIdx"].as_u64().map(|n| n as usize) {
                        if let Some(ponds) = unidad["ponderaciones"].as_object() {
                            for (col_key, vals) in ponds {
                                let ci: usize = col_key.parse().unwrap_or(0);
                                s = set_xml_cell(&s, ri, ci, Some(&json!(parse_decimal(&vals["ponderacion"]))), "number")?;
                                s = set_xml_cell(&s, ri, ci + 1, Some(&json!(parse_decimal(&vals["ponderacionInstituto"]))), "number")?;
                                s = set_xml_cell(&s, ri, ci + 2, Some(&json!(parse_decimal(&vals["ponderacionEmpresa"]))), "number")?;
                            }
                        }
                    }
                }
                Ok(s)
            }
        })),
    ])
}

// ---------------------------------------------------------------------------
// save_notas_actividad
// ---------------------------------------------------------------------------

fn normalize_grade(value: &Value) -> Option<f64> {
    match value {
        Value::Number(n) => n.as_f64(),
        Value::String(s) => { let s2 = s.replace(',', "."); if s2.trim().is_empty() { None } else { s2.trim().parse().ok() } }
        _ => None,
    }
}

#[tauri::command]
fn excel_save_notas_actividad(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();
    let tipo = payload["tipo"].as_str().ok_or("Falta tipo")?.to_string();
    let actividad = payload["actividad"].as_i64().unwrap_or(1);
    let notas = payload["notas"].as_array().ok_or("Falta notas")?.clone();
    let nombre_actividad = payload["nombreActividad"].as_str().map(|s| s.to_string());
    let incluida = payload["incluida"].as_bool().unwrap_or(false);

    let rows = read_sheet_rows(&path, &unidad).map_err(|_| format!("El archivo no tiene la hoja \"{unidad}\"."))?;
    let at = get_activity_type(&tipo);
    let blocks = find_activity_blocks(&rows, at.key);
    let block = blocks.iter().find(|b| b.numero == actividad)
        .ok_or_else(|| format!("No se encontro la actividad {actividad} de {} en {unidad}.", at.label))?.clone();

    let path_clone = path.clone(); let unidad_clone = unidad.clone(); let tipo_clone = tipo.clone();

    edit_workbook_sheets_xml(&path, vec![(&unidad, Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        let nv = nombre_actividad.as_deref().map(|n| json!(n));
        s = set_xml_cell(&s, block.number_row, block.name_value_col, nv.as_ref(), "text")?;
        let x_val = json!("X");
        s = set_xml_cell(&s, block.included_row, block.included_col, if incluida { Some(&x_val) } else { None }, "text")?;
        for nota_item in &notas {
            if let Some(ri) = nota_item["rowIdx"].as_u64().map(|n| n as usize) {
                if ri < block.first_student_row { continue; }
                if cell_str(&rows, ri, block.name_col).is_empty() { continue; }
                match normalize_grade(&nota_item["nota"]) {
                    Some(n) => { s = set_xml_cell(&s, ri, block.note_col, Some(&json!(n)), "number")?; }
                    None    => { s = set_xml_cell(&s, ri, block.note_col, None, "number")?; }
                }
            }
        }
        Ok(s)
    }) as Box<dyn Fn(&str) -> Result<String, String>>)])?;

    load_notas_actividad(&path_clone, &unidad_clone, &tipo_clone, actividad, None)
}

// ---------------------------------------------------------------------------
// add_actividad
// ---------------------------------------------------------------------------

const ACTIVITY_BLOCK_ROWS: usize = 41;
const ACTIVITY_BLOCK_STRIDE: usize = 44;

fn shift_formula_rows(formula: &str, row_delta: i64, src_start: i64, src_end: i64) -> String {
    Regex::new(r"(\$?[A-Z]{1,3})(\$?)(\d+)").unwrap().replace_all(formula, |caps: &regex::Captures| {
        let col = &caps[1]; let abs_row = &caps[2]; let row_num: i64 = caps[3].parse().unwrap_or(0);
        if !abs_row.is_empty() || row_num < src_start || row_num > src_end { caps[0].to_string() }
        else { format!("{}{}", col, row_num + row_delta) }
    }).to_string()
}

fn shift_formula_ref_attr(tag: &str, row_delta: i64, src_start: i64, src_end: i64) -> String {
    if let Some(ref_val) = get_xml_attr(tag, "ref") {
        set_xml_attr(tag, "ref", &shift_formula_rows(&ref_val, row_delta, src_start, src_end))
    } else { tag.to_string() }
}

fn extract_and_adjust_cells(source_row_xml: &str, dst_row: i64, row_delta: i64, src_start: i64, src_end: i64, col_start: usize, col_end: usize) -> Vec<String> {
    let cell_re = Regex::new(r#"<c\b[^>/]*(?:/>|>[\s\S]*?</c>)"#).unwrap();
    let mut cells = Vec::new();
    for m in cell_re.find_iter(source_row_xml) {
        let cell_xml = m.as_str();
        let ref_attr = match get_xml_attr(cell_xml, "r") { Some(r) => r, None => continue };
        let col_part: String = ref_attr.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
        let ci = col_index(&col_part);
        if ci < col_start || ci > col_end { continue; }
        let new_ref = format!("{}{}", col_part, dst_row);
        let mut adjusted = cell_xml.replacen(&format!(r#"r="{ref_attr}""#), &format!(r#"r="{new_ref}""#), 1);
        adjusted = Regex::new(r#"(<f\b[^>/]*>)([\s\S]*?)(</f>)"#).unwrap().replace_all(&adjusted, |caps: &regex::Captures| {
            format!("{}{}{}", shift_formula_ref_attr(&caps[1], row_delta, src_start, src_end), shift_formula_rows(&caps[2], row_delta, src_start, src_end), &caps[3])
        }).to_string();
        cells.push(adjusted);
    }
    cells
}

fn insert_xml_row_ordered(sheet_xml: &str, row_xml: &str) -> Result<String, String> {
    let target: i64 = get_xml_attr(row_xml, "r").and_then(|s| s.parse().ok()).unwrap_or(0);
    let row_re = Regex::new(r#"<row\b[^>/]*\br="(\d+)"[^>/]*>[\s\S]*?</row>"#).unwrap();
    for m in row_re.find_iter(sheet_xml) {
        let n: i64 = get_xml_attr(m.as_str(), "r").and_then(|s| s.parse().ok()).unwrap_or(0);
        if n > target { return Ok(sheet_xml.replacen(m.as_str(), &format!("{row_xml}{}", m.as_str()), 1)); }
    }
    if sheet_xml.contains("</sheetData>") { Ok(sheet_xml.replace("</sheetData>", &format!("{row_xml}</sheetData>"))) }
    else { Err("La hoja no tiene una estructura sheetData valida.".to_string()) }
}

fn paste_range_cells(sheet_xml: &str, dst_row: usize, cells: &[String], col_start: usize, col_end: usize) -> Result<String, String> {
    let cell_re = Regex::new(r#"<c\b[^>/]*(?:/>|>[\s\S]*?</c>)"#).unwrap();
    match get_xml_row(sheet_xml, dst_row) {
        None => {
            let new_row = format!("<row r=\"{dst_row}\">{}</row>", cells.join(""));
            insert_xml_row_ordered(sheet_xml, &new_row)
        }
        Some(target_row_xml) => {
            let mut updated = cell_re.replace_all(&target_row_xml, |caps: &regex::Captures| {
                if let Some(r) = get_xml_attr(&caps[0], "r") {
                    let ci = col_index_from_ref(&r);
                    if ci >= col_start && ci <= col_end { return String::new(); }
                }
                caps[0].to_string()
            }).to_string();
            for cell in cells {
                let ref_attr = get_xml_attr(cell, "r").unwrap_or_default();
                updated = insert_xml_cell_in_row(&updated, cell, col_index_from_ref(&ref_attr));
            }
            Ok(sheet_xml.replacen(&target_row_xml, &updated, 1))
        }
    }
}

fn copy_activity_merges_xml(sheet_xml: &str, source_start: usize, source_end: usize, target_start: usize, type_start_col: usize, type_end_col: usize) -> Result<String, String> {
    let row_delta = (target_start as i64) - (source_start as i64);
    let src_s = (source_start + 1) as i64; let src_e = (source_end + 1) as i64;
    let tgt_s = (target_start + 1) as i64; let tgt_e = src_e + row_delta;

    let decode = |s: &str| -> Option<(usize, i64)> {
        let col: String = s.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
        let row: i64 = s.chars().skip_while(|c| c.is_ascii_alphabetic()).collect::<String>().parse().ok()?;
        Some((col_index(&col), row))
    };

    let merge_re = Regex::new(r#"<mergeCell\b[^>]*\bref="([^"]+)"[^>]*/>"#).unwrap();
    let mut keep: Vec<String> = Vec::new(); let mut new_refs: Vec<String> = Vec::new();

    for m in merge_re.find_iter(sheet_xml) {
        let ref_val = get_xml_attr(m.as_str(), "ref").unwrap_or_default();
        let parts: Vec<&str> = ref_val.split(':').collect();
        let (sc, sr) = match decode(parts[0]) { Some(v) => v, None => { keep.push(ref_val); continue; } };
        let (ec, er) = match decode(parts.get(1).copied().unwrap_or(parts[0])) { Some(v) => v, None => { keep.push(ref_val); continue; } };
        let tgt_overlap = !(er < tgt_s || sr > tgt_e || ec < type_start_col || sc > type_end_col);
        if !tgt_overlap { keep.push(ref_val.clone()); }
        if sr >= src_s && er <= src_e && sc >= type_start_col && ec <= type_end_col {
            let ns = format!("{}{}", col_name(sc), sr + row_delta);
            let ne = format!("{}{}", col_name(ec), er + row_delta);
            new_refs.push(if ns == ne { ns } else { format!("{ns}:{ne}") });
        }
    }

    let all_refs: Vec<String> = { let mut seen = std::collections::HashSet::new(); keep.into_iter().chain(new_refs).filter(|r| seen.insert(r.clone())).collect() };
    let merge_xml = if all_refs.is_empty() { String::new() } else {
        format!("<mergeCells count=\"{}\">{}</mergeCells>", all_refs.len(), all_refs.iter().map(|r| format!("<mergeCell ref=\"{}\"/>", escape_xml(r))).collect::<String>())
    };
    let full_re = Regex::new(r#"<mergeCells\b[\s\S]*?</mergeCells>"#).unwrap();
    if full_re.is_match(sheet_xml) { Ok(full_re.replace(sheet_xml, merge_xml.as_str()).to_string()) }
    else if merge_xml.is_empty() { Ok(sheet_xml.to_string()) }
    else { Ok(sheet_xml.replace("</sheetData>", &format!("</sheetData>{merge_xml}"))) }
}

// Expande las "shared formulas" de Excel dentro del rango dado.
// Celdas secundarias tienen <f t="shared" si=N/> (sin fórmula). Sin expandir, al copiar
// quedan vacías y el Excel se corrompe.
fn expand_shared_formulas_in_range(sheet_xml: &str, src_row_start: i64, src_row_end: i64, col_start: usize, col_end: usize) -> String {
    let cell_re = Regex::new(r#"<c\b[^>/]*(?:/>|>[\s\S]*?</c>)"#).unwrap();

    // Paso 1: recopilar masters (celdas con <f t="shared" si=N ref=...>FORMULA</f>)
    struct Master { formula: String, master_col: String, master_row: i64 }
    let mut masters: HashMap<String, Master> = HashMap::new();

    for m in cell_re.find_iter(sheet_xml) {
        let cell = m.as_str();
        let ref_attr = match get_xml_attr(cell, "r") { Some(r) => r, None => continue };
        let f_match = Regex::new(r#"<f\b([^>/]*)>([\s\S]*?)</f>"#).unwrap();
        if let Some(fm) = f_match.captures(cell) {
            let f_tag_inner = fm.get(1).map(|x| x.as_str()).unwrap_or("");
            let formula = fm.get(2).map(|x| x.as_str().trim().to_string()).unwrap_or_default();
            let f_tag = format!("<f {f_tag_inner}>");
            let t_attr = get_xml_attr(&f_tag, "t");
            let si_attr = get_xml_attr(&f_tag, "si");
            let ref2 = get_xml_attr(&f_tag, "ref");
            if t_attr.as_deref() == Some("shared") && si_attr.is_some() && ref2.is_some() && !formula.is_empty() {
                let col_part: String = ref_attr.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
                let row_num: i64 = ref_attr.chars().skip_while(|c| c.is_ascii_alphabetic()).collect::<String>().parse().unwrap_or(0);
                masters.insert(si_attr.unwrap(), Master { formula, master_col: col_part, master_row: row_num });
            }
        }
    }

    if masters.is_empty() { return sheet_xml.to_string(); }

    // Paso 2: reemplazar celdas secundarias (self-closing <f t="shared" si=N/>) en el rango
    cell_re.replace_all(sheet_xml, |caps: &regex::Captures| {
        let cell = &caps[0];
        let ref_attr = match get_xml_attr(cell, "r") { Some(r) => r, None => return cell.to_string() };
        let col_part: String = ref_attr.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
        let row_num: i64 = ref_attr.chars().skip_while(|c| c.is_ascii_alphabetic()).collect::<String>().parse().unwrap_or(0);
        let col_idx = col_index(&col_part);
        if col_idx < col_start || col_idx > col_end || row_num < src_row_start || row_num > src_row_end {
            return cell.to_string();
        }
        // Buscar self-closing shared formula: <f ... />
        let sc_re = Regex::new(r#"<f\b([^>/]*)/>"#).unwrap();
        if let Some(fm) = sc_re.captures(cell) {
            let f_inner = fm.get(1).map(|x| x.as_str()).unwrap_or("");
            let f_tag = format!("<f {f_inner}>");
            let t_attr = get_xml_attr(&f_tag, "t");
            let si_attr = get_xml_attr(&f_tag, "si");
            if t_attr.as_deref() == Some("shared") {
                if let Some(si) = si_attr {
                    if let Some(master) = masters.get(&si) {
                        let row_delta = row_num - master.master_row;
                        let concrete = Regex::new(r"(\$?[A-Z]{1,3})(\$?)(\d+)").unwrap().replace_all(&master.formula, |c2: &regex::Captures| {
                            let col = &c2[1]; let abs_row = &c2[2]; let rn: i64 = c2[3].parse().unwrap_or(0);
                            if !abs_row.is_empty() { c2[0].to_string() } else { format!("{}{}", col, rn + row_delta) }
                        }).to_string();
                        let expanded_f = format!("<f>{concrete}</f>");
                        return cell.replacen(fm.get(0).unwrap().as_str(), &expanded_f, 1);
                    }
                }
            }
        }
        cell.to_string()
    }).to_string()
}

fn copy_activity_block_xml(sheet_xml: &str, source_start: usize, source_end: usize, target_start: usize,
    type_start_col: usize, type_end_col: usize, number_col: usize, name_value_col: usize,
    included_row_offset: usize, included_col: usize, first_student_row_offset: usize, note_col: usize,
    notes_to_clear: usize, new_number: i64, nombre_actividad: &str, incluida: bool
) -> Result<String, String> {
    let row_delta = (target_start as i64) - (source_start as i64);
    let src_s = (source_start + 1) as i64; let src_e = (source_end + 1) as i64;
    // Expandir shared formulas antes de copiar para evitar corrupción del Excel
    let mut xml = expand_shared_formulas_in_range(sheet_xml, src_s, src_e, type_start_col, type_end_col);

    for row_idx in source_start..=source_end {
        let src_row = row_idx + 1;
        let dst_row = (src_row as i64 + row_delta) as usize;
        let source_row_xml = match get_xml_row(&xml, src_row) { Some(r) => r, None => continue };
        let cloned = extract_and_adjust_cells(&source_row_xml, dst_row as i64, row_delta, src_s, src_e, type_start_col, type_end_col);
        xml = paste_range_cells(&xml, dst_row, &cloned, type_start_col, type_end_col)?;
    }

    xml = copy_activity_merges_xml(&xml, source_start, source_end, target_start, type_start_col, type_end_col)?;
    xml = set_xml_cell(&xml, target_start + 1, number_col, Some(&json!(new_number)), "number")?;
    let nombre_val = json!(nombre_actividad);
    xml = set_xml_cell(&xml, target_start + 1, name_value_col, if nombre_actividad.is_empty() { None } else { Some(&nombre_val) }, "text")?;
    let x_val = json!("X");
    xml = set_xml_cell(&xml, target_start + included_row_offset, included_col, if incluida { Some(&x_val) } else { None }, "text")?;

    let student_start_row = target_start + first_student_row_offset + 1;
    let student_end_row = student_start_row + notes_to_clear;
    let row_re = Regex::new(r#"<row\b[^>/]*\br="(\d+)"[^>/]*>[\s\S]*?</row>"#).unwrap();
    let cell_re2 = Regex::new(r#"<c\b[^>/]*(?:/>|>[\s\S]*?</c>)"#).unwrap();
    xml = row_re.replace_all(&xml, |caps: &regex::Captures| {
        let row_xml = &caps[0];
        let rn: usize = get_xml_attr(row_xml, "r").and_then(|s| s.parse().ok()).unwrap_or(0);
        if rn < student_start_row || rn >= student_end_row { return row_xml.to_string(); }
        cell_re2.replace_all(row_xml, |caps2: &regex::Captures| {
            let cell = &caps2[0];
            if let Some(r) = get_xml_attr(cell, "r") {
                if col_index_from_ref(&r) == note_col {
                    let style = get_xml_attr(cell, "s").map(|s| format!(" s=\"{}\"", escape_xml(&s))).unwrap_or_default();
                    return format!("<c r=\"{r}\"{style}><v>0</v></c>");
                }
            }
            cell.to_string()
        }).to_string()
    }).to_string();

    Ok(xml)
}

#[tauri::command]
fn excel_add_actividad(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().unwrap_or("U1").to_string();
    let tipo = payload["tipo"].as_str().unwrap_or("practicas").to_string();
    let nombre_actividad = payload["nombreActividad"].as_str().unwrap_or("").to_string();
    let incluida = payload["incluida"].as_bool().unwrap_or(true);

    let rows = read_sheet_rows(&path, &unidad).map_err(|_| format!("El archivo no tiene la hoja \"{unidad}\"."))?;
    let at = get_activity_type(&tipo);
    let fixed = activity_fixed_cols(at.key);
    let mut blocks = find_activity_blocks(&rows, at.key);
    blocks.sort_by_key(|b| b.numero);
    let prev_block = blocks.last().ok_or_else(|| format!("No se encontro ninguna actividad previa de {} en {unidad}.", at.label))?.clone();
    let new_number = prev_block.numero + 1;
    let source_start = prev_block.title_row;
    let source_end = source_start + ACTIVITY_BLOCK_ROWS - 1;
    let target_start = source_start + ACTIVITY_BLOCK_STRIDE;
    let notes_to_clear = extract_activity_notes(&rows, &prev_block, None).len();
    let included_row_offset = prev_block.included_row - prev_block.title_row;
    let first_student_row_offset = prev_block.first_student_row - prev_block.title_row;
    let path_clone = path.clone(); let unidad_clone = unidad.clone(); let tipo_clone = tipo.clone();

    edit_workbook_sheets_xml(&path, vec![(&unidad, Box::new(move |xml: &str| {
        copy_activity_block_xml(xml, source_start, source_end, target_start, fixed.start, fixed.end,
            prev_block.number_col, prev_block.name_value_col, included_row_offset, prev_block.included_col,
            first_student_row_offset, prev_block.note_col, notes_to_clear, new_number, &nombre_actividad, incluida)
    }))])?;

    load_notas_actividad(&path_clone, &unidad_clone, &tipo_clone, new_number, None)
}

// ---------------------------------------------------------------------------
// save_ce_notas (batch: escribe la nota de un CE en todas las filas de alumnos)
// ---------------------------------------------------------------------------

#[tauri::command]
fn excel_save_ce_notas(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();
    let tipo = payload["tipo"].as_str().ok_or("Falta tipo")?.to_string();
    let actividad = payload["actividad"].as_i64().unwrap_or(1);
    let ce_notas = payload["ceNotas"].as_object().cloned().unwrap_or_default();

    if ce_notas.is_empty() {
        return load_notas_actividad(&path, &unidad, &tipo, actividad, None);
    }

    let rows = read_sheet_rows(&path, &unidad).map_err(|_| format!("El archivo no tiene la hoja \"{unidad}\"."))?;
    let at = get_activity_type(&tipo);
    let blocks = find_activity_blocks(&rows, at.key);
    let block = blocks.iter().find(|b| b.numero == actividad)
        .ok_or_else(|| format!("No se encontro la actividad {actividad} de {} en {unidad}.", at.label))?.clone();

    let path_clone = path.clone(); let unidad_clone = unidad.clone(); let tipo_clone = tipo.clone();

    edit_workbook_sheets_xml(&path, vec![(&unidad, Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        let mut ri = block.first_student_row;
        loop {
            if ri >= rows.len() { break; }
            let nombre = cell_str(&rows, ri, block.name_col);
            if nombre.is_empty() || nombre == "0" { break; }
            for (code, ci) in &block.ce_cols {
                if let Some(val) = ce_notas.get(code) {
                    if let Some(n) = normalize_grade(val) {
                        s = set_xml_cell(&s, ri, *ci, Some(&json!(n)), "number")?;
                    }
                }
            }
            ri += 1;
        }
        Ok(s)
    }) as Box<dyn Fn(&str) -> Result<String, String>>)])?;

    load_notas_actividad(&path_clone, &unidad_clone, &tipo_clone, actividad, None)
}

// ---------------------------------------------------------------------------
// open external + main
// ---------------------------------------------------------------------------

#[tauri::command]
fn app_open_external(url: String) -> Result<(), String> {
    webbrowser::open(&url).map_err(|e| format!("No se pudo abrir el enlace: {e}"))
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            excel_select_file, excel_get_selected_file, excel_set_selected_file,
            excel_verify_file_exists, excel_save_alumnos, excel_get_unidades,
            excel_save_unidades, excel_get_rraa_criterios, excel_save_rraa_criterios,
            excel_get_notas_actividad, excel_get_notas_actividades_tipo,
            excel_save_notas_actividad, excel_save_ce_notas, excel_add_actividad,
            excel_get_notas_evaluacion, excel_get_notas_evaluacion_alumno,
            excel_get_notas_unidad, excel_get_alumnos_informes, app_open_external
        ])
        .run(tauri::generate_context!())
        .expect("error al ejecutar la aplicacion Tauri");
}
