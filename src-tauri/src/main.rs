use serde::Deserialize;
use serde_json::{json, Value};
use std::io::Write;
use std::path::PathBuf;
use std::process::{Command, Stdio};

#[derive(Debug, Deserialize)]
struct BridgeResponse {
    ok: bool,
    data: Option<Value>,
    error: Option<String>,
}

fn project_root() -> Result<PathBuf, String> {
    let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    manifest_dir
        .parent()
        .map(PathBuf::from)
        .ok_or_else(|| "No se pudo localizar la carpeta del proyecto.".to_string())
}

fn call_node_backend(command: &str, payload: Value) -> Result<Value, String> {
    let root = project_root()?;
    let script = root.join("tauri-node-backend.js");
    let request = json!({
        "command": command,
        "payload": payload
    });

    let mut child = Command::new("node")
        .arg(script)
        .current_dir(root)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|error| format!("No se pudo lanzar Node para leer el Excel: {error}"))?;

    if let Some(stdin) = child.stdin.as_mut() {
        stdin
            .write_all(request.to_string().as_bytes())
            .map_err(|error| format!("No se pudo enviar el comando al puente Tauri: {error}"))?;
    }

    let output = child
        .wait_with_output()
        .map_err(|error| format!("No se pudo recibir la respuesta del puente Tauri: {error}"))?;

    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);

    if stdout.trim().is_empty() {
        return Err(format!("El puente Tauri no devolvio datos. {stderr}"));
    }

    let response: BridgeResponse = serde_json::from_str(stdout.trim())
        .map_err(|error| format!("Respuesta no valida del puente Tauri: {error}. {stderr}"))?;

    if response.ok {
        Ok(response.data.unwrap_or(Value::Null))
    } else {
        Err(response.error.unwrap_or_else(|| "Error desconocido en el puente Tauri.".to_string()))
    }
}

#[tauri::command]
fn excel_select_file() -> Result<Value, String> {
    let file = rfd::FileDialog::new()
        .set_title("Selecciona la plantilla Excel")
        .add_filter("Excel", &["xlsx", "xlsm", "xls"])
        .pick_file();

    match file {
        Some(path) => call_node_backend("selectFile", json!({ "filePath": path.to_string_lossy() })),
        None => Ok(Value::Null),
    }
}

#[tauri::command]
fn excel_get_selected_file() -> Result<Value, String> {
    call_node_backend("getSelectedFile", json!({}))
}

#[tauri::command]
fn excel_save_alumnos(alumnos: Value) -> Result<Value, String> {
    call_node_backend("saveAlumnos", alumnos)
}

#[tauri::command]
fn excel_get_unidades() -> Result<Value, String> {
    call_node_backend("getUnidades", json!({}))
}

#[tauri::command]
fn excel_save_unidades(unidades: Value) -> Result<Value, String> {
    call_node_backend("saveUnidades", unidades)
}

#[tauri::command]
fn excel_get_rraa_criterios() -> Result<Value, String> {
    call_node_backend("getRraaCriterios", json!({}))
}

#[tauri::command]
fn excel_save_rraa_criterios(payload: Value) -> Result<Value, String> {
    call_node_backend("saveRraaCriterios", payload)
}

#[tauri::command]
fn excel_get_notas_actividad(payload: Value) -> Result<Value, String> {
    call_node_backend("getNotasActividad", payload)
}

#[tauri::command]
fn excel_save_notas_actividad(payload: Value) -> Result<Value, String> {
    call_node_backend("saveNotasActividad", payload)
}

#[tauri::command]
fn excel_get_notas_evaluacion(payload: Value) -> Result<Value, String> {
    call_node_backend("getNotasEvaluacion", payload)
}

#[tauri::command]
fn app_open_external(url: String) -> Result<(), String> {
    webbrowser::open(&url).map_err(|error| format!("No se pudo abrir el enlace: {error}"))
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            excel_select_file,
            excel_get_selected_file,
            excel_save_alumnos,
            excel_get_unidades,
            excel_save_unidades,
            excel_get_rraa_criterios,
            excel_save_rraa_criterios,
            excel_get_notas_actividad,
            excel_save_notas_actividad,
            excel_get_notas_evaluacion,
            app_open_external
        ])
        .run(tauri::generate_context!())
        .expect("error al ejecutar la aplicacion Tauri");
}
