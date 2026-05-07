# Migracion a Tauri

La aplicacion ya tiene una primera base Tauri en `src-tauri/`.

## Estado actual

- La interfaz HTML se reutiliza casi completa.
- `app-bridge.js` mantiene la API `window.electronExcel` para que las pantallas funcionen en Electron y Tauri.
- Tauri llama a comandos Rust.
- Los comandos Rust llaman temporalmente a `tauri-node-backend.js`, que reutiliza el motor actual de Excel de `main.js`.

Este enfoque permite seguir desarrollando con Tauri sin reescribir de golpe toda la manipulacion XML del Excel.

## Requisito pendiente

En este equipo todavia no esta instalado Rust:

- `rustc`: no instalado
- `cargo`: no instalado

Instalalo desde:

```powershell
winget install Rustlang.Rustup
```

Despues cierra y vuelve a abrir Visual Studio Code o PowerShell.

## Comandos

```powershell
npm install
npm run tauri -- info
npm run tauri:dev
npm run tauri:build
```

## Siguiente fase recomendada

Cuando Tauri arranque bien, iremos pasando el motor Excel de JavaScript a Rust por partes:

1. Lectura de alumnos, unidades, RRAA y notas.
2. Escritura XML directa de `DATOS`.
3. Escritura XML directa de hojas `U1`, `U2`, etc.
4. Eliminacion del puente temporal `tauri-node-backend.js`.
