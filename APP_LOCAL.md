# App local con Electron

Esta version permite abrir el frontend como una aplicacion de escritorio y guardar cambios directamente en un Excel del ordenador.

Para una explicacion exhaustiva del proceso, arquitectura, comandos, empaquetado `.exe` y errores frecuentes, consulta `DOCUMENTACION_APP_LOCAL_ELECTRON.md`.

## Que necesitas

- Node.js instalado en Windows.
- El archivo Excel cerrado en Excel mientras la app escribe cambios.
- Si el Excel esta en una carpeta de OneDrive sincronizada, OneDrive subira los cambios despues de guardar.

## Arrancar en desarrollo

```bash
npm install
npm start
```

Tambien puedes abrirla con doble clic en `Abrir Gestor Local.bat`.

## Crear una aplicacion instalable

Ejecuta:

```bash
npm install
npm run dist
```

El resultado aparecera en la carpeta `dist/`. Se generara un instalador de Windows y una version portable.

## Flujo de uso

1. Ejecuta la app con `npm start`.
2. Abre `Gestion de Alumnos`.
3. Pulsa en el area de carga y selecciona el Excel local.
4. Agrega o elimina alumnos.
5. Pulsa `Guardar cambios en el Excel`.

## Importante

La app no usa Microsoft Graph ni login de Microsoft. Escribe en el archivo local que selecciones. Si ese archivo esta dentro de OneDrive, la sincronizacion la hace OneDrive Desktop.
