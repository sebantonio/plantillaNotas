# Documentacion de la app local con Electron

Esta documentacion explica que se ha construido, por que se ha construido asi y como convertirlo mas adelante en una aplicacion `.exe` de Windows.

## 1. Objetivo del cambio

El proyecto empezo como una web estatica formada por archivos HTML:

- `index.html`
- `gestor-alumnos.html`
- `gestor-rraa-criterios.html`

En ese modo, el navegador puede leer un Excel seleccionado por el usuario, modificarlo en memoria y descargar una copia nueva. El problema es que una web abierta en el navegador no puede escribir directamente sobre un archivo del ordenador por seguridad.

El nuevo objetivo es:

1. Mantener la interfaz HTML existente.
2. Abrirla como una aplicacion de escritorio.
3. Seleccionar un Excel local del ordenador.
4. Insertar/modificar alumnos.
5. Guardar directamente en el mismo archivo Excel.
6. Si el archivo esta dentro de una carpeta sincronizada de OneDrive, dejar que OneDrive sincronice los cambios.

Para eso se ha incorporado Electron.

## 2. Que es Electron

Electron es una tecnologia que permite crear aplicaciones de escritorio usando tecnologias web:

- HTML para la estructura.
- CSS para el diseno.
- JavaScript para la interaccion.
- Node.js para acceder a funciones del sistema operativo, como leer y escribir archivos.

Una forma sencilla de entenderlo:

- Una web normal vive dentro del navegador y tiene permisos limitados.
- Una app Electron abre una ventana propia, parecida a un navegador integrado.
- Dentro de esa ventana se carga tu `index.html`.
- Ademas, Electron tiene una parte local con permisos para hablar con Windows.

En este proyecto:

- La interfaz sigue siendo `index.html` y `gestor-alumnos.html`.
- Electron abre esos HTML dentro de una ventana de escritorio.
- El archivo `main.js` se encarga de leer y escribir el Excel local.
- El archivo `preload.js` crea un puente seguro entre la interfaz y Electron.

## 3. Diferencia entre modo navegador y modo Electron

### Modo navegador

Se activa cuando abres directamente:

```text
index.html
```

En este modo:

- Se abre en Chrome, Edge, Firefox, etc.
- El gestor usa la libreria SheetJS cargada desde internet.
- Puedes cargar un Excel.
- Puedes modificar datos.
- Al final se descarga una copia.
- No se puede sobrescribir el Excel original.

### Modo Electron

Se activa cuando arrancas con:

```powershell
npm start
```

O cuando haces doble clic en:

```text
Abrir Gestor Local.bat
```

En este modo:

- Se abre una ventana de aplicacion.
- La interfaz detecta `window.electronExcel`.
- Al cargar Excel se abre un selector nativo de archivos.
- Al guardar se escribe directamente en el archivo local seleccionado.
- No se descarga una copia.

## 4. Archivos anadidos o modificados

### `package.json`

Archivo de configuracion del proyecto Node/Electron.

Define:

- Nombre de la app.
- Version.
- Archivo principal: `main.js`.
- Comandos disponibles.
- Dependencias necesarias.
- Configuracion de empaquetado para generar `.exe`.

Comandos importantes:

```json
"start": "electron ."
"pack": "electron-builder --dir"
"dist": "electron-builder"
```

Significado:

- `npm start`: abre la app en modo desarrollo.
- `npm run pack`: crea una carpeta con la app desempaquetada.
- `npm run dist`: crea instalador y version portable.

### `main.js`

Es el proceso principal de Electron.

Responsabilidades:

- Crear la ventana de la aplicacion.
- Cargar `index.html`.
- Abrir el selector de archivos de Windows.
- Recordar la ruta del Excel seleccionado.
- Leer la hoja `DATOS`.
- Buscar la seccion `ALUMNADO`.
- Devolver la lista de alumnos al HTML.
- Guardar cambios directamente en el Excel.
- Abrir enlaces externos en el navegador.

Puntos clave:

- Usa `BrowserWindow` para crear la ventana.
- Usa `dialog.showOpenDialog` para seleccionar el Excel.
- Usa `ipcMain.handle` para recibir peticiones desde la interfaz.
- Usa `xlsx` para leer y escribir el libro Excel.

### `preload.js`

Es el puente seguro entre la interfaz HTML y Electron.

Electron no expone directamente Node.js dentro de la pagina, por seguridad. En su lugar, `preload.js` publica una pequena API controlada:

```js
window.electronExcel.selectFile()
window.electronExcel.getSelectedFile()
window.electronExcel.saveAlumnos(alumnos)
window.electronExcel.openExternal(url)
```

La interfaz solo puede hacer esas operaciones, no acceder libremente al sistema.

### `gestor-alumnos.html`

Se ha adaptado para funcionar en dos modos:

- Modo navegador.
- Modo Electron.

La deteccion se hace con:

```js
const isElectronMode = Boolean(window.electronExcel);
```

Si esta en Electron:

- No depende de la carga de `XLSX` en el navegador.
- El area de carga abre el selector local.
- El boton final cambia a `Guardar cambios en el Excel`.
- El guardado llama a `window.electronExcel.saveAlumnos(...)`.

Si no esta en Electron:

- Sigue funcionando como antes.
- Carga el Excel en el navegador.
- Descarga una copia actualizada.

### `Abrir Gestor Local.bat`

Archivo para abrir la app con doble clic durante desarrollo.

Contenido:

```bat
@echo off
cd /d "%~dp0"
npm start
```

Hace dos cosas:

1. Se coloca en la carpeta donde esta el `.bat`.
2. Ejecuta `npm start`.

Importante: este `.bat` necesita que Node.js y npm esten instalados.

### `APP_LOCAL.md`

Guia breve de uso de la app local.

### `DOCUMENTACION_APP_LOCAL_ELECTRON.md`

Este documento.

## 5. Requisitos para trabajar en este proyecto

En el equipo donde quieras desarrollar o empaquetar la app necesitas:

1. Windows.
2. Node.js instalado.
3. npm disponible en PowerShell.
4. Conexion a internet la primera vez para instalar dependencias.
5. El proyecto completo en una carpeta local.

Node.js se descarga desde:

```text
https://nodejs.org/
```

Recomendado:

- Instalar la version LTS.
- Dejar marcada la opcion de anadir Node/npm al PATH.
- Cerrar y abrir PowerShell despues de instalar.

Comprobar instalacion:

```powershell
node -v
npm -v
```

Si ambos comandos devuelven una version, Node esta bien instalado.

## 6. Como abrir la app en desarrollo

Desde PowerShell:

```powershell
cd "C:\OneDrive\OneDrive - Consejería de Educación, Cultura y Deportes Castilla La-Mancha\Github\plantillaNotas\plantillaNotas"
npm start
```

Tambien puedes hacer doble clic en:

```text
Abrir Gestor Local.bat
```

Si haces doble clic en `index.html`, se abrira el navegador normal. Eso no es Electron.

## 7. Primer arranque en un PC nuevo

Cuando copies el proyecto a otro ordenador, puede que no exista la carpeta `node_modules`.

`node_modules` contiene las librerias descargadas:

- Electron.
- SheetJS/XLSX.
- Electron Builder.
- Otras dependencias internas.

En un PC nuevo, ejecuta:

```powershell
cd "RUTA\A\LA\CARPETA\plantillaNotas"
npm install
npm start
```

`npm install` lee `package.json` y descarga todo lo necesario.

## 8. Flujo de uso de la app local

1. Abre la app con `npm start` o con `Abrir Gestor Local.bat`.
2. En el panel principal, pulsa `Abrir Gestor de Alumnos`.
3. En el gestor, pulsa el area de seleccion de Excel.
4. Selecciona el archivo `.xlsx` local.
5. Comprueba que se cargan los alumnos.
6. Agrega un alumno nuevo.
7. Pulsa `Guardar cambios en el Excel`.
8. Abre el Excel para verificar el resultado.

Recomendacion importante:

Mantener el Excel cerrado mientras guardas desde la app. Si el archivo esta abierto en Excel, Windows puede bloquearlo o OneDrive puede crear un conflicto de sincronizacion.

## 9. Como funciona la insercion de alumnos

La app espera encontrar una hoja llamada:

```text
DATOS
```

Dentro de esa hoja busca una fila donde la columna B contenga:

```text
ALUMNADO
```

Desde la fila siguiente empieza a leer alumnos:

- Columna A: numero.
- Columna B: nombre del alumno.
- Columna C: fecha de nacimiento.

Al guardar:

1. Lee de nuevo el Excel desde disco.
2. Busca otra vez la seccion `ALUMNADO`.
3. Limpia las filas de alumnos existentes.
4. Escribe la lista actual.
5. Guarda el archivo en la misma ruta.

## 10. Limitacion tecnica actual sobre formatos

Actualmente se usa la libreria `xlsx`.

Funciona bien para leer y escribir datos, pero puede no preservar perfectamente:

- Estilos.
- Formatos complejos.
- Validaciones.
- Comentarios.
- Algunas formulas o metadatos avanzados.

El codigo intenta conservar propiedades especiales de la hoja que empiezan por `!`, como rangos o anchos, pero reconstruye la hoja desde arrays.

Si mas adelante la plantilla real depende mucho de formato, formulas o validaciones, conviene estudiar una alternativa:

- `exceljs`
- `xlsx-populate`
- Backend Python con `openpyxl`

Para el primer prototipo local, `xlsx` es suficiente para validar el flujo.

## 11. Como crear el `.exe`

La herramienta configurada para crear el `.exe` es:

```text
electron-builder
```

Ya esta declarada en `package.json`:

```json
"electron-builder": "^24.13.3"
```

### Paso 1. Abrir PowerShell

Abre PowerShell en la carpeta del proyecto:

```powershell
cd "C:\OneDrive\OneDrive - Consejería de Educación, Cultura y Deportes Castilla La-Mancha\Github\plantillaNotas\plantillaNotas"
```

### Paso 2. Instalar dependencias

```powershell
npm install
```

Esto descarga Electron, XLSX y electron-builder.

### Paso 3. Probar antes de empaquetar

```powershell
npm start
```

Antes de generar el `.exe`, conviene comprobar:

- Que abre la ventana.
- Que se puede entrar al gestor de alumnos.
- Que se puede seleccionar un Excel.
- Que se puede guardar en el Excel.

### Paso 4. Generar instalador y portable

```powershell
npm run dist
```

Esto ejecuta:

```powershell
electron-builder
```

### Paso 5. Revisar la carpeta `dist`

Al terminar deberia aparecer una carpeta:

```text
dist/
```

Dentro deberias encontrar archivos similares a:

```text
Plantilla Notas Local Setup 0.1.0.exe
Plantilla Notas Local 0.1.0.exe
win-unpacked/
```

Puede variar el nombre exacto segun la version y la configuracion.

Significado:

- `Setup ... .exe`: instalador normal de Windows.
- `... portable.exe`: version portable si electron-builder la genera con ese nombre.
- `win-unpacked/`: app desempaquetada, util para pruebas.

## 12. Diferencia entre instalador y portable

### Instalador

Ventajas:

- Crea acceso directo en escritorio.
- Crea entrada en menu Inicio.
- Se puede desinstalar desde Windows.
- Mejor para uso habitual.

Inconvenientes:

- Instala archivos en el sistema del usuario.
- Puede requerir permisos segun el equipo.

### Portable

Ventajas:

- Es un `.exe` que se puede ejecutar sin instalar.
- Facil de copiar a otro ordenador.
- Muy util para pruebas.

Inconvenientes:

- No crea accesos directos automaticamente.
- Puede ser menos comodo para usuarios finales.

## 13. Configuracion actual de empaquetado

En `package.json` se ha anadido:

```json
"build": {
  "appId": "es.seba.plantillanotas",
  "productName": "Plantilla Notas Local",
  "files": [
    "index.html",
    "gestor-alumnos.html",
    "gestor-rraa-criterios.html",
    "main.js",
    "preload.js",
    "package.json"
  ],
  "win": {
    "target": [
      "nsis",
      "portable"
    ]
  },
  "nsis": {
    "oneClick": false,
    "allowToChangeInstallationDirectory": true,
    "createDesktopShortcut": true,
    "createStartMenuShortcut": true
  }
}
```

Explicacion:

- `appId`: identificador interno de la aplicacion.
- `productName`: nombre visible de la app.
- `files`: archivos que se meten dentro de la aplicacion.
- `nsis`: genera instalador Windows.
- `portable`: genera ejecutable portable.
- `createDesktopShortcut`: crea acceso directo en escritorio.
- `createStartMenuShortcut`: crea acceso en menu Inicio.

## 14. Que copiar a otro PC

Opcion A: copiar el proyecto para seguir desarrollando.

Copiar:

- Todo el repositorio.
- No es obligatorio copiar `node_modules`.
- Si no copias `node_modules`, ejecuta `npm install` en el nuevo PC.

Despues:

```powershell
npm install
npm start
```

Opcion B: copiar solo la app ya generada.

Despues de `npm run dist`, copiar:

- El instalador `.exe`, o
- El portable `.exe`, o
- La carpeta `win-unpacked`.

En ese caso, el otro PC no necesita Node.js para ejecutar la app ya empaquetada.

## 15. Errores frecuentes

### `npm : El termino 'npm' no se reconoce`

Significa que npm no esta instalado o no esta en el PATH.

Solucion:

1. Instalar Node.js LTS desde `https://nodejs.org/`.
2. Cerrar PowerShell.
3. Abrir PowerShell otra vez.
4. Comprobar:

```powershell
node -v
npm -v
```

### La app se abre en navegador al pulsar `index.html`

Es normal. `index.html` siempre se abre con navegador.

Para Electron:

```powershell
npm start
```

O doble clic en:

```text
Abrir Gestor Local.bat
```

### `npm install` tarda mucho

Normal. Electron descarga bastante contenido.

Recomendaciones:

- Tener conexion estable.
- Evitar carpetas con sincronizacion agresiva si se vuelve muy lento.
- Si OneDrive bloquea archivos de `node_modules`, mover el proyecto temporalmente a una carpeta local como `C:\dev\plantillaNotas`.

### Error al guardar Excel

Posibles causas:

- El Excel esta abierto en Microsoft Excel.
- OneDrive esta sincronizando el archivo.
- No tienes permisos de escritura.
- El archivo esta en modo solo lectura.
- La hoja `DATOS` no existe.
- No existe la seccion `ALUMNADO`.

Soluciones:

1. Cerrar Excel.
2. Esperar a que OneDrive termine.
3. Copiar el archivo a una carpeta local de prueba.
4. Verificar que la hoja se llama exactamente `DATOS`.
5. Verificar que la columna B contiene una celda con `ALUMNADO`.

### Windows SmartScreen avisa al abrir el `.exe`

Es normal en aplicaciones no firmadas.

Motivo:

- El `.exe` no tiene firma digital de una entidad certificadora.
- Windows no conoce aun la app.

Soluciones:

- Para uso personal, se puede permitir manualmente.
- Para distribuir a terceros, habria que firmar el ejecutable con un certificado de firma de codigo.

### El antivirus bloquea el `.exe`

Puede ocurrir con apps Electron generadas localmente.

Soluciones:

- Crear una exclusion temporal durante pruebas.
- Firmar el ejecutable si se va a distribuir.
- Generar el instalador en un entorno limpio.

### OneDrive crea conflictos

Puede pasar si:

- El Excel esta abierto en varios equipos.
- La app guarda mientras OneDrive sincroniza.
- Otro usuario edita el mismo archivo en la nube.

Recomendaciones:

- Cerrar Excel antes de guardar.
- Esperar a que OneDrive este en estado sincronizado.
- Evitar editar el mismo archivo simultaneamente desde Excel web y desde la app local.

## 16. Buenas practicas antes de empaquetar

Antes de ejecutar `npm run dist`:

1. Probar `npm start`.
2. Probar con una copia del Excel real.
3. Agregar un alumno.
4. Eliminar un alumno.
5. Guardar.
6. Abrir el Excel y comprobar los cambios.
7. Probar que OneDrive sincroniza.
8. Cerrar y reabrir la app.
9. Comprobar que no se esta usando accidentalmente el navegador.

## 17. Posibles mejoras futuras

### Recordar el ultimo Excel usado

Ahora la app guarda la ruta en memoria mientras esta abierta.

Mejora:

- Guardar la ruta en una configuracion local.
- Al abrir la app, cargar automaticamente el ultimo Excel.

### Mejorar preservacion de formato

Evaluar cambio de libreria si la plantilla pierde estilos o formulas.

### Crear una pantalla especifica de configuracion

Podria incluir:

- Ruta actual del Excel.
- Boton `Cambiar archivo`.
- Estado de guardado.
- Advertencia si el archivo esta en OneDrive.

### Crear icono de la app

Electron Builder permite anadir un icono `.ico`.

Se configuraria en `package.json`:

```json
"win": {
  "icon": "assets/icon.ico"
}
```

### Firmar el ejecutable

Para distribucion formal, firmar el `.exe` evita muchos avisos de Windows.

Necesita:

- Certificado de firma de codigo.
- Configuracion adicional en `electron-builder`.

## 18. Resumen rapido de comandos

Instalar dependencias:

```powershell
npm install
```

Abrir app en desarrollo:

```powershell
npm start
```

Crear carpeta desempaquetada:

```powershell
npm run pack
```

Crear instalador y portable:

```powershell
npm run dist
```

Comprobar Node y npm:

```powershell
node -v
npm -v
```

## 19. Estado actual

La base Electron esta creada.

Ya existen:

- `package.json`
- `main.js`
- `preload.js`
- `Abrir Gestor Local.bat`
- Adaptacion de `gestor-alumnos.html`

Pendiente de validar en profundidad:

- Guardado sobre una copia real de la plantilla.
- Conservacion de formato y formulas.
- Generacion del `.exe` con `npm run dist`.
- Prueba del instalador en otro PC.

## 20. Idea principal para recordar

Hay dos mundos:

- Navegador: abre `index.html`, trabaja con copias descargadas.
- Electron: abre con `npm start` o `.exe`, trabaja con archivos locales reales.

Si quieres que guarde directamente en el Excel del ordenador, tienes que usar Electron o el `.exe`, no abrir el HTML en el navegador.
