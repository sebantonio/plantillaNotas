# ⚡ QUICK START: Tu .EXE en 5 pasos (15 minutos)

## 🎯 Objetivo
Obtener un único archivo `Plantilla Notas Local.exe` que puedas llevar en un USB y distribuir.

---

## 📋 REQUISITOS PREVIOS

Verifica que tengas instalado:
- ✅ **Node.js** (versión 14+) → https://nodejs.org/
- ✅ **NPM** (viene con Node.js)
- ✅ **Git** (opcional pero recomendado)

### Verificar instalación:
```bash
node --version    # Debe mostrar v14.0.0 o superior
npm --version     # Debe mostrar 6.0.0 o superior
```

---

## 🚀 PASO 1: Preparar el proyecto (2 minutos)

### 1.1 Abre terminal/CMD en la carpeta del proyecto

**En Windows:**
- Abre File Explorer
- Navega a tu carpeta `visor-notas`
- Click derecho → "Abrir terminal aquí" o "Open PowerShell here"

**En Mac/Linux:**
```bash
cd /ruta/a/visor-notas
```

### 1.2 Verificar estructura
```bash
ls -la
# Deberías ver:
# - package.json
# - main.js
# - preload.js
# - index.html
# - gestor-alumnos.html
# - gestor-notas.html
# - etc.
```

Si ves estos archivos, ✅ **estás listo**.

---

## 🖥️ PASO 2: Instalar dependencias (3-5 minutos)

```bash
npm install
```

Esto descargará:
- Electron (~170 MB)
- electron-builder (~30 MB)
- Otros paquetes necesarios

**Espera a que termine.** Verás un mensaje como:
```
added 200+ packages
```

---

## 📦 PASO 3: Crear el .EXE (5-7 minutos)

### Opción A: Archivo PORTABLE (Recomendado)
Un único archivo `.exe` que **no necesita instalación**.

```bash
npm run dist
```

O si tienes permisos, también puedes:
```bash
npx electron-builder --win -c.win.target=portable
```

### Opción B: Instalador NSIS
Para que usuarios instalen como programa normal.

```bash
npx electron-builder --win -c.win.target=nsis
```

### Opción C: Ambos (portable + instalador)
```bash
npm run dist
```

**Espera a que termine.** Verás algo como:
```
Building for Windows
  × signing is not supported on macOS, skipped
  ⨯ Execute vendor/wine/bin/wine64 failed
  (ignorar estos errores, no afectan)
  ✔ Packaged into ./dist/
```

---

## ✅ PASO 4: Localizar tu .EXE

Abre el navegador de archivos y ve a:
```
tu_proyecto/dist/
```

Ahí encontrarás:

### Si hiciste **portable**:
```
Plantilla Notas Local 0.1.0.exe    ← ESTE ES TU ARCHIVO
```

### Si hiciste **NSIS instalador**:
```
Plantilla Notas Local Setup 0.1.0.exe    ← INSTALADOR
```

### Si hiciste **ambos**:
Ambos archivos están ahí.

---

## 🧪 PASO 5: Probar (2 minutos)

### 5.1 Ejecutar localmente
```bash
npm start
```

Esto abre la aplicación en modo desarrollo. Prueba:
- ✅ Cargar un Excel
- ✅ Ver alumnos
- ✅ Hacer cambios
- ✅ Cerrar correctamente

### 5.2 Probar el .EXE generado

1. Navega a `./dist/`
2. Doble-click en `Plantilla Notas Local 0.1.0.exe`
3. Prueba toda la funcionalidad

**¡Si todo funciona, ya tienes tu aplicación empaquetada!**

---

## 📤 PASO 6: Distribuir (Bonus)

### Opción A: Copiar en USB
1. Copia `Plantilla Notas Local 0.1.0.exe` a tu USB
2. Dáselo a tus estudiantes
3. Ellos hacen doble-click y ¡funciona!

### Opción B: Subir a un servidor
```bash
# Ejemplo con Firebase Hosting
firebase deploy --only hosting
```

### Opción C: Crear instalador bonito (Opcional)
El archivo generado ya es un ejecutable, pero si quieres un instalador tradicional:

El archivo `Plantilla Notas Local Setup 0.1.0.exe` (si lo creaste con NSIS) es un instalador que:
- Pide permisos de administrador
- Crea acceso directo en Escritorio
- Se instala en `C:\Program Files\`

---

## 🎁 BONUS: Automatizar con script

### En Windows:
Crea archivo `empaquetar.bat` en tu carpeta raíz con:

```batch
@echo off
npm install
npm run dist
echo Done! Archivos en ./dist/
pause
```

Luego solo necesitas hacer **doble-click** en `empaquetar.bat` para compilar.

### En Mac/Linux:
Crea archivo `empaquetar.sh`:

```bash
#!/bin/bash
npm install
npm run dist
echo "Done! Archivos en ./dist/"
```

Ejecutar:
```bash
chmod +x empaquetar.sh
./empaquetar.sh
```

---

## ⚠️ SOLUCIÓN DE PROBLEMAS

### Error: "npm: command not found"
→ Node.js no está instalado. Descárgalo desde https://nodejs.org/

### Error: "No se encontró package.json"
→ Ejecuta el comando en la carpeta correcta (donde está `main.js`)

### El .EXE no se genera
→ Verifica que tengas:
- Node.js 14+ instalado
- 2 GB de espacio libre en disco
- Antivirus no está bloqueando (desactiva temporalmente)

### El .EXE pesa mucho (150+ MB)
→ Normal. Electron incluye todo el navegador Chrome. Es el trade-off.

### Quiero reducir tamaño
→ Mira la guía completa (`GUIA_TECNICA_ANDROID.md`)

---

## 📊 TIEMPO TOTAL

| Paso | Tiempo | Acción |
|------|--------|--------|
| 1. Preparar | 2 min | Abrir terminal |
| 2. Instalar deps | 5 min | `npm install` |
| 3. Compilar | 5 min | `npm run dist` |
| 4. Localizar | 1 min | Ver en `./dist/` |
| 5. Probar | 2 min | Ejecutar el .exe |
| **TOTAL** | **~15 minutos** | ✅ **Listo para distribuir** |

---

## 🎉 RESULTADO FINAL

Tendrás:
```
Plantilla Notas Local 0.1.0.exe    (10-20 MB)
```

- ✅ Un único archivo
- ✅ Sin instalación requerida
- ✅ Funciona en cualquier Windows
- ✅ Listo para distribuir en USB o email

---

## 📱 Y PARA ANDROID?

Una vez tengas el .EXE listo, cuando quieras hacer versión Android:

1. **Opción fácil** (Capacitor): 1 semana
2. **Opción profesional** (React Native): 2-3 semanas
3. **Opción rápida** (PWA): 2 días

Ver documentación en `GUIA_EMPAQUETADO.md` y `GUIA_TECNICA_ANDROID.md`

---

## ✨ SIGUIENTES PASOS (Después de tener el .EXE)

1. **Mejorar icono**: Cambiar ícono del ejecutable
2. **Agregar updater**: Automático `electron-updater`
3. **Versión móvil**: Migrar a React Native o Capacitor
4. **Publicar**: Subir a GitHub Releases para distribución

---

## 🆘 ¿NECESITAS AYUDA?

- Abre una terminal en la carpeta del proyecto
- Copia el error exacto
- Envíalo junto con output de: `npm list`

---

## 🚀 ¡VAMOS!

```bash
cd tu_carpeta_proyecto
npm install
npm run dist
```

**En 15 minutos tendrás tu aplicación empaquetada.** 

¿Alguna pregunta?
