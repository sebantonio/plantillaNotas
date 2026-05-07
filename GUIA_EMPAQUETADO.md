# 📦 GUÍA COMPLETA: Empaquetado de Aplicación Electron a Desktop y Android

## 📋 RESUMEN EJECUTIVO

Tu proyecto **"Plantilla Notas Local"** es una aplicación **Electron** (Node.js + Chrome) que gestiona alumnos y notas desde Excel. Ya tiene configurado `electron-builder` en `package.json`, lo que significa que **puedes empaquetar para Windows en un único `.exe` con solo un comando**.

Para **Android**, necesitarás migrar a un framework multiplataforma como **React Native**, **Tauri**, o **Flutter**.

---

## 🖥️ OPCIÓN 1: EMPAQUETADO PARA WINDOWS (ELECTRON)

### ✅ Lo bueno: Ya está casi todo hecho
- Tu `package.json` ya tiene `electron-builder` configurado
- Ya genera NSIS (instalador) y portable (ejecutable único)

### 📦 Paso a paso: Crear el archivo .EXE único

#### **1. Verificar dependencias instaladas**
```bash
npm install
```

#### **2. Crear el ejecutable PORTABLE (recomendado)**
```bash
npm run dist
```

Esto genera en la carpeta `dist/`:
- `Plantilla Notas Local Setup 0.1.0.exe` (instalador NSIS)
- `Plantilla Notas Local 0.1.0.exe` (portable, sin instalación)

**El portable es lo que buscas**: 1 archivo, cópialo y listo.

#### **3. Optimizar el tamaño (opcional pero recomendado)**

Ahora mismo Electron incluye todo: Node, Chrome, etc. (~150-200 MB). Para reducirlo:

**En `package.json`, agregar en la sección `build.win`:**
```json
"build": {
  "win": {
    "target": [
      "nsis",
      "portable"
    ],
    "certificateFile": null,
    "certificatePassword": null
  },
  "nsis": {
    "oneClick": false,
    "allowToChangeInstallationDirectory": true,
    "createDesktopShortcut": true,
    "createStartMenuShortcut": true,
    "installerIcon": "path/to/icon.ico",
    "uninstallerIcon": "path/to/icon.ico"
  },
  "files": [
    "index.html",
    "gestor-alumnos.html",
    "gestor-rraa-criterios.html",
    "gestor-unidades.html",
    "gestor-notas.html",
    "visor-notas.html",
    "main.js",
    "preload.js",
    "package.json"
  ]
}
```

---

## 📱 OPCIÓN 2: MIGRACIÓN A ANDROID (Multiplataforma)

### ⚠️ La realidad: No hay solución directa Electron → Android

Electron **solo funciona en Windows, Mac, Linux**. No existe versión nativa para Android.

### 🎯 Las mejores alternativas:

| Opción | Esfuerzo | Código reutilizable | Recomendación |
|--------|----------|-------------------|---------------|
| **React Native** | Medio (70% del código) | Sí, HTML+JS reutilizable | ✅ Mejor para tu caso |
| **Flutter** | Alto (reescribir todo) | No mucho | Si quieres máximo rendimiento |
| **Tauri** | Medio | Sí, pero sin Android | Solo Windows/Mac/Linux |
| **Capacitor** | Medio | Sí, reutiliza web | Buena para híbridas |
| **PWA + Web** | Bajo | 100% reutilizable | Solución más rápida |

---

## 🚀 OPCIÓN 2A: React Native (RECOMENDADA)

### Ventajas:
- ✅ Reutilizas **70-80% del código JavaScript**
- ✅ Funciona en iOS y Android
- ✅ Acceso nativo (lectura/escritura de archivos Excel)
- ✅ Mejor rendimiento que web app

### Desventajas:
- ❌ UI requiere reescritura con componentes nativos
- ❌ Manejo de Excel más complejo que en Node.js

### Camino de migración:

#### **1. Crear proyecto React Native**
```bash
npx create-expo-app plantilla-notas-mobile
cd plantilla-notas-mobile
npm install xlsx jszip
```

#### **2. Instalar librerías clave para Excel**
```bash
npm install expo-document-picker expo-file-system
npm install xlsx jszip
```

#### **3. Estructura del código (Pseudocódigo)**

**`App.js`** - Lógica reutilizable de Excel:
```javascript
import * as FileSystem from 'expo-file-system';
import * as DocumentPicker from 'expo-document-picker';
import XLSX from 'xlsx';

const selectAndLoadExcel = async () => {
  const doc = await DocumentPicker.getDocumentAsync({ type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const fileContent = await FileSystem.readAsStringAsync(doc.uri, {encoding: FileSystem.EncodingType.Base64});
  const workbook = XLSX.read(fileContent, {type: 'base64'});
  // ... procesar datos igual que en Electron
};
```

#### **4. Compilar para Android**
```bash
eas build --platform android
# o
expo build:android
```

**Resultado**: `.apk` o `.aab` para instalar en Android.

---

## 🚀 OPCIÓN 2B: Capacitor (Alternativa + simple)

Convierte tu **app web actual** a móvil sin prácticamente cambios.

### Ventajas:
- ✅ Código web reutilizable 100%
- ✅ Menos reescritura que React Native
- ✅ Acceso a archivos nativos

### Desventajas:
- ❌ Rendimiento inferior a React Native
- ❌ Limitaciones en APIs nativas complejas

### Paso a paso:

#### **1. Preparar tu proyecto como SPA (Single Page App)**
Tu proyecto Electron ya es casi una SPA. Necesitas:
- Mover todo a **una única `index.html`** con navegación interna
- O crear rutas con un router JavaScript simple

#### **2. Agregar Capacitor**
```bash
npm install @capacitor/core @capacitor/cli
npx cap init
# Nombre: "Plantilla Notas Local"
# Package ID: es.seba.plantillanotas
```

#### **3. Agregar plataformas**
```bash
npm install @capacitor/android @capacitor/ios
npx cap add android
npx cap add ios
```

#### **4. Adaptar para lectura de Excel**
```bash
npm install @capacitor/filesystem
```

**Código (JavaScript):**
```javascript
import { Filesystem, Directory, Encoding } from '@capacitor/filesystem';

const readExcelFile = async (fileUri) => {
  const contents = await Filesystem.readFile({
    path: fileUri,
    directory: Directory.Documents,
    encoding: Encoding.UTF8
  });
  const workbook = XLSX.read(contents);
  // ... procesar
};
```

#### **5. Compilar**
```bash
npx cap build android
# Se abre Android Studio automáticamente
```

---

## 🌐 OPCIÓN 2C: PWA (Progressive Web App) - LA MÁS RÁPIDA

Si necesitas **algo rápido** sin reescribir, convierte tu app a **PWA** que funcione offline.

### Ventajas:
- ✅ **Cero cambios de código**
- ✅ Funciona en cualquier navegador
- ✅ Instalable como app en móvil desde Chrome
- ✅ Funciona offline

### Desventajas:
- ❌ No es una "app real" (es web)
- ❌ Menos acceso a APIs nativas

### Setup mínimo:

#### **1. Agregar Service Worker**
Crea `service-worker.js`:
```javascript
const CACHE_NAME = 'plantilla-notas-v1';
const urlsToCache = [
  '/',
  '/index.html',
  '/gestor-notas.html',
  '/gestor-alumnos.html',
  // ... agregar todos tus HTML
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(urlsToCache))
  );
});

self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request).then(response => response || fetch(event.request))
  );
});
```

#### **2. Registrar Service Worker en `index.html`**
```html
<script>
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('/service-worker.js');
  }
</script>
```

#### **3. Agregar `manifest.json`**
```json
{
  "name": "Plantilla Notas Local",
  "short_name": "Notas",
  "start_url": "/",
  "display": "standalone",
  "background_color": "#ffffff",
  "theme_color": "#0078D4",
  "icons": [
    {
      "src": "/icon-192.png",
      "sizes": "192x192",
      "type": "image/png"
    },
    {
      "src": "/icon-512.png",
      "sizes": "512x512",
      "type": "image/png"
    }
  ]
}
```

#### **4. Linkar en `index.html`**
```html
<link rel="manifest" href="/manifest.json">
```

#### **5. Desplegar como web**
```bash
npm install -g serve
serve -s . -p 3000
# Luego abrir en Android: https://tu-ip:3000
# Chrome sugiere "Instalar app"
```

---

## 📊 COMPARATIVA: ¿Cuál elegir?

| Escenario | Solución | Razón |
|-----------|----------|-------|
| **Solo Windows, 1 archivo** | Electron `npm run dist` | ✅ Lista ya |
| **Windows + Android, máx compatibilidad** | React Native | ✅ Recomendado |
| **Windows + Android, máx rapidez** | Capacitor | ✅ Menos cambios |
| **Web + Android, sin instalación** | PWA | ✅ Más simple |
| **Máximo rendimiento Android** | Flutter | ✅ Si reescribes todo |

---

## 🛠️ PLAN DE ACCIÓN RECOMENDADO

### **Fase 1: Empaquetar Windows (Hoy)**
```bash
# 1. Verificar todo está ok
npm install

# 2. Crear portable
npm run dist

# 3. El archivo estará en ./dist/
# Ejemplo: "Plantilla Notas Local 0.1.0.exe"
```

**Resultado**: 1 archivo `.exe` para distribución.

### **Fase 2: Preparar Android (Próximas semanas)**

**Opción A (Recomendada): React Native**
```bash
# 1. Duplicar proyecto
cp -r . ../plantilla-notas-mobile

# 2. Reescribir UI con React Native
# (70% del código JavaScript reutilizable)

# 3. Compilar para Android
eas build --platform android
```

**Opción B (Más rápida): Capacitor**
```bash
# 1. Instalar Capacitor
npm install @capacitor/core @capacitor/cli

# 2. Adaptar estructu web
npx cap init

# 3. Agregar Android
npx cap add android

# 4. Compilar
npx cap build android
```

---

## 📝 CONFIGURACIÓN FINAL RECOMENDADA PARA `package.json`

```json
{
  "name": "plantilla-notas-local",
  "version": "1.0.0",
  "description": "Aplicación para gestionar notas de alumnos con Excel",
  "main": "main.js",
  "scripts": {
    "start": "electron .",
    "pack": "electron-builder --dir",
    "dist": "electron-builder",
    "dist:win": "electron-builder --win --publish never",
    "dist:portable": "electron-builder -c.win.target=portable --win --publish never"
  },
  "build": {
    "appId": "es.seba.plantillanotas",
    "productName": "Plantilla Notas Local",
    "files": [
      "index.html",
      "gestor-alumnos.html",
      "gestor-rraa-criterios.html",
      "gestor-unidades.html",
      "gestor-notas.html",
      "visor-notas.html",
      "main.js",
      "preload.js",
      "package.json"
    ],
    "win": {
      "target": [
        {
          "target": "nsis",
          "arch": ["x64"]
        },
        {
          "target": "portable",
          "arch": ["x64"]
        }
      ],
      "certificateFile": null,
      "certificatePassword": null,
      "signingHashAlgorithms": ["sha256"]
    },
    "nsis": {
      "oneClick": false,
      "allowToChangeInstallationDirectory": true,
      "createDesktopShortcut": true,
      "createStartMenuShortcut": true,
      "shortcutName": "Plantilla Notas Local"
    }
  },
  "dependencies": {
    "jszip": "^3.10.1",
    "xlsx": "^0.18.5"
  },
  "devDependencies": {
    "electron": "^30.5.1",
    "electron-builder": "^24.13.3"
  }
}
```

---

## ⚡ DISTRIBUCIÓN Y ACTUALIZACIONES

### Para Windows:
1. Distribuir `.exe` portable por email o web
2. O usar **electron-updater** para actualizaciones automáticas:
```bash
npm install electron-updater
```

### Para Android:
1. Publicar en **Google Play Store** (requiere cuenta de desarrollador, ~$25)
2. O distribuir `.apk` directamente (usuarios instalan de "fuentes desconocidas")

---

## ❓ PREGUNTAS FRECUENTES

**P: ¿El Excel tiene que estar empaquetado en el .exe?**
R: No, el usuario selecciona su Excel en tiempo de ejecución. Pero puedes empaquetar archivos por defecto.

**P: ¿Qué pesa más: Electron o React Native?**
R: Electron (~150-200 MB), React Native (~50-80 MB).

**P: ¿Puedo hacer lo mismo con .NET?**
R: Sí, con WPF o UWP, pero pierdes portabilidad a Android sin bibliotecas externas.

**P: ¿Se puede hacer un único APK para Android?**
R: Sí, un archivo `.apk` es el equivalente a tu `.exe`.

---

## 📚 RECURSOS Y ENLACES

- Electron Builder: https://www.electron.build/
- React Native: https://reactnative.dev/
- Capacitor: https://capacitorjs.com/
- Tauri: https://tauri.app/
- PWA: https://developer.mozilla.org/en-US/docs/Web/Progressive_web_apps/

---

## 📌 SIGUIENTE PASO

**Mañana mismo puedes:**
1. Ejecutar `npm install` (si no lo hiciste)
2. Ejecutar `npm run dist`
3. Encontrar tu `.exe` en `./dist/`

**Tu aplicación en 1 solo archivo = LISTO.**

¿Necesitas ayuda con algún paso específico?
