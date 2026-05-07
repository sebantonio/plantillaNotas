# 🔧 GUÍA TÉCNICA: Migración Android con React Native y Capacitor

## PARTE 1: REACT NATIVE (Opción completa y profesional)

### Arquitectura: Reutilizar lógica, reescribir UI

Tu proyecto Electron tiene esta estructura:
```
main.js             ← Lógica de Electron (IPC, archivo)
preload.js          ← Bridge entre Electron y HTML
HTML (5 archivos)   ← UI para web
JavaScript inline   ← Lógica de datos y Excel
```

En React Native:
```
App.js              ← Componente raíz
/screens            ← UI equivalente a tus HTML
/utils/excelHelper  ← Lógica reutilizable de main.js
/services/fileService ← Equivalente a preload.js
```

---

### Paso 1: Crear el proyecto base

```bash
npx create-expo-app plantilla-notas-mobile
cd plantilla-notas-mobile

npm install \
  expo-file-system \
  expo-document-picker \
  expo-sharing \
  react-native-paper \
  @react-navigation/native \
  @react-navigation/bottom-tabs \
  xlsx \
  jszip
```

---

### Paso 2: Crear servicio para leer Excel (reutilizable)

**`services/excelService.js`**:
```javascript
import * as FileSystem from 'expo-file-system';
import * as DocumentPicker from 'expo-document-picker';
import XLSX from 'xlsx';

class ExcelService {
  constructor() {
    this.selectedPath = null;
    this.workbook = null;
    this.alumnosData = [];
  }

  /**
   * Seleccionar archivo Excel (equivalente a ipcMain.handle)
   */
  async selectExcelFile() {
    try {
      const result = await DocumentPicker.getDocumentAsync({
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });

      if (result.canceled) {
        return null;
      }

      this.selectedPath = result.assets[0].uri;
      return await this.loadExcel(this.selectedPath);
    } catch (error) {
      console.error('Error selecting file:', error);
      throw error;
    }
  }

  /**
   * Cargar Excel desde URI (reutilizable de main.js)
   */
  async loadExcel(filePath) {
    try {
      // Leer como Base64
      const fileContent = await FileSystem.readAsStringAsync(filePath, {
        encoding: FileSystem.EncodingType.Base64,
      });

      // Procesar con XLSX (igual que en Electron)
      this.workbook = XLSX.read(fileContent, { type: 'base64' });
      
      // Extraer datos de alumnos (mismo código que main.js)
      const alumnosSheet = this.workbook.Sheets['alumnos'] || this.workbook.Sheets[0];
      const alumnosData = XLSX.utils.sheet_to_json(alumnosSheet);

      this.alumnosData = alumnosData;
      return alumnosData;
    } catch (error) {
      console.error('Error loading Excel:', error);
      throw error;
    }
  }

  /**
   * Obtener notas de un alumno (lógica extraída de main.js)
   */
  getAlumnoNotas(alumnoId, activityType = 'practicas') {
    const ACTIVITY_TYPES = {
      practicas: { label: 'Practicas', baseCol: 0 },
      memorias: { label: 'Memorias', baseCol: 112 },
      otros: { label: 'Otras actividades', baseCol: 223 },
      controles: { label: 'Control teorico/practico', baseCol: 334 },
    };

    if (!this.workbook) {
      throw new Error('No Excel file loaded');
    }

    const activity = ACTIVITY_TYPES[activityType];
    const sheet = this.workbook.Sheets['notas'];
    
    // Extraer datos desde Excel (columna dinámica basada en baseCol)
    const data = XLSX.utils.sheet_to_json(sheet);
    const alumno = data.find(a => a['ID'] === alumnoId);

    return alumno ? alumno : null;
  }

  /**
   * Guardar cambios en Excel
   */
  async saveChanges() {
    try {
      const wbout = XLSX.write(this.workbook, { bookType: 'xlsx', type: 'base64' });
      
      await FileSystem.writeAsStringAsync(this.selectedPath, wbout, {
        encoding: FileSystem.EncodingType.Base64,
      });

      return true;
    } catch (error) {
      console.error('Error saving Excel:', error);
      throw error;
    }
  }

  /**
   * Exportar datos a ZIP (similar a main.js)
   */
  async exportToZip() {
    try {
      const JSZip = require('jszip');
      const zip = new JSZip();

      // Agregar Excel
      const wbout = XLSX.write(this.workbook, { bookType: 'xlsx', type: 'base64' });
      zip.file('datos.xlsx', wbout, { base64: true });

      // Agregar JSON de alumnos
      zip.file('alumnos.json', JSON.stringify(this.alumnosData, null, 2));

      const zipContent = await zip.generateAsync({ type: 'base64' });

      // Guardar
      const zipUri = `${FileSystem.documentDirectory}export_${Date.now()}.zip`;
      await FileSystem.writeAsStringAsync(zipUri, zipContent, {
        encoding: FileSystem.EncodingType.Base64,
      });

      return zipUri;
    } catch (error) {
      console.error('Error exporting to ZIP:', error);
      throw error;
    }
  }
}

export default new ExcelService();
```

---

### Paso 3: Crear pantalla principal (equivalente a index.html)

**`screens/HomeScreen.js`**:
```javascript
import React, { useState, useEffect } from 'react';
import { View, ScrollView, StyleSheet, Alert } from 'react-native';
import { Button, Card, Title, Paragraph, Appbar } from 'react-native-paper';
import excelService from '../services/excelService';

const HomeScreen = ({ navigation }) => {
  const [alumnosData, setAlumnosData] = useState([]);
  const [loading, setLoading] = useState(false);

  const handleSelectFile = async () => {
    setLoading(true);
    try {
      const data = await excelService.selectExcelFile();
      setAlumnosData(data);
      Alert.alert('Éxito', `Se cargaron ${data.length} alumnos`);
    } catch (error) {
      Alert.alert('Error', 'No se pudo cargar el archivo');
    } finally {
      setLoading(false);
    }
  };

  const handleExport = async () => {
    setLoading(true);
    try {
      const zipUri = await excelService.exportToZip();
      Alert.alert('Éxito', `Archivo exportado: ${zipUri}`);
    } catch (error) {
      Alert.alert('Error', 'No se pudo exportar');
    } finally {
      setLoading(false);
    }
  };

  return (
    <>
      <Appbar.Header>
        <Appbar.Content title="Plantilla Notas Local" />
      </Appbar.Header>

      <ScrollView style={styles.container}>
        <Card style={styles.card}>
          <Card.Content>
            <Title>Gestión de Alumnos</Title>
            <Paragraph>
              Cargue su plantilla Excel para gestionar notas y calificaciones
            </Paragraph>
          </Card.Content>
          <Card.Actions>
            <Button 
              mode="contained" 
              onPress={handleSelectFile}
              loading={loading}
            >
              Seleccionar Excel
            </Button>
          </Card.Actions>
        </Card>

        {alumnosData.length > 0 && (
          <Card style={styles.card}>
            <Card.Content>
              <Title>Alumnos Cargados: {alumnosData.length}</Title>
              <Button onPress={() => navigation.navigate('Alumnos')}>
                Ver Alumnos
              </Button>
            </Card.Content>
          </Card>
        )}

        <Button 
          mode="outlined" 
          onPress={handleExport}
          style={styles.button}
        >
          Exportar a ZIP
        </Button>
      </ScrollView>
    </>
  );
};

const styles = StyleSheet.create({
  container: {
    flex: 1,
    padding: 16,
    backgroundColor: '#f5f5f5',
  },
  card: {
    marginBottom: 16,
  },
  button: {
    marginTop: 16,
  },
});

export default HomeScreen;
```

---

### Paso 4: Crear navegación

**`App.js`**:
```javascript
import React from 'react';
import { NavigationContainer } from '@react-navigation/native';
import { createNativeStackNavigator } from '@react-navigation/native-stack';
import { createBottomTabNavigator } from '@react-navigation/bottom-tabs';
import { PaperProvider } from 'react-native-paper';
import MaterialCommunityIcons from '@expo/vector-icons/MaterialCommunityIcons';

import HomeScreen from './screens/HomeScreen';
import AlumnosScreen from './screens/AlumnosScreen';
import NotasScreen from './screens/NotasScreen';

const Stack = createNativeStackNavigator();
const Tab = createBottomTabNavigator();

function HomeStack() {
  return (
    <Stack.Navigator>
      <Stack.Screen 
        name="Home" 
        component={HomeScreen}
        options={{ title: 'Inicio' }}
      />
      <Stack.Screen 
        name="Alumnos" 
        component={AlumnosScreen}
        options={{ title: 'Alumnos' }}
      />
      <Stack.Screen 
        name="Notas" 
        component={NotasScreen}
        options={{ title: 'Notas del Alumno' }}
      />
    </Stack.Navigator>
  );
}

export default function App() {
  return (
    <PaperProvider>
      <NavigationContainer>
        <Tab.Navigator
          screenOptions={({ route }) => ({
            tabBarIcon: ({ focused, color, size }) => {
              let iconName;
              if (route.name === 'Home') {
                iconName = 'home';
              } else if (route.name === 'Settings') {
                iconName = 'cog';
              }
              return (
                <MaterialCommunityIcons 
                  name={iconName} 
                  size={size} 
                  color={color} 
                />
              );
            },
            tabBarActiveTintColor: '#007AFF',
            tabBarInactiveTintColor: 'gray',
          })}
        >
          <Tab.Screen 
            name="Home" 
            component={HomeStack}
            options={{ headerShown: false }}
          />
          <Tab.Screen 
            name="Settings" 
            component={SettingsScreen}
            options={{ title: 'Configuración' }}
          />
        </Tab.Navigator>
      </NavigationContainer>
    </PaperProvider>
  );
}
```

---

### Paso 5: Compilar para Android

#### **Opción A: Expo (más fácil)**
```bash
# Instalar Expo CLI
npm install -g eas-cli

# Configurar proyecto
eas build --platform android --local

# Se genera: plantilla-notas-mobile.apk
```

#### **Opción B: Android Studio (más control)**
```bash
# Generar proyecto nativo
npx react-native init plantilla-notas-mobile
npx react-native link

# Compilar
cd android
./gradlew assembleRelease

# El APK estará en: android/app/build/outputs/apk/release/app-release.apk
```

---

## PARTE 2: CAPACITOR (Más rápido, menos cambios)

### Idea: Web app + capa nativa

Capacitor permite usar **tu HTML actual** (con ajustes mínimos) en móvil.

---

### Paso 1: Preparar estructura web

Primero, convierte tus múltiples HTML en una **SPA única**:

**`index.html`** (nueva versión):
```html
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Plantilla Notas Local</title>
  <style>
    body { margin: 0; font-family: Arial, sans-serif; }
    .view { display: none; }
    .view.active { display: block; }
    nav { padding: 10px; background: #007AFF; color: white; }
    nav button { background: none; border: none; color: white; cursor: pointer; margin: 0 10px; }
  </style>
</head>
<body>
  <nav>
    <button onclick="showView('home')">Inicio</button>
    <button onclick="showView('alumnos')">Alumnos</button>
    <button onclick="showView('notas')">Notas</button>
  </nav>

  <div id="home" class="view active">
    <h1>Plantilla Notas Local</h1>
    <button onclick="selectFile()">Cargar Excel</button>
  </div>

  <div id="alumnos" class="view">
    <h1>Alumnos</h1>
    <div id="alumnosList"></div>
  </div>

  <div id="notas" class="view">
    <h1>Notas</h1>
    <div id="notasContent"></div>
  </div>

  <script src="js/app.js"></script>
  <script src="js/ui.js"></script>
</body>
</html>
```

**`js/ui.js`**:
```javascript
function showView(viewName) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById(viewName).classList.add('active');
}

function selectFile() {
  // Captura el evento del documento picker
  window.electronAPI.selectFile().then(data => {
    renderAlumnos(data);
    showView('alumnos');
  });
}

function renderAlumnos(alumnos) {
  const list = document.getElementById('alumnosList');
  list.innerHTML = alumnos.map(a => 
    `<div onclick="selectAlumno('${a.id}')">
      <strong>${a.nombre}</strong> - ID: ${a.id}
    </div>`
  ).join('');
}
```

---

### Paso 2: Instalar Capacitor

```bash
npm install @capacitor/core @capacitor/cli
npx cap init

# Cuando pregunte:
# App name: Plantilla Notas Local
# App Package ID: es.seba.plantillanotas
# Web dir: . (punto, porque tus HTML están en la raíz)
```

---

### Paso 3: Agregar plataforma Android

```bash
npm install @capacitor/android
npx cap add android

# Instala herramientas de Android necesarias
```

---

### Paso 4: Adaptar lectura de archivos

Instalar plugin de filesystem:
```bash
npm install @capacitor/filesystem
npm install @capacitor/document-picker
```

**`js/fileService.js`** (con Capacitor):
```javascript
import { Filesystem, Directory, Encoding } from '@capacitor/filesystem';
import { DocumentPicker } from '@capacitor/document-picker';
import XLSX from 'xlsx';

class CapacitorFileService {
  async selectExcelFile() {
    const result = await DocumentPicker.pickFiles({
      types: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
    });

    if (!result.files || result.files.length === 0) {
      return null;
    }

    return this.loadExcelFile(result.files[0].path);
  }

  async loadExcelFile(filePath) {
    const contents = await Filesystem.readFile({
      path: filePath,
      directory: Directory.Documents,
      encoding: Encoding.UTF8,
    });

    const workbook = XLSX.read(contents.data, { type: 'string' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    return data;
  }

  async saveFile(filename, content) {
    await Filesystem.writeFile({
      path: `${filename}`,
      data: content,
      directory: Directory.Documents,
    });
  }
}

export default new CapacitorFileService();
```

---

### Paso 5: Compilar para Android

```bash
# Sincronizar archivos web con Android Studio
npx cap sync android

# Copiar proyecto a Android Studio
npx cap open android
```

En Android Studio:
1. Click en "Build" → "Generate Signed Bundle / APK"
2. Seleccionar "APK"
3. Crear o seleccionar firma
4. Elegir "release"
5. Listo: `app-release.apk`

---

## PARTE 3: PWA (La opción más rápida sin cambios)

Si tu app es principalmente **lectura de datos**, una PWA funciona perfecto.

### Paso 1: Service Worker

**`service-worker.js`**:
```javascript
const CACHE_NAME = 'plantilla-notas-v1';
const FILES = [
  '/',
  '/index.html',
  '/gestor-notas.html',
  '/gestor-alumnos.html',
  '/gestor-rraa-criterios.html',
  '/gestor-unidades.html',
  '/visor-notas.html',
  '/main.js',
  '/preload.js',
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(FILES))
  );
});

self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request).then(response => {
      return response || fetch(event.request).then(response => {
        // Cachear dinámicamente
        if (response.ok && event.request.method === 'GET') {
          const responseClone = response.clone();
          caches.open(CACHE_NAME).then(cache => {
            cache.put(event.request, responseClone);
          });
        }
        return response;
      }).catch(() => {
        return caches.match(event.request);
      });
    })
  );
});
```

### Paso 2: Manifest

**`manifest.json`**:
```json
{
  "name": "Plantilla Notas Local",
  "short_name": "Notas",
  "start_url": "/",
  "display": "standalone",
  "background_color": "#ffffff",
  "theme_color": "#007AFF",
  "orientation": "portrait-primary",
  "icons": [
    {
      "src": "/icon-192.png",
      "sizes": "192x192",
      "type": "image/png",
      "purpose": "any"
    },
    {
      "src": "/icon-512.png",
      "sizes": "512x512",
      "type": "image/png",
      "purpose": "maskable"
    }
  ],
  "screenshots": [
    {
      "src": "/screenshot-540.png",
      "sizes": "540x720",
      "type": "image/png",
      "form_factor": "narrow"
    }
  ]
}
```

### Paso 3: Registrar en index.html

```html
<link rel="manifest" href="/manifest.json">
<link rel="icon" href="/icon-192.png">
<meta name="theme-color" content="#007AFF">

<script>
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('/service-worker.js')
      .then(() => console.log('PWA instalada'))
      .catch(err => console.log('Error instalando PWA', err));
  }
</script>
```

### Paso 4: Desplegar

```bash
# Opción A: GitHub Pages
npm install -g gh-pages
gh-pages -d .

# Opción B: Servidor local
npx http-server .

# Opción C: Firebase Hosting
npm install -g firebase-tools
firebase init
firebase deploy
```

En Android, abrir en Chrome y tocar "Instalar app".

---

## COMPARATIVA DE COMPLEJIDAD

```
┌─────────────────┬──────────┬───────────┬──────────┬─────────┐
│ Opción          │ Tiempo   │ Cambios   │ Tamaño   │ Calidad │
├─────────────────┼──────────┼───────────┼──────────┼─────────┤
│ Electron (dist) │ 5 min    │ Ninguno   │ 150 MB   │ ★★★★★  │
│ React Native    │ 3 semanas│ 80%       │ 60 MB    │ ★★★★★  │
│ Capacitor       │ 1 semana │ 30%       │ 80 MB    │ ★★★★   │
│ PWA             │ 2 días   │ 10%       │ 5 MB     │ ★★★    │
│ Flutter         │ 1 mes    │ 100%      │ 40 MB    │ ★★★★★  │
└─────────────────┴──────────┴───────────┴──────────┴─────────┘
```

---

## RECOMENDACIÓN FINAL

**Para ti (profesor con proyecto ya hecho):**

1. **Hoy**: `npm run dist` → Tienes Windows listo
2. **Mes 1**: Migra a **Capacitor** (30% cambios, 1 semana)
3. **Resultado**: Windows + Android + Web

**Total de esfuerzo**: 2-3 semanas para soporte multiplataforma.

¿Necesitas ayuda con alguno de estos pasos?
