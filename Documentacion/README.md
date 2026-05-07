# 📋 Plantilla Notas Local - Gestor de Evaluación FP

**Versión**: 2.0.0  
**Estado**: ✅ Producción  
**Tecnología**: Tauri v2 + Node.js + XLSX  
**Última actualización**: 2026-05-07

---

## 🎯 ¿Qué es Plantilla Notas Local?

Sistema profesional de gestión de calificaciones para módulos de **Formación Profesional** que funciona como **aplicación de escritorio local** con sincronización de archivos Excel.

### Características principales:

✅ **Gestión completa de evaluación**
- Alumnos y calificaciones
- Resultados de Aprendizaje (RRAA)
- Criterios de Evaluación (CE)
- Ponderaciones por unidad

✅ **Sincronización Excel** 
- Carga/descarga de ficheros `.xlsx`
- Formato compatible con OneDrive
- Preserva estructura original

✅ **Interfaz intuitiva**
- Panel de inicio con navegación
- Múltiples gestores especializados
- Diseño responsivo
- Sistema de pestañas

✅ **Exportación de reportes**
- Notas por alumno
- Análisis por unidad
- Visualización de evaluaciones

---

## 📁 Estructura del Proyecto

```
plantillaNotas/
│
├── 📄 HTML Principales (Gestores)
│   ├── index.html                      ← Panel de inicio
│   ├── gestor-alumnos.html             ← Gestión de alumnos
│   ├── gestor-rraa-criterios.html      ← RRAA y Criterios ⭐
│   ├── gestor-unidades.html            ← Unidades didácticas
│   ├── gestor-notas.html               ← Calificaciones
│   ├── visor-notas.html                ← Visualización notas
│   └── informes.html                   ← Reportes exportables
│
├── 🔧 Backend & Integration
│   ├── main.js                         ← Backend Electron/Node.js
│   ├── preload.js                      ← Bridge seguro Electron
│   ├── app-bridge.js                   ← API Tauri ↔ Frontend
│   └── tauri-node-backend.js           ← Logica compartida
│
├── 📦 Tauri (Framework nativo)
│   └── src-tauri/
│       ├── src/
│       │   ├── main.rs                 ← Punto entrada Rust
│       │   └── build.rs                ← Build script
│       ├── Cargo.toml                  ← Deps Rust
│       └── tauri.conf.json             ← Config Tauri
│
├── 📚 Documentación
│   └── Documentacion/
│       ├── README.md                   ← Este archivo
│       ├── TAURI_MIGRACION.md          ← Detalles migración
│       ├── MAPA_PROYECTO.md            ← Guía visual
│       ├── CHECKLIST_SETUP.md          ← Setup paso a paso
│       └── INSTRUCCIONES_GITHUB.md     ← Git/GitHub
│
├── 📋 Config
│   ├── package.json                    ← Deps Node.js
│   └── .gitignore                      ← Archivos ignorados
│
└── 📄 Raíz del proyecto
    ├── index.html                      (Se copia a tauri-web/)
    └── [gestores HTML + assets]
```

---

## 🚀 Inicio Rápido

### Requisitos
- Node.js 18+
- Rust 1.70+ (para compilar Tauri)
- Windows 10+ / macOS 10.15+ / Linux

### Instalación local

```bash
# 1. Clonar repositorio
git clone https://github.com/tu-usuario/plantillaNotas.git
cd plantillaNotas

# 2. Instalar dependencias
npm install

# 3. Desarrollo (con hot reload)
npm run tauri:dev

# 4. Compilar para distribución
npm run tauri:build
```

### Distribución
El ejecutable se genera en:
```
src-tauri/target/release/bundle/nsis/
```

---

## 📖 Gestores Disponibles

### 1️⃣ **Gestor de Alumnos** (`gestor-alumnos.html`)
Administra la lista de alumnos del módulo.
- ✅ Cargar/importar alumnos
- ✅ Agregar nuevos
- ✅ Eliminar duplicados
- ✅ Exportar actualizado

### 2️⃣ **Gestor RRAA y Criterios** (`gestor-rraa-criterios.html`) ⭐
**Más importante**. Gestiona Resultados de Aprendizaje y Criterios.
- ✅ Edición de RRAA
- ✅ Edición de Criterios (código + texto)
- ✅ Ponderaciones por unidad
- ✅ Importación de datos
- ✅ Guardado automático

### 3️⃣ **Gestor de Unidades** (`gestor-unidades.html`)
Define las unidades didácticas del módulo.
- ✅ Crear/editar unidades
- ✅ Asignar horas
- ✅ Tipo de evaluación

### 4️⃣ **Gestor de Notas** (`gestor-notas.html`)
Registra calificaciones de actividades evaluables.
- ✅ Matriz alumno × criterio
- ✅ Edición inline
- ✅ Cálculo automático
- ✅ Validación 0-10

### 5️⃣ **Visor de Notas** (`visor-notas.html`)
Visualiza notas consolidadas por evaluación.
- ✅ Filtrado por evaluación
- ✅ Gráficos de distribución
- ✅ Exportación PDF

### 6️⃣ **Generador de Informes** (`informes.html`)
Crea reportes exportables.
- ✅ Informes por alumno
- ✅ Análisis por unidad
- ✅ Exportación Excel/PDF

---

## 🔄 Flujo de Trabajo Típico

```
1. Abre Plantilla Notas Local (EXE)
   ↓
2. Carga archivo Excel (drag & drop)
   ↓
3. Valida/importa estructura (RRAA, Criterios, Unidades)
   ↓
4. Edita en gestores:
   - Alumnos: lista de estudiantes
   - Criterios: definición y pesos
   - Unidades: calendario
   - Notas: calificaciones
   ↓
5. Visualiza en visor de notas
   ↓
6. Genera reportes para acta
   ↓
7. Descarga Excel actualizado
```

---

## 💾 Persistencia de Datos

### Almacenamiento local
- Archivo Excel seleccionado
- Última carpeta abierta
- Preferencias usuario

**Ubicación** (Windows):
```
%APPDATA%\Roaming\Plantilla Notas Local\
```

### Sincronización
- ⚠️ **No automática** (por diseño)
- Usuario controla cuándo descargar/subir a OneDrive
- Respeta versiones en caché local

---

## 🛠️ Arquitectura Técnica

### Stack actual
```
┌─────────────────────────────────┐
│   UI (HTML5 + CSS + JS)         │  ← Gestores web
├─────────────────────────────────┤
│   Tauri (app-bridge.js)         │  ← IPC, APIs nativas
├─────────────────────────────────┤
│   Rust (Tauri v2)               │  ← Ventana, eventos
├─────────────────────────────────┤
│   Node.js Backend (main.js)     │  ← Excel I/O
├─────────────────────────────────┤
│   XLSX + JSZip Libraries        │  ← Parseo Excel
└─────────────────────────────────┘
```

### IPC (Inter-Process Communication)
Tauri maneja:
- File dialogs (seleccionar Excel)
- File I/O (lectura/escritura)
- Notificaciones
- Sistema de ventanas

---

## 📊 Estructura Excel Esperada

### Hoja "DATOS"
```
Columnas:
- 0-2:   Nº Alumno, Nombre, Fecha Nac
- 5-6:   Nº RRAA, Descripción RRAA
- 8-11:  Código Unit, Nombre Unit, Eval, Horas
- 21-22: Código Criterio, Texto Criterio
```

### Hoja "PESOS"
```
Fila 3:  Códigos de criterios
Fila 4:  Reservada
Fila 5-20: Ponderaciones por unidad
Col X:   Ponderación Final
Col X+1: Ponderación Instituto
Col X+2: Ponderación Empresa
```

---

## 🔐 Seguridad

✅ **CSP deshabilitado en Tauri** (para compatibilidad Excel parsing)  
✅ **Archivos Excel locales** (no subidos)  
✅ **Sin conexión a internet** requerida  
✅ **Código open-source** (auditable)

---

## 📦 Dependencias Principales

```json
{
  "tauri": "2.11.0",           // Framework nativo
  "xlsx": "0.18.5",            // Lectura Excel
  "jszip": "3.10.1",           // Manipulación ZIP
  "electron": "30.5.1",        // Fallback (legacy)
  "electron-builder": "24.13"  // Empaquetado
}
```

**Rust** (`src-tauri/Cargo.toml`):
- `tauri` = 2
- `serde` / `serde_json` = serialización
- `rfd` = file dialogs nativos

---

## 🐛 Troubleshooting

### "No se abre el archivo Excel"
→ Verifica que esté en formato `.xlsx` (no `.xls`)  
→ Comprueba que no esté abierto en otra aplicación

### "Los cambios no se guardan"
→ Usa botón "💾 Guardar Cambios Ahora"  
→ Verifica que tengas permisos de escritura

### "Errores al compilar (Rust)"
→ Actualiza Rust: `rustup update`  
→ Limpia caché: `cargo clean`

### "Tauri no inicia"
→ Reinstala deps: `npm install`  
→ Reconstruye: `npm run tauri:build`

---

## 🚀 Próximas Mejoras (v2.1+)

- [ ] Sincronización automática con OneDrive
- [ ] Historial de cambios con undo/redo
- [ ] Validaciones complejas (ponderaciones suma 100%)
- [ ] Tema oscuro
- [ ] Soporte multiidioma
- [ ] API REST (para integración externa)

---

## 📞 Contacto y Contribución

**Autor**: Sebastián (@sebantonio)  
**Email**: sebantonio@gmail.com  
**GitHub**: https://github.com/sebantonio/plantillaNotas  
**Licencia**: MIT

### Contribuir
1. Fork del repositorio
2. Crea rama: `git checkout -b feature/mi-mejora`
3. Commit: `git commit -m "feat: Descripción"`
4. Push: `git push origin feature/mi-mejora`
5. Pull Request en GitHub

---

## 📚 Documentación Relacionada

- [TAURI_MIGRACION.md](TAURI_MIGRACION.md) — Detalles de la migración a Tauri
- [MAPA_PROYECTO.md](MAPA_PROYECTO.md) — Guía visual del proyecto
- [CHECKLIST_SETUP.md](CHECKLIST_SETUP.md) — Setup local paso a paso
- [INSTRUCCIONES_GITHUB.md](INSTRUCCIONES_GITHUB.md) — Git/GitHub/Claude Code

---

**¡Gracias por usar Plantilla Notas Local!** 🎓
