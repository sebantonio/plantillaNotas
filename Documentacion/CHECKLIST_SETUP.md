# ✅ CHECKLIST: GitHub + Claude Code Setup

## 📋 FASE 1: PREPARACIÓN (15 minutos)

### 1.1 Crear repositorio en GitHub
- [ ] Ir a https://github.com/new
- [ ] **Name**: `gestor-notas-fp`
- [ ] **Visibility**: Public
- [ ] Click **Create repository**
- [ ] Copiar URL (formato: https://github.com/tu-usuario/gestor-notas-fp.git)

### 1.2 Instalar herramientas
- [ ] Descargar e instalar Git: https://git-scm.com/downloads
- [ ] Descargar e instalar VS Code: https://code.visualstudio.com/
- [ ] Descargar e instalar Claude Code (en VS Code)

### 1.3 Configurar Git (solo primera vez)
```bash
git config --global user.name "Tu Nombre"
git config --global user.email "tu@email.com"
git config --global user.name          # Verificar
```
- [ ] Nombre configurado ✅
- [ ] Email configurado ✅

---

## 📁 FASE 2: CLONAR Y COPIAR (10 minutos)

### 2.1 Clonar repositorio
```bash
# En tu terminal/CMD
cd Documentos
git clone https://github.com/tu-usuario/gestor-notas-fp.git
cd gestor-notas-fp
```
- [ ] Carpeta clonada ✅

### 2.2 Verificar archivos
```bash
ls                    # Mac/Linux
dir                   # Windows
```
Deberías ver:
- [ ] `public/` (carpeta con HTML)
- [ ] `src/` (carpeta)
- [ ] `docs/` (carpeta)
- [ ] `README.md`
- [ ] `package.json`
- [ ] `.gitignore`

---

## 🚀 FASE 3: PRIMER PUSH (5 minutos)

### 3.1 Agregar cambios
```bash
git status              # Ver cambios
git add .              # Agregar todos
```
- [ ] Cambios agregados ✅

### 3.2 Crear commit
```bash
git commit -m "Initial commit: Gestores de alumnos y RRAA"
```
- [ ] Commit creado ✅

### 3.3 Enviar a GitHub
```bash
git push -u origin main
```
- [ ] Archivos en GitHub ✅

### 3.4 Verificar en GitHub
- [ ] Ve a https://github.com/tu-usuario/gestor-notas-fp
- [ ] Deberías ver todos los archivos
- [ ] README visible
- [ ] Carpeta `public/` con los HTMLs

---

## 💻 FASE 4: VS CODE (10 minutos)

### 4.1 Abrir VS Code
```bash
code .
```
- [ ] VS Code abierto ✅

### 4.2 Instalar extensiones
Click en el icono **Extensiones** (lado izquierdo)

Buscar e instalar:
- [ ] **Live Server** by Ritwick Dey
- [ ] **Prettier** by esbenp  
- [ ] **GitHub Copilot** by GitHub (opcional)
- [ ] **Auto Rename Tag** by formulahendry

**Total: 4+ extensiones instaladas**

### 4.3 Verificar configuración
- [ ] `.vscode/settings.json` visible en VS Code
- [ ] `.vscode/extensions.json` visible en VS Code
- [ ] Live Server en barra inferior (o en menú)

---

## 🤖 FASE 5: CLAUDE CODE (5 minutos)

### 5.1 Abrir Claude Code
```bash
claude-code open .
```
- [ ] Claude Code abierto ✅

### 5.2 Verificar archivos
- [ ] Ver `public/gestor-alumnos.html`
- [ ] Ver `public/gestor-rraa-criterios.html`
- [ ] Ver `README.md`

### 5.3 Test simple
```
Mensaje a Claude: "¿Qué archivos HTML hay en este proyecto?"
Respuesta esperada: Lista los 2 gestores
```
- [ ] Claude Code responde correctamente ✅

---

## 🧪 FASE 6: PRUEBAS (10 minutos)

### 6.1 Probar Gestor de Alumnos

En VS Code:
1. Click derecho en `public/gestor-alumnos.html`
2. Selecciona **Open with Live Server**
3. Abre una ventana de navegador
- [ ] Página carga correctamente
- [ ] Se ve la interfaz colorida
- [ ] Puedes cargar un Excel
- [ ] Live Server funciona

### 6.2 Probar Gestor RRAA
1. Click derecho en `public/gestor-rraa-criterios.html`
2. Selecciona **Open with Live Server**
- [ ] Página carga correctamente
- [ ] Se ven las pestañas (RRAA | Criterios)
- [ ] Interfaz funciona

### 6.3 Probar Excel real
- [ ] Carga tu `plantilla313_dual.xlsx`
- [ ] Muestra alumnos existentes
- [ ] Puedes agregar un alumno
- [ ] Puedes descargar actualizado

---

## 📝 FASE 7: PRIMER COMMIT (5 minutos)

### 7.1 Hacer un cambio pequeño
En VS Code:
1. Abre `public/gestor-alumnos.html`
2. Busca `<title>` (línea ~3)
3. Cambia algo pequeño (comentario, etc.)

### 7.2 Guardar y commitear
```bash
git status                  # Ver cambio
git add public/gestor-alumnos.html
git commit -m "fix: Mejorar descripción del título"
git push origin main
```
- [ ] Cambio enviado a GitHub ✅

### 7.3 Verificar en GitHub
- [ ] Vuelve a https://github.com/tu-usuario/gestor-notas-fp
- [ ] Verifica que el cambio esté ahí

---

## 🌐 FASE 8: GITHUB PAGES (OPCIONAL, 5 minutos)

### 8.1 Habilitar GitHub Pages
1. Ve a Settings del repositorio
2. Busca **Pages** en la barra lateral
3. Source: **main** → Save

### 8.2 Acceder a tu sitio
```
https://tu-usuario.github.io/gestor-notas-fp/
```

Para abrir las aplicaciones:
```
https://tu-usuario.github.io/gestor-notas-fp/public/gestor-alumnos.html
https://tu-usuario.github.io/gestor-notas-fp/public/gestor-rraa-criterios.html
```

- [ ] GitHub Pages funciona (espera 1-5 minutos)
- [ ] Puedes acceder desde cualquier lugar

---

## 📚 FASE 9: DOCUMENTACIÓN (Lectura)

- [ ] Leer `README.md` (2 min)
- [ ] Leer `docs/GUIA.md` (5 min)
- [ ] Guardar `INSTRUCCIONES_GITHUB.md` (referencia)
- [ ] Guardar `PASOS_RAPIDOS.md` (referencia)

---

## 🎯 LISTO PARA TRABAJAR

### Checklist final
- [ ] Git instalado y configurado
- [ ] Repositorio GitHub creado
- [ ] Archivos clonados localmente
- [ ] VS Code con extensiones
- [ ] Claude Code funciona
- [ ] Primer commit hecho
- [ ] Aplicaciones probadas
- [ ] GitHub Pages (opcional)
- [ ] Documentación leída

### Siguiente sesión será:
```bash
cd ~/Documentos/gestor-notas-fp
git pull origin main            # Descargar cambios
code .                          # Abrir VS Code
claude-code open .              # O Claude Code
# Editar y mejorar...
git add .
git commit -m "feat: Nueva funcionalidad"
git push origin main
```

---

## 🎁 BONUS: Atajos útiles

### Windows/Mac
- **Abrir terminal en VS Code**: Ctrl + ` (backtick)
- **Formatear código**: Alt + Shift + F
- **Crear archivo nuevo**: Ctrl + N
- **Guardar**: Ctrl + S
- **Buscar**: Ctrl + F

### Git rápido
```bash
# Ver estado actual
git status

# Ver últimos commits
git log --oneline -5

# Deshacer último commit (mantener cambios)
git reset --soft HEAD~1

# Cambiar rama
git checkout -b feature/nombre

# Volver a main
git checkout main
```

---

## ❓ SI ALGO NO FUNCIONA

1. **Error de Git**: Recarga y vuelve a intentar
2. **Live Server no aparece**: Recarga VS Code
3. **No puedo hacer push**: Verifica que Git esté configurado
4. **GitHub Pages no funciona**: Espera 5 minutos y recarga
5. **Claude Code no abre**: Instala extensión primero en VS Code

---

## 📊 ESTADO ACTUAL

```
✅ Proyecto: LISTO
✅ 2 Gestores: FUNCIONALES
✅ Documentación: COMPLETA
✅ GitHub: CONFIGURADO
✅ VS Code: LISTO
✅ Claude Code: INTEGRADO

📊 Progreso: ████████░░ 80%
```

---

## 🚀 ¡ENHORABUENA!

Ahora tienes:
- Un proyecto profesional con versionado Git
- Dos aplicaciones web funcionales
- Documentación clara
- Infraestructura para crecer
- Setup listo para Claude Code

**¡A programar se ha dicho!** 💪

---

**Tiempo total**: ~60 minutos  
**Versión**: 2.0.0  
**Fecha**: 2026-05-07  
**Status**: ✅ LISTO (Tauri v2)
