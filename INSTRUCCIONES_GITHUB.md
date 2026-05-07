# 🚀 Guía Completa: GitHub + Claude Code + Visual Studio

## Paso 1: Crear Repositorio en GitHub

### 1.1 Crear el repositorio en GitHub.com

1. **Ir a GitHub**: https://github.com/new
2. **Rellenar formulario**:
   - **Repository name**: `gestor-notas-fp`
   - **Description**: Sistema web para gestionar notas en FP
   - **Visibility**: Public (o Private si prefieres)
   - **Initialize with README**: NO (ya tenemos uno)
   - **Add .gitignore**: NO (ya tenemos uno)
   - **Choose a license**: MIT (ya lo tenemos)

3. **Click en "Create repository"**

### 1.2 URL del repositorio
Después de crear, tendrás una URL como:
```
https://github.com/tu-usuario/gestor-notas-fp.git
```

---

## Paso 2: Clonar Repositorio Localmente

### En Terminal / CMD:

```bash
# Ir al directorio donde quieras el proyecto
cd Documentos
# O donde prefieras

# Clonar el repositorio vacío que acabas de crear
git clone https://github.com/tu-usuario/gestor-notas-fp.git
cd gestor-notas-fp
```

---

## Paso 3: Copiar Archivos Locales

El repositorio que creé está en `/home/claude/gestor-notas-fp`

### Opción A: Copiar archivos manualmente
```bash
# Copiar todo el contenido del repositorio local al nuevo
# Reemplaza SOURCE con la ruta real

# En Windows PowerShell:
Copy-Item -Path "C:\Users\tu-usuario\Downloads\gestor-notas-fp\*" `
          -Destination "C:\ruta\a\tu\repo\gestor-notas-fp" -Recurse -Force

# En Mac/Linux:
cp -r ~/descargas/gestor-notas-fp/* ~/Documentos/gestor-notas-fp/
```

### Opción B: Usar Git directamente (Recomendado)

```bash
# Dentro de tu repositorio clonado:
cd gestor-notas-fp

# Agregar todos los archivos
git add .

# Crear primer commit
git commit -m "Initial commit: Gestores de alumnos y RRAA"

# Enviar a GitHub
git push -u origin main
```

---

## Paso 4: Configurar VS Code / Claude Code

### 4.1 Abrir en Visual Studio Code

```bash
# En la carpeta del repositorio
code .
```

**Extensiones recomendadas**:
1. **Live Server** (Ritwick Dey)
   - Para servir los archivos HTML localmente
   - Click derecho en `public/gestor-alumnos.html` → "Open with Live Server"

2. **Prettier** (esbenp)
   - Formateo automático de código

3. **GitHub Copilot** (Microsoft)
   - Autocompletado inteligente

4. **REST Client** (Huachao Mao)
   - Para probar APIs (cuando las agregues)

### 4.2 Abrir en Claude Code

```bash
# En la carpeta del repositorio
claude-code open .
```

---

## Paso 5: Configurar Git Localmente

### Primera vez:

```bash
git config --global user.name "Tu Nombre"
git config --global user.email "tu-email@example.com"
```

### Verificar:

```bash
git config --global user.name
git config --global user.email
```

---

## Paso 6: Flujo de Trabajo Diario

### 6.1 Antes de empezar a trabajar:

```bash
# Descargar cambios de GitHub (si trabajas en múltiples máquinas)
git pull origin main
```

### 6.2 Mientras trabajas:

```bash
# Ver cambios
git status

# Agregar cambios específicos
git add public/gestor-alumnos.html

# O agregar todo
git add .

# Commit con mensaje descriptivo
git commit -m "feat: Agregar validación de emails en alumnos"

# Enviar a GitHub
git push origin main
```

### 6.3 Convenciones de commit:

```
feat:     Nueva funcionalidad
fix:      Corrección de bug
docs:     Cambios en documentación
style:    Formateo, sin cambios de código
refactor: Refactorización de código
test:     Agregar o actualizar tests
chore:    Cambios de dependencias o configuración
```

**Ejemplos**:
```bash
git commit -m "feat: Agregar gestor de notas"
git commit -m "fix: Corregir error al cargar Excel"
git commit -m "docs: Actualizar guía de uso"
git commit -m "refactor: Mejorar función de validación"
```

---

## Paso 7: Estructura Final del Proyecto

```
gestor-notas-fp/
├── public/
│   ├── gestor-alumnos.html           ✅ Listo
│   ├── gestor-rraa-criterios.html    ✅ Listo
│   └── gestor-notas.html             ⏳ Próximo
├── src/
│   ├── components/
│   ├── pages/
│   ├── utils/
│   └── styles/
├── docs/
│   ├── GUIA.md                       ✅ Completa
│   ├── CAMBIOS.md                    ✅ Actualizado
│   └── API.md                        ⏳ Para API
├── tests/
├── .vscode/
│   ├── settings.json
│   └── extensions.json
├── README.md                         ✅ Profesional
├── LICENSE                           ✅ MIT
├── package.json                      ✅ Configurado
└── .gitignore                        ✅ Completo
```

---

## Paso 8: Próximos Pasos para Expandir

### 8.1 Agregar página de notas

Crear `public/gestor-notas.html` con:
- Tabla de notas editable
- Cálculo automático de promedios
- Validación de rangos (0-10)
- Exportación de resultados

### 8.2 Crear API con Node.js

```bash
npm install express cors xlsx
```

Estructura:
```
api/
├── server.js
├── routes/
│   ├── alumnos.js
│   ├── criterios.js
│   └── notas.js
└── middleware/
```

### 8.3 Conectar con OneDrive

```javascript
// Usar Azure SDK
npm install @azure/identity @microsoft/microsoft-graph-client
```

---

## Paso 9: Usa Claude Code para Desarrollo

### 9.1 Abrir archivos en Claude Code

```bash
claude-code open public/gestor-alumnos.html
```

### 9.2 Peticiones útiles a Claude:

```
"Mejora la función cargarAlumnos para manejar errores de Excel"

"Refactoriza el código de validación en una función separada"

"Agrega comentarios JSDoc a todas las funciones principales"

"Crea una función para calcular el promedio ponderado de notas"

"Integra autenticación con GitHub usando OAuth"
```

---

## Paso 10: Tips y Trucos

### 10.1 Ver el historial de cambios
```bash
git log --oneline
git log --graph --all --decorate
```

### 10.2 Deshacer cambios
```bash
# Deshacer cambios locales (antes de commit)
git restore archivo.html

# Deshacer el último commit (mantener cambios)
git reset --soft HEAD~1

# Deshacer el último commit (perder cambios)
git reset --hard HEAD~1
```

### 10.3 Crear ramas de desarrollo
```bash
# Crear rama nueva
git checkout -b feature/gestor-notas

# Trabajar en la rama
git add .
git commit -m "feat: Agregar gestor de notas"

# Volver a main
git checkout main

# Fusionar rama
git merge feature/gestor-notas

# Eliminar rama
git branch -d feature/gestor-notas
```

### 10.4 Publicar en GitHub Pages (Gratis)

```bash
# En GitHub.com:
# 1. Ve a Settings del repositorio
# 2. En GitHub Pages, selecciona "main branch"
# 3. Tu sitio estará en: https://tu-usuario.github.io/gestor-notas-fp/
# 4. Accede a: https://tu-usuario.github.io/gestor-notas-fp/public/gestor-alumnos.html
```

---

## Paso 11: Checklist Final

- [ ] Repositorio creado en GitHub
- [ ] Archivos enviados (git push)
- [ ] VS Code configurado
- [ ] Claude Code funcionando
- [ ] Live Server probado
- [ ] README visible en GitHub
- [ ] GitHub Pages habilitado (opcional)
- [ ] Extensiones instaladas
- [ ] Primera contribución hecha

---

## 📞 Comandos Git Más Usados

```bash
# Clonar repositorio
git clone https://github.com/usuario/repo.git

# Ver estado
git status

# Agregar cambios
git add .
git add archivo.html
git add "*.js"

# Commit
git commit -m "mensaje"

# Enviar
git push origin main

# Descargar
git pull origin main

# Ver historial
git log

# Ver diferencias
git diff

# Crear rama
git checkout -b nombre-rama

# Cambiar rama
git checkout main

# Listar ramas
git branch

# Eliminar rama
git branch -d nombre-rama

# Fusionar rama
git merge nombre-rama
```

---

## 🎯 Objetivo Final

Con esta configuración tienes:
- ✅ Versionado de código con Git/GitHub
- ✅ Colaboración fácil
- ✅ Historial de cambios
- ✅ Desarrollo con Claude Code
- ✅ Editor moderno (VS Code)
- ✅ Sitio web público (GitHub Pages)
- ✅ Documentación profesional
- ✅ Estructura escalable para agregar backend

¡Ahora puedes trabajar profesionalmente en tu proyecto FP! 🚀
