# ⚡ PASOS RÁPIDOS: GitHub + Claude Code

## 🔥 RESUMEN EJECUTIVO (5 MINUTOS)

### Paso 1: Crear repositorio en GitHub
```
1. Ve a https://github.com/new
2. Repository name: gestor-notas-fp
3. Visibility: Public
4. Create repository
```

### Paso 2: Clonar y configurar

```bash
# En tu terminal
cd Documentos

# Clonar el repo (cambiar URL)
git clone https://github.com/tu-usuario/gestor-notas-fp.git
cd gestor-notas-fp

# Configurar git (primera vez)
git config --global user.name "Tu Nombre"
git config --global user.email "tu@email.com"
```

### Paso 3: Enviar archivos a GitHub

```bash
# Los archivos ya están en el directorio
# Solo hacer:

git add .
git commit -m "Initial commit: Gestores de alumnos y RRAA"
git push -u origin main
```

### Paso 4: Abrir en VS Code

```bash
code .
```

Instala extensiones:
- Live Server
- Prettier
- GitHub Copilot

### Paso 5: Usar con Claude Code

```bash
claude-code open .
```

---

## 🎯 FLUJO DIARIO DE TRABAJO

### Mañana cuando empieces:
```bash
cd ~/ruta/a/gestor-notas-fp
git pull origin main
```

### Durante el día:
```bash
# Editar archivos en VS Code / Claude Code
# Probar con Live Server

# Cuando termines una funcionalidad:
git add .
git commit -m "feat: descripción breve"
git push origin main
```

### Si trabajas en múltiples máquinas:
```bash
git pull origin main  # Descargar cambios
# Editar
git push origin main  # Enviar
```

---

## 📁 ARCHIVO ESTRUCTURA ACTUAL

```
gestor-notas-fp/
├── public/
│   ├── gestor-alumnos.html              ✅ Funciona
│   └── gestor-rraa-criterios.html       ✅ Funciona
├── src/                                  (Para expandir)
├── docs/
│   ├── GUIA.md
│   └── CAMBIOS.md
├── .vscode/
├── README.md
├── package.json
├── LICENSE
└── .gitignore
```

---

## 🚀 PRÓXIMAS TAREAS

1. **Crear gestor de notas** → `public/gestor-notas.html`
2. **Backend con Node.js** → `api/server.js`
3. **Sincronización OneDrive** → `src/utils/onedrive.js`
4. **API REST** → `api/routes/*.js`
5. **Base de datos** → MongoDB o Firebase

---

## ⚠️ PROBLEMAS COMUNES

### Regla importante al modificar el Excel desde Electron

No reescribir nunca el libro completo con `XLSX.writeFile()` o ExcelJS cuando se trabaje con la plantilla real.

Procedimiento correcto:

1. Leer con `xlsx` solo para localizar filas, columnas y validar estructura.
2. Abrir el `.xlsx` como ZIP con `JSZip`.
3. Modificar solo el XML de la hoja necesaria, por ejemplo:
   - `DATOS` para alumnos, RRAA y unidades.
   - `PESOS` para criterios de evaluación.
4. Reempaquetar el ZIP dejando intactos los XML de las demás hojas.

Motivo: la plantilla tiene muchas fórmulas y hojas grandes. Si una librería reescribe todo el libro, puede inflar el archivo hasta decenas de MB y corromper fórmulas en las hojas de evaluación.

### "fatal: repository not found"
```
❌ Verificar que la URL sea correcta
✅ git clone https://github.com/TU-USUARIO/gestor-notas-fp.git
```

### "Permission denied (publickey)"
```
✅ Configurar SSH key:
https://docs.github.com/en/authentication/connecting-to-github-with-ssh
```

### No puedo hacer push
```
✅ Primero hacer pull:
git pull origin main
git push origin main
```

---

## 📚 REFERENCIAS RÁPIDAS

**Git Docs**: https://git-scm.com/doc
**GitHub Docs**: https://docs.github.com
**VS Code**: https://code.visualstudio.com/docs
**Claude Code**: Dentro de tu IDE

---

## 💡 TIPS

✅ Commit frecuente (cada funcionalidad pequeña)
✅ Mensajes descriptivos
✅ Un cambio por commit
✅ Push al final del día
✅ README actualizado
✅ .gitignore configurado

---

## 🎓 EJEMPLO REAL

```bash
# Empezar a trabajar
cd ~/Documentos/gestor-notas-fp
git pull origin main

# Editar archivo en VS Code
# Agregar validación a gestor-alumnos.html

# Ver cambios
git status

# Hacer commit
git add public/gestor-alumnos.html
git commit -m "feat: Agregar validación de email en alumnos"

# Enviar
git push origin main

# Ver en GitHub
# https://github.com/tu-usuario/gestor-notas-fp
```

---

Contacto: Si tienes dudas sobre Git → 
https://git-scm.com/book/es/v2 (muy completo)

¡Éxito con tu proyecto! 🚀
