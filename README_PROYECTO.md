# 📊 Gestor de Notas FP

Sistema web para gestionar notas, alumnos, criterios y resultados de aprendizaje en módulos de formación profesional.

## ✨ Características

- 👥 **Gestión de Alumnos**: Carga, agregar y eliminar alumnos
- 🎯 **RRAA y Criterios**: Gestionar Resultados de Aprendizaje y criterios de evaluación
- 📝 **Notas**: Sistema de calificación (en desarrollo)
- 📥 **Importar/Exportar**: Excel integrado con SheetJS
- 🚀 **Sin servidor**: Funciona 100% en el navegador
- 🎨 **Interfaz moderna**: Responsive y accesible

## 🚀 Comenzar

### Requisitos
- Navegador moderno (Chrome, Firefox, Edge, Safari)
- No necesita instalación

### Uso Rápido

1. **Descargar el proyecto**
   ```bash
   git clone https://github.com/tu-usuario/gestor-notas-fp.git
   cd gestor-notas-fp
   ```

2. **Abrir en navegador**
   - Gestor de Alumnos: Abre `public/gestor-alumnos.html`
   - Gestor de RRAA: Abre `public/gestor-rraa-criterios.html`

3. **Cargar tu Excel** y comenzar a gestionar

## 📁 Estructura del Proyecto

```
gestor-notas-fp/
├── public/
│   ├── gestor-alumnos.html
│   ├── gestor-rraa-criterios.html
│   └── gestor-notas.html (próximamente)
├── src/
│   ├── components/
│   ├── utils/
│   ├── styles/
│   └── pages/
├── docs/
│   ├── GUIA.md
│   ├── API.md
│   └── CAMBIOS.md
├── tests/
├── .vscode/
│   └── settings.json
├── .gitignore
├── README.md
├── package.json
└── LICENSE
```

## 🛠️ Desarrollo

### Con Claude Code (Recomendado)
```bash
claude-code open .
```

### Con Visual Studio Code
```bash
code .
```

### Configuración recomendada
- Extensión: **Live Server** para pruebas locales
- Extensión: **Prettier** para formateo
- Node.js v16+ (opcional, para herramientas)

## 📝 Funcionalidades por Módulo

### Gestor de Alumnos (`gestor-alumnos.html`)
- Cargar Excel desde OneDrive o local
- Ver lista de alumnos existentes
- Agregar nuevos alumnos
- Eliminar alumnos
- Descargar Excel actualizado

### Gestor RRAA y Criterios (`gestor-rraa-criterios.html`)
- Gestionar Resultados de Aprendizaje
- Gestionar Criterios de Evaluación
- Editar ponderaciones
- Descargar Excel actualizado

### Gestor de Notas (En desarrollo)
- Interfaz para introducir calificaciones
- Cálculo automático de promedios
- Validación de datos
- Exportar informes

## 🔄 Flujo de Trabajo

1. **Cargar** archivo Excel desde tu computadora
2. **Gestionar** alumnos, RRAA y criterios
3. **Descargar** el archivo actualizado
4. Importar en Moodle o tu sistema de gestión

## 🤝 Contribuir

Las contribuciones son bienvenidas. Por favor:

1. Fork el proyecto
2. Crea una rama (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto está bajo la licencia MIT - ver el archivo [LICENSE](LICENSE) para detalles.

## 👨‍💼 Autor

**Seba** - Profesor FP, Especialista en Electrónica y Desarrollo
- Email: contacto@tumail.com
- GitHub: [@tu-usuario](https://github.com/tu-usuario)

## 📞 Soporte

- 📧 Issues: [GitHub Issues](https://github.com/tu-usuario/gestor-notas-fp/issues)
- 💬 Discussions: [GitHub Discussions](https://github.com/tu-usuario/gestor-notas-fp/discussions)

## 🗺️ Hoja de Ruta

- [x] Gestor de Alumnos
- [x] Gestor RRAA y Criterios
- [ ] Gestor de Notas
- [ ] Sincronización con OneDrive
- [ ] API REST (Node.js)
- [ ] Base de datos (MongoDB)
- [ ] Autenticación de usuarios
- [ ] Integración con Moodle

---
**Versión**: 0.1.0 | **Última actualización**: $(date +%Y-%m-%d)
