#!/bin/bash

# Script para empaquetar aplicacion Electron de Plantilla Notas Local
# Uso en Linux/Mac: chmod +x empaquetar.sh && ./empaquetar.sh

set -e

# Colores
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo ""
echo "===================================="
echo "  Empaquetador: Plantilla Notas"
echo "===================================="
echo ""

# Verificar si npm esta instalado
if ! command -v npm &> /dev/null; then
    echo -e "${RED}[ERROR]${NC} NPM no esta instalado"
    echo "Descargalo desde: https://nodejs.org/"
    exit 1
fi

echo -e "${GREEN}[OK]${NC} NPM detectado"

# Verificar si estamos en el directorio correcto
if [ ! -f package.json ]; then
    echo -e "${RED}[ERROR]${NC} No se encontro package.json"
    echo "Este script debe ejecutarse en la raiz del proyecto"
    exit 1
fi

echo -e "${GREEN}[OK]${NC} Proyecto detectado"

# Menu de opciones
echo ""
echo "Elige una opcion:"
echo ""
echo "1) Instalar dependencias (npm install)"
echo "2) Crear portable para Windows"
echo "3) Crear instalador NSIS (Windows)"
echo "4) Crear ambos formatos"
echo "5) Compilar para Linux"
echo "6) Compilar para macOS"
echo "7) Compilar todo (Windows + Linux + Mac)"
echo "8) Limpiar compilaciones previas"
echo "9) Salir"
echo ""

read -p "Opcion [1-9]: " choice

case $choice in
  1)
    install_deps
    ;;
  2)
    install_deps
    build_portable
    ;;
  3)
    install_deps
    build_nsis
    ;;
  4)
    install_deps
    build_windows
    ;;
  5)
    install_deps
    build_linux
    ;;
  6)
    install_deps
    build_macos
    ;;
  7)
    install_deps
    build_all
    ;;
  8)
    clean_build
    ;;
  9)
    echo "Hasta luego!"
    exit 0
    ;;
  *)
    echo -e "${RED}[ERROR]${NC} Opcion no valida"
    exit 1
    ;;
esac

# ================== FUNCIONES ==================

function install_deps() {
    echo ""
    echo -e "${YELLOW}[*]${NC} Instalando dependencias..."
    npm install
    if [ $? -eq 0 ]; then
        echo -e "${GREEN}[OK]${NC} Dependencias instaladas"
    else
        echo -e "${RED}[ERROR]${NC} Fallo la instalacion"
        exit 1
    fi
}

function build_portable() {
    echo ""
    echo -e "${YELLOW}[*]${NC} Creando portable para Windows..."
    echo "Esto puede tomar 2-3 minutos..."
    npm run dist -- -c.win.target=portable
    show_result
}

function build_nsis() {
    echo ""
    echo -e "${YELLOW}[*]${NC} Creando instalador NSIS (Windows)..."
    echo "Esto puede tomar 2-3 minutos..."
    npm run dist -- -c.win.target=nsis
    show_result
}

function build_windows() {
    echo ""
    echo -e "${YELLOW}[*]${NC} Creando formatos Windows (portable + NSIS)..."
    echo "Esto puede tomar 3-5 minutos..."
    npm run dist -- --win
    show_result
}

function build_linux() {
    echo ""
    echo -e "${YELLOW}[*]${NC} Creando paquete para Linux..."
    echo "Esto puede tomar 2-3 minutos..."
    npm run dist -- --linux
    show_result
}

function build_macos() {
    echo ""
    echo -e "${YELLOW}[*]${NC} Creando paquete para macOS..."
    echo "Esto puede tomar 2-3 minutos..."
    npm run dist -- --mac
    show_result
}

function build_all() {
    echo ""
    echo -e "${YELLOW}[*]${NC} Compilando para todas las plataformas..."
    echo "Esto puede tomar 10+ minutos..."
    npm run dist
    show_result
}

function clean_build() {
    echo ""
    echo -e "${YELLOW}[*]${NC} Limpiando compilaciones previas..."
    
    if [ -d dist ]; then
        rm -rf dist
        echo -e "${GREEN}[OK]${NC} Directorio ./dist eliminado"
    fi
    
    if [ -d out ]; then
        rm -rf out
        echo -e "${GREEN}[OK]${NC} Directorio ./out eliminado"
    fi
    
    echo -e "${GREEN}[OK]${NC} Limpieza completada"
}

function show_result() {
    echo ""
    echo "===================================="
    echo -e "${GREEN}   COMPILACION EXITOSA${NC}"
    echo "===================================="
    echo ""
    echo "Los archivos estan en: ./dist/"
    echo ""
    echo "Archivos generados:"
    ls -lh dist/ 2>/dev/null || echo "(No se encontraron archivos)"
    echo ""
}

# Si llegamos aca, mostrar resultado
show_result
