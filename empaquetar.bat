@echo off
REM Script para empaquetar aplicacion Electron de Plantilla Notas Local
REM Uso: Simplemente ejecutar este archivo o ejecutar como administrador

setlocal enabledelayedexpansion

echo.
echo ====================================
echo   Empaquetador: Plantilla Notas Local
echo ====================================
echo.

REM Verificar si npm esta instalado
where npm >nul 2>nul
if !errorlevel! neq 0 (
  echo [ERROR] NPM no esta instalado.
  echo Descargalo desde: https://nodejs.org/
  pause
  exit /b 1
)

echo [OK] NPM detectado

REM Verificar si estamos en el directorio correcto
if not exist package.json (
  echo [ERROR] No se encontro package.json
  echo Este script debe ejecutarse en la raiz del proyecto
  pause
  exit /b 1
)

echo [OK] Proyecto detectado

REM Menu de opciones
echo.
echo Elige una opcion:
echo.
echo 1) Instalar dependencias (npm install)
echo 2) Crear portable.exe (archivo unico)
echo 3) Crear instalador NSIS (programa)
echo 4) Crear ambos (portable + instalador)
echo 5) Limpiar compilaciones previas
echo 6) Salir
echo.

set /p choice="Opcion [1-6]: "

if "%choice%"=="1" (
  call :install_deps
  goto :end
)

if "%choice%"=="2" (
  call :build_portable
  goto :end
)

if "%choice%"=="3" (
  call :build_nsis
  goto :end
)

if "%choice%"=="4" (
  call :install_deps
  call :build_all
  goto :end
)

if "%choice%"=="5" (
  call :clean_build
  goto :end
)

if "%choice%"=="6" (
  goto :end
)

echo [ERROR] Opcion no valida
goto :end

REM ================== FUNCIONES ==================

:install_deps
  echo.
  echo [*] Instalando dependencias...
  call npm install
  if !errorlevel! equ 0 (
    echo [OK] Dependencias instaladas
  ) else (
    echo [ERROR] Fallo la instalacion
  )
  exit /b !errorlevel!

:build_portable
  echo.
  echo [*] Creando portable.exe...
  echo Esto puede tomar 2-3 minutos...
  call npm run dist -- -c.win.target=portable
  if !errorlevel! equ 0 (
    echo [OK] Portable creado en ./dist/
    call :show_result
  ) else (
    echo [ERROR] Fallo la compilacion
  )
  exit /b !errorlevel!

:build_nsis
  echo.
  echo [*] Creando instalador NSIS...
  echo Esto puede tomar 2-3 minutos...
  call npm run dist -- -c.win.target=nsis
  if !errorlevel! equ 0 (
    echo [OK] Instalador creado en ./dist/
    call :show_result
  ) else (
    echo [ERROR] Fallo la compilacion
  )
  exit /b !errorlevel!

:build_all
  echo.
  echo [*] Creando ambos formatos (portable + NSIS)...
  echo Esto puede tomar 3-5 minutos...
  call npm run dist
  if !errorlevel! equ 0 (
    echo [OK] Archivos creados en ./dist/
    call :show_result
  ) else (
    echo [ERROR] Fallo la compilacion
  )
  exit /b !errorlevel!

:clean_build
  echo.
  echo [*] Limpiando compilaciones previas...
  if exist dist (
    rmdir /s /q dist
    echo [OK] Directorio ./dist eliminado
  )
  if exist out (
    rmdir /s /q out
    echo [OK] Directorio ./out eliminado
  )
  echo [OK] Limpieza completada
  exit /b 0

:show_result
  echo.
  echo ====================================
  echo   COMPILACION EXITOSA
  echo ====================================
  echo.
  echo Los archivos estan en: ./dist/
  echo.
  echo Archivos generados:
  dir /b dist\*.exe
  echo.
  pause
  exit /b 0

:end
echo.
echo Programa finalizado.
pause
