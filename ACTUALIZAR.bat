@echo off
title BRIAV32 - Actualizador de Dashboard
color 1F
cd /d "%~dp0"

echo.
echo  ================================================================
echo   BRIAV32 - Actualizador de Dashboard
echo  ================================================================
echo.

:: Verificar que Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python no está instalado.
    echo  Descárgalo en: python.org
    echo.
    pause
    exit /b
)

:: Buscar el archivo Excel en la carpeta
set EXCEL=
for %%f in (*.xlsx) do (
    if not defined EXCEL set EXCEL=%%f
)

if not defined EXCEL (
    echo  ERROR: No se encontró ningún archivo Excel en esta carpeta.
    echo  Asegúrate de que el archivo .xlsx esté aquí.
    echo.
    pause
    exit /b
)

echo  Archivo encontrado: %EXCEL%
echo.
echo  Ejecutando actualizador...
echo.

python actualizar.py "%EXCEL%"

echo.
if errorlevel 1 (
    echo  Algo salió mal. Revisa los mensajes de arriba.
) else (
    echo  ================================================================
    echo   index.html actualizado correctamente.
    echo   Sube el index.html a GitHub para ver los cambios en línea.
    echo  ================================================================
)

echo.
pause
