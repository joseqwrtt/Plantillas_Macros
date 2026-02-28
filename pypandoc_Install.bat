@echo off
chcp 65001 >nul
title Instalación manual de Pandoc portable
echo ======================================================
echo 🧩 Instalación manual de Pandoc portable
echo ======================================================
echo.

:: Paso 1: Instalar pypandoc
echo Ejecutando: pip install pypandoc
pip install pypandoc

:: Paso 2: Descargar Pandoc
echo.
echo Ejecutando: python -c "import pypandoc; pypandoc.download_pandoc()"
python -c "import pypandoc; pypandoc.download_pandoc()"

:: Paso 3: Eliminar el archivo .msi descargado
echo.
echo Eliminando el archivo .msi descargado...
del /q "%~dp0pandoc-*.msi"

echo.
echo ======================================================
echo ✅ Proceso finalizado.
echo Pandoc portable se ha descargado correctamente.
echo ======================================================
pause
exit /b
