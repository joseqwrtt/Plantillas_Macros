@echo off
title Instalador de Plantillas - Dependencias Python y Pandoc
color 0A
chcp 65001 >nul
setlocal enabledelayedexpansion

echo =============================================
echo   🧩 Instalador de Plantillas - Dependencias
echo =============================================
echo.

:: === RUTAS PRINCIPALES ===
set "BASE_DIR=%~dp0"
set "AJUSTES_DIR=%BASE_DIR%Ajustes"
set "EDITOR_DIR=%BASE_DIR%Editor"
set "LOG_FILE=%AJUSTES_DIR%\install_log.txt"
set "PYPANDOC_INSTALL=%BASE_DIR%pypandoc_Install.bat"

if not exist "%AJUSTES_DIR%" mkdir "%AJUSTES_DIR%"
echo 🕒 Inicio instalación dependencias: %date% %time% > "%LOG_FILE%"

:: === LIMPIEZA LIGERA ===
echo.
echo 🧹 Limpiando archivos temporales antiguos...
for %%F in ("%BASE_DIR%*.tmp" "%BASE_DIR%*.bak" "%BASE_DIR%*.old") do (
    if exist "%%~fF" (
        del "%%~fF" >nul 2>&1
        echo Eliminado: %%~nxF
    )
)
if exist "%EDITOR_DIR%\ajustes\temp\" (
    echo Limpiando carpeta temp...
    del /q "%EDITOR_DIR%\ajustes\temp\*" >nul 2>&1
)
if exist "%EDITOR_DIR%\ajustes\pandoc.zip" del "%EDITOR_DIR%\ajustes\pandoc.zip" >nul 2>&1
echo 🧼 Limpieza completada.
echo 🕒 %date% %time% - 🧹 Limpieza ligera ejecutada >> "%LOG_FILE%"

:: === 1. Verificar Python ===
echo.
echo Verificando Python...
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ❌ Python no encontrado. Se abrirá el instalador.
    start "" "%AJUSTES_DIR%\python-3.13.2-amd64.exe"
    echo 🕒 %date% %time% - Instalador Python lanzado >> "%LOG_FILE%"
    pause
) else (
    for /f "tokens=*" %%i in ('where python') do set PYTHON_EXE=%%i
    for /f "tokens=*" %%v in ('%PYTHON_EXE% --version') do set PY_VER=%%v
    echo ✅ Python detectado: %PY_VER%
    echo 🕒 %date% %time% - ✅ Python detectado: %PY_VER% (%PYTHON_EXE%) >> "%LOG_FILE%"
)

:: === 2. Actualizar pip ===
echo.
echo Actualizando pip...
%PYTHON_EXE% -m pip install --upgrade pip >nul 2>&1
if %errorlevel%==0 (
    echo ✅ pip actualizado correctamente.
    echo 🕒 %date% %time% - ✅ pip actualizado >> "%LOG_FILE%"
) else (
    echo ⚠ No se pudo actualizar pip.
    echo 🕒 %date% %time% - ⚠ pip fallo al actualizar >> "%LOG_FILE%"
)

:: === 3. Instalar dependencias ===
set "DEPS=Flask pywin32 python-docx win10toast pypandoc"

echo.
echo 📦 Instalando dependencias requeridas...
for %%D in (%DEPS%) do (
    %PYTHON_EXE% -m pip show %%D >nul 2>&1
    if errorlevel 1 (
        echo [Instalando] %%D...
        %PYTHON_EXE% -m pip install %%D >nul 2>&1
        echo 🕒 %date% %time% - ✅ %%D instalado >> "%LOG_FILE%"
    ) else (
        echo [OK] %%D ya está instalado.
        echo 🕒 %date% %time% - ⚙ %%D ya presente >> "%LOG_FILE%"
    )
)

:: === 4. Verificar Pandoc local ===
echo.
echo Verificando Pandoc...
set "PANDOC_EXE=%EDITOR_DIR%\ajustes\pandoc\pandoc.exe"
if exist "%PANDOC_EXE%" (
    echo ✅ Pandoc detectado en ruta local.
    echo 🕒 %date% %time% - ⚙ Pandoc detectado localmente >> "%LOG_FILE%"
) else (
    echo ⚠ Pandoc no encontrado, se instalará con pypandoc...
)

:: === 5. Ejecutar instalador de pypandoc (descarga real de Pandoc) ===
if exist "%PYPANDOC_INSTALL%" (
    echo.
    echo 🚀 Ejecutando instalador de pypandoc para descarga de Pandoc...
    call "%PYPANDOC_INSTALL%"
    echo 🕒 %date% %time% - ✅ pypandoc_Install.bat ejecutado >> "%LOG_FILE%"
) else (
    echo ❌ No se encontró pypandoc_Install.bat en la raíz.
    echo 🕒 %date% %time% - ⚠ No se ejecutó pypandoc_Install.bat >> "%LOG_FILE%"
)

echo.
echo =============================================
echo ✅ Instalación completada correctamente.
echo 📄 Log disponible en: %LOG_FILE%
echo =============================================
pause
endlocal
exit /b
