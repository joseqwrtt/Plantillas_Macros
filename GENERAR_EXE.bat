@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title Compilador Plantillas Macro

echo.
echo ===================================================
echo    COMPILADOR - Plantillas Macro v3.0
echo    Ofuscacion + Empaquetado con PyInstaller
echo ===================================================
echo.

:: -- Verificar Python ------------------------------------------------------
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado.
    pause & exit /b 1
)

:: -- Instalar dependencias -------------------------------------------------
echo [1/5] Instalando dependencias...
pip install pyinstaller pyarmor customtkinter --quiet --upgrade
if errorlevel 1 ( echo [ERROR] Fallo dependencias. & pause & exit /b 1 )
echo       OK

:: -- Rutas (con comillas para manejar espacios en el path) -----------------
set "SRC=%~dp0"
if "%SRC:~-1%"=="\" set "SRC=%SRC:~0,-1%"

set "BUILD=%SRC%\_build"
set "DIST=%SRC%\dist"
set "OBFUSC=%SRC%\_obfuscated"

if exist "%BUILD%"   rmdir /s /q "%BUILD%"
if exist "%DIST%"    rmdir /s /q "%DIST%"
if exist "%OBFUSC%"  rmdir /s /q "%OBFUSC%"
mkdir "%BUILD%"
mkdir "%OBFUSC%"

:: -- Copiar fuentes al directorio de trabajo -------------------------------
echo [2/5] Preparando fuentes...
for %%F in (main.py core.py core_odt.py interfaz_ctk.py interfaz_base.py licencia.py ventana_licencia.py) do (
    if exist "%SRC%\%%F" (
        copy /y "%SRC%\%%F" "%BUILD%\%%F" >nul
        echo       Copiado: %%F
    ) else (
        echo [AVISO] No encontrado: %%F - continuando sin el
    )
)

:: -- Eliminar bloque TEST_FORZAR_MODO de licencia.py -----------------------
echo       Eliminando codigo de pruebas de licencia.py...
python "%SRC%\_strip_test.py" "%BUILD%\licencia.py"
if errorlevel 1 ( echo       [AVISO] No se pudo limpiar TEST_FORZAR_MODO ) else ( echo       OK )

echo       OK

:: -- Ofuscacion con PyArmor ------------------------------------------------
echo [3/5] Ofuscando codigo con PyArmor...
cd /d "%BUILD%"

pyarmor gen --output "%OBFUSC%" --obf-module 1 --obf-code 1 main.py core.py core_odt.py interfaz_ctk.py interfaz_base.py licencia.py ventana_licencia.py

if errorlevel 1 (
    echo [AVISO] PyArmor fallo - se compilara sin ofuscacion
    set "COMPILE_DIR=%BUILD%"
) else (
    echo       OK - Codigo ofuscado correctamente
    set "COMPILE_DIR=%OBFUSC%"
)

:: -- Detectar carpeta runtime de PyArmor (nombre varia segun version) ------
set "RUNTIME_OPT="
for /d %%D in ("%OBFUSC%\pyarmor_runtime*") do (
    echo       Runtime PyArmor: %%~nxD
    set "RUNTIME_OPT=--add-data "%%D;%%~nxD""
)

:: -- Detectar icono opcional -----------------------------------------------
set "ICON_OPT="
if exist "%SRC%\icon.ico" (
    set "ICON_OPT=--icon="%SRC%\icon.ico""
    echo       Icono encontrado: icon.ico
)

:: -- Compilar con PyInstaller ----------------------------------------------
echo [4/5] Generando ejecutable con PyInstaller...
cd /d "%COMPILE_DIR%"

pyinstaller ^
    --onefile ^
    --windowed ^
    --name "PlantillasMacro" ^
    --distpath "%DIST%" ^
    --workpath "%BUILD%\pyi_work" ^
    --specpath "%BUILD%" ^
    %ICON_OPT% ^
    %RUNTIME_OPT% ^
    --hidden-import=customtkinter ^
    --hidden-import=win32clipboard ^
    --hidden-import=win32api ^
    --hidden-import=win32con ^
    --hidden-import=pywintypes ^
    --hidden-import=keyboard ^
    --hidden-import=docx ^
    --hidden-import=lxml ^
    --hidden-import=lxml.etree ^
    --hidden-import=odf ^
    --hidden-import=odf.opendocument ^
    --collect-all customtkinter ^
    --noconfirm ^
    main.py

if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller fallo. Revisa los mensajes arriba.
    cd /d "%SRC%"
    pause & exit /b 1
)

:: -- Limpiar temporales ----------------------------------------------------
echo [5/5] Limpiando temporales...
cd /d "%SRC%"
rmdir /s /q "%BUILD%"   >nul 2>&1
rmdir /s /q "%OBFUSC%"  >nul 2>&1
echo       OK

:: -- Mover EXE a la carpeta raiz y borrar dist ---------------------------
echo       Moviendo EXE a carpeta raiz...
move /y "%DIST%\PlantillasMacro.exe" "%SRC%\PlantillasMacro.exe" >nul
if errorlevel 1 (
    echo [ERROR] No se pudo mover el EXE.
    pause & exit /b 1
)
rmdir /s /q "%DIST%" >nul 2>&1
echo       OK

:: -- Resultado -------------------------------------------------------------
echo.
echo ===================================================
echo    COMPILACION COMPLETADA
echo    EXE: %SRC%\PlantillasMacro.exe
echo ===================================================
echo.
explorer /select,"%SRC%\PlantillasMacro.exe"
pause