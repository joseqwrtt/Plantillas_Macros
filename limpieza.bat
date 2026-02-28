@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

echo =============================================
echo 🧹 Limpieza de Python conflictivo / PATH
echo =============================================
echo.

:: ==== 1. Eliminar accesos directos de la Store ====
echo Comprobando Python de la Store...
for %%P in (%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe %LOCALAPPDATA%\Microsoft\WindowsApps\python3.exe) do (
    if exist "%%P" (
        echo ⚠ Se encontro Python de la Store en %%P
        echo ⚠ Eliminando acceso directo conflictivo...
        del /f /q "%%P" >nul 2>&1
        if not errorlevel 1 (
            echo 🟢 Eliminado correctamente.
        ) else (
            echo ❌ Error al eliminar (puede que no tengas permisos).
        )
    )
)

:: ==== 2. Limpiar posibles rutas antiguas de Python en el PATH de usuario ====
echo.
echo Comprobando PATH de usuario por entradas conflictivas de Python...
set "OLD_PATH=%PATH%"
set "NEW_PATH="
:: Tokenizar el PATH por punto y coma (;) para revisar cada directorio.
for %%D in (%PATH%) do (
    echo %%D | findstr /i "Python" >nul
    if errorlevel 1 (
        :: La ruta no contiene "Python", la conservamos.
        if defined NEW_PATH (
            set "NEW_PATH=!NEW_PATH!;%%D"
        ) else (
            set "NEW_PATH=%%D"
        )
    ) else (
        echo ⚠ Se elimino del PATH: %%D
    )
)

:: Usamos SETX para guardar el nuevo PATH de forma permanente en el registro del usuario.
setx PATH "!NEW_PATH!" >nul
echo 🟢 PATH de usuario actualizado.

:: ==== 3. Confirmacion final ====
echo.
echo =============================================
echo ✅ Limpieza completada. Reinicie CMD o la PC si es necesario.
echo =============================================
pause
endlocal
exit /b
