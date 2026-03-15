@echo off
chcp 65001 >nul
title Limpiar Trial - Plantillas Macro

echo.
echo ================================================
echo   Limpieza de registros de prueba
echo   Plantillas Macro - Solo para desarrollo
echo ================================================
echo.
echo Borrando entradas del registro...

:: Registro 1: Office UserInfo
reg delete "HKCU\Software\Microsoft\Office\16.0\Common\UserInfo" /v "RecentTemplateTS" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Common\UserInfo" /v "RecentTemplateID" /f >nul 2>&1
echo   [1/5] Office UserInfo - OK

:: Registro 2: .NET Framework AppPerf
reg delete "HKCU\Software\Microsoft\.NETFramework\Policy\AppPerf" /v "InitStamp" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\.NETFramework\Policy\AppPerf" /v "PerfIndex" /f >nul 2>&1
echo   [2/5] .NET Framework AppPerf - OK

:: Registro 3: Explorer UserAssist
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\Config" /v "SessionToken" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\Config" /v "SessionSeq" /f >nul 2>&1
echo   [3/5] Explorer UserAssist - OK

:: Registro 4: Visual Studio Setup
reg delete "HKCU\Software\Microsoft\VisualStudio\14.0\Setup\VS" /v "BuildTimestamp" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\VisualStudio\14.0\Setup\VS" /v "BuildRevision" /f >nul 2>&1
echo   [4/5] Visual Studio Setup - OK

:: Registro 5: Licencia activada (WindowsUpdate AppCategories)
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RequestedAppCategories" /v "AppID" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RequestedAppCategories" /v "CategoryGUID" /f >nul 2>&1
echo   [5/5] WindowsUpdate AppCategories - OK

echo.
echo Borrando archivos de datos...

set "AJUSTES=%~dp0Ajustes"

set "TRIAL=%AJUSTES%\trial.dat"
if exist "%TRIAL%" (
    del /f /q "%TRIAL%" >nul 2>&1
    echo   [+] trial.dat borrado
) else (
    echo   [+] trial.dat no encontrado
)

set "LIC=%AJUSTES%\licencia.dat"
if exist "%LIC%" (
    del /f /q "%LIC%" >nul 2>&1
    echo   [+] licencia.dat borrado
) else (
    echo   [+] licencia.dat no encontrado
)

echo.
echo ================================================
echo   Listo - El trial se reiniciara en el
echo   proximo arranque de la aplicacion
echo ================================================
echo.
pause
