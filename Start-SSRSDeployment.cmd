@echo off
:: ============================================================
:: Start-SSRSDeployment.cmd
:: ============================================================
:: Kopiert das Tool nach C:\ProgramData\SSRSDeployment und
:: startet es als Administrator (UAC).
::
:: Warum ProgramData?
::   - Nach UAC-Elevation ist W:\ nicht mehr erreichbar
::   - AppLocker/AV-unbedenklich
::   - Nicht von Cleanup-Scripts betroffen
::
:: Verwendung: Doppelklick vom Share genuegt.
:: ============================================================
setlocal EnableDelayedExpansion

set "SRCDIR=%~dp0"
set "LOCALDIR=%ProgramData%\SSRSDeployment"
set "LOCALPS=%LOCALDIR%\ReportDeplyment.ps1"

echo.
echo  Start-SSRSDeployment
echo  ============================================================
echo  Quelle : %SRCDIR%
echo  Ziel   : %LOCALDIR%
echo.

if not exist "%LOCALDIR%" (
    mkdir "%LOCALDIR%"
    if errorlevel 1 (
        echo  FEHLER: Verzeichnis konnte nicht angelegt werden: %LOCALDIR%
        pause
        exit /b 1
    )
)

xcopy /Y /Q /E "%SRCDIR%." "%LOCALDIR%\" >nul 2>&1
if errorlevel 1 (
    echo  FEHLER: Kopieren fehlgeschlagen.
    pause
    exit /b 1
)

echo  Dateien bereit - starte als Administrator ...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%LOCALPS%""' -Verb RunAs"

endlocal
