@echo off
:: IBM Industrial WFM — Compliance Launcher starter
:: Double-click this file to open the interactive launcher.

set "DIR=%~dp0"
set "HTA=%DIR%WFM_Launcher.hta"

if not exist "%HTA%" (
    echo ERROR: WFM_Launcher.hta not found in %DIR%
    pause
    exit /b 1
)

:: Launch the HTA application (Windows built-in engine, no install required)
start "" mshta.exe "%HTA%"
