@echo off
rem Double-click this file to open PMRS with the Excel extension loaded (Edge).
rem Pass "chrome" as the first argument to use Chrome instead.
cd /d "%~dp0"
if /i "%~1"=="chrome" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run-pmrs.ps1" -Browser Chrome
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run-pmrs.ps1" -Browser Edge
)
if errorlevel 1 pause
