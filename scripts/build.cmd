@echo off
setlocal

@echo Build Add-on Manager

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"

if %errorlevel% neq 0 (
    echo.
    echo ========================================
    echo   BUILD FAILED
    echo ========================================
    echo.
)

if not "%1"=="/nopause" pause
