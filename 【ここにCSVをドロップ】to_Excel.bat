@echo off
setlocal
title CSV to Excel Converter

REM --- Check if file is dropped ---
if "%~1" == "" (
    color 0C
    echo [ERROR] No file detected.
    echo Please drop a CSV file onto this icon.
    pause
    exit /b
)

REM --- Processing ---
echo Processing: %~nx1...

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0CsvToExcel_backend.ps1" "%~1" >nul 2>&1

REM --- Result Check ---
if %errorlevel% equ 0 (
    echo [SUCCESS] Opening Excel...
    timeout /t 2 >nul
) else (
    color 0C
    echo [ERROR] PowerShell script failed.
    pause
)

exit /b