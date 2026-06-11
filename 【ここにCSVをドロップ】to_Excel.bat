@echo off
setlocal EnableExtensions

title CSV to Excel Converter

set "SCRIPT=%~dp0CsvToExcel_backend.ps1"
set "LOG=%TEMP%\CsvToExcel_%RANDOM%%RANDOM%.log"

if "%~1"=="" (
    color 0C
    echo [ERROR] CSV file was not specified.
    echo Please drag and drop a CSV file onto this BAT file.
    echo.
    pause
    exit /b 1
)

if not exist "%~1" (
    color 0C
    echo [ERROR] The specified file was not found.
    echo "%~1"
    echo.
    pause
    exit /b 1
)

if /i not "%~x1"==".csv" (
    color 0C
    echo [ERROR] Only CSV files are supported.
    echo File: "%~nx1"
    echo.
    pause
    exit /b 1
)

if not exist "%SCRIPT%" (
    color 0C
    echo [ERROR] Backend PowerShell script was not found.
    echo "%SCRIPT%"
    echo.
    pause
    exit /b 1
)

echo Processing: "%~nx1"
echo Log file: "%LOG%"
echo.

powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%SCRIPT%" -CsvPath "%~f1" -LogPath "%LOG%"

if errorlevel 1 (
    color 0C
    echo.
    echo [ERROR] Failed to open the CSV file in Excel.
    echo Please check the log file below.
    echo "%LOG%"
    echo.
    pause
    exit /b 1
)

echo.
echo [SUCCESS] The CSV file has been opened in Excel.
echo All columns are imported as text.
echo.
timeout /t 2 >nul
exit /b 0