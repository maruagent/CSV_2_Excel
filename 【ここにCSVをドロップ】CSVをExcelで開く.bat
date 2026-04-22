@echo off
setlocal
REM ---------------------------------------------------------
REM  CSVを PowerShell スクリプトに渡して Excel で開く
REM  - ExecutionPolicy Bypass で実行制限を回避
REM  - 窓を隠して実行
REM ---------------------------------------------------------

set "PS_FILE=%~dp0CsvToExcel_backend.ps1"

if "%~1"=="" (
    echo 【エラー】CSVファイルをこのアイコンにドロップしてください。
    pause
    exit /b
)

powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "%PS_FILE%" "%~1"

exit /b