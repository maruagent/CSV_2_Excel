<#
    CsvToExcel_backend.ps1
    ---------------------------------------------------------
    ■ 目的
      - CSV の文字コードを自動判定（UTF-8 BOM / UTF-8 / UTF-16 / Shift-JIS）
      - UTF-8 BOM に変換してから Excel で開く（文字化け防止）
      - 一時ファイルに GUID を付与し衝突を防止
#>

param([string]$CsvPath)

# --- 0. GUIライブラリの読み込み ---
Add-Type -AssemblyName System.Windows.Forms

# --- 1. 入力ファイル存在チェック ---
if (-not $CsvPath -or -not (Test-Path $CsvPath)) {
    [System.Windows.Forms.MessageBox]::Show("ファイルが見つかりません。`n対象のCSVをバッチファイルにドロップしてください。")
    exit
}

# --- 2. 文字コード自動判定関数 ---
function Detect-And-ReadText {
    param([string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)

    # UTF-8 BOMあり
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes)
    }
    # UTF-16 LE
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode.GetString($bytes)
    }

    # UTF-8 (BOMなし) か Shift-JIS かを判定
    try {
        $utf8NoBOM = New-Object System.Text.UTF8Encoding($false, $true)
        return $utf8NoBOM.GetString($bytes)
    } catch {
        # UTF-8として失敗した場合は Shift-JIS(CP932) で読み込む
        $sjis = [System.Text.Encoding]::GetEncoding(932)
        return $sjis.GetString($bytes)
    }
}

# --- 3. 実行処理 ---
try {
    # テキストの読み込み
    $text = Detect-And-ReadText $CsvPath

    # 一時ファイルのパス作成 (TEMPフォルダ + GUID：ランダム値)
    $tempDir = [System.IO.Path]::GetTempPath()
    $guid = [guid]::NewGuid().ToString().Substring(0,8)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
    $tempPath = Join-Path $tempDir "${baseName}_${guid}.csv"

    # Excelが確実に認識する「UTF-8 BOM付き」で保存
    $utf8bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($tempPath, $text, $utf8bom)

    # Excelで開く
    # Start-Processを使うことで、Excelのインストール場所を問わず起動可能
    Start-Process "excel.exe" -ArgumentList "`"$tempPath`""

} catch {
    [System.Windows.Forms.MessageBox]::Show("処理中にエラーが発生しました。`n$($_.Exception.Message)")
}