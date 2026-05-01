<#
    CsvToExcel_Fixed.ps1
    - 文字化け（UTF-8/Shift-JIS自動判定）の解消
    - 0落ち・指数表記の防止（全セルを文字列型として展開）
#>

param([string]$CsvPath)

Add-Type -AssemblyName System.Windows.Forms

if (-not $CsvPath -or -not (Test-Path $CsvPath)) {
    [System.Windows.Forms.MessageBox]::Show("ファイルパスが無効です。")
    exit 1
}

# --- 文字コード判定・テキスト取得（既存ロジックを継承） ---
function Read-CsvText {
    param([string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes)
    }
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
        return $utf8.GetString($bytes)
    } catch {
        return [System.Text.Encoding]::GetEncoding(932).GetString($bytes)
    }
}

try {
    $csvContent = Read-CsvText $CsvPath
    
    # --- Excel COM オブジェクトの生成 ---
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.ActiveSheet

    # --- 全セルを「文字列型(@)」に設定 ---
    # これにより、後から入力される数字の「0落ち」を防止する
    $sheet.Cells.NumberFormat = "@"

    # --- クリップボードを介した高速展開 ---
    # 1セルずつ書き込むと極めて遅いため、タブ区切り（TSV）形式でクリップボード経由で貼り付ける
    # CSVのカンマをタブに置換（簡易実装：データ内にカンマがある場合は考慮が必要）
    $tsvContent = $csvContent -replace ',', "`t"
    [System.Windows.Forms.Clipboard]::SetText($tsvContent)
    
    $sheet.Paste()
    [System.Windows.Forms.Clipboard]::Clear()

    # --- 列幅の自動調整 ---
    $sheet.Columns.AutoFit()

} catch {
    [System.Windows.Forms.MessageBox]::Show("エラーが発生しました。`n$($_.Exception.Message)")
    if ($excel) { $excel.Quit() }
    exit 1
}

# COMオブジェクトの解放（メモリ管理）
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

exit 0