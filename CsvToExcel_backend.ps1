<#
    CsvToExcel_backend.ps1
    ---------------------------------------------------------
    ■ 目的
      - CSV の文字コードを自動判定（UTF-8 BOM / UTF-8 / UTF-16 / Shift-JIS）
      - UTF-8 BOM に変換してから Excel で開く（文字化け防止）
      - 一時ファイルに GUID を付与し衝突を防止

    ■ 修正履歴
      - $pid 予約変数の衝突を解消（$targetPid に変更）
      - 監視スクリプトを -EncodedCommand で渡すよう変更
      - UTF-8 / Shift-JIS 判定を厳密デコード＋フォールバック方式に修正
      - Excel 起動失敗時のエラーハンドリングを追加
      - exit の終了コードを統一（正常:0 / 異常:1）
      - 関数名を承認済み動詞に準拠（Read-CsvText）
      - GUID をフル長（32文字）に変更
#>

param([string]$CsvPath)

# --- 0. GUI ライブラリの読み込み ---
Add-Type -AssemblyName System.Windows.Forms

# --- 1. 入力ファイル存在チェック ---
if (-not $CsvPath -or -not (Test-Path $CsvPath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "ファイルが見つかりません。`n対象の CSV をバッチファイルにドロップしてください。"
    )
    exit 1
}

# --- 2. 文字コード自動判定関数 ---
# BOM なし UTF-8 と Shift-JIS の判定：
#   UTF-8Encoding($false, $true) で厳密デコードを試みる。
#   Shift-JIS のバイト列は UTF-8 として不正なシーケンスを含む場合に例外を発生させるため、
#   例外発生時のみ Shift-JIS にフォールバックする。
#   ※ ASCII のみのファイルは例外が発生しないが、日本語を含まない CSV では
#      どちらで読んでも結果が同じため実害はない。
function Read-CsvText {
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)

    # UTF-8 BOM あり（EF BB BF）
    if ($bytes.Length -ge 3 -and
        $bytes[0] -eq 0xEF -and
        $bytes[1] -eq 0xBB -and
        $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes)
    }

    # UTF-16 LE BOM あり（FF FE）
    if ($bytes.Length -ge 2 -and
        $bytes[0] -eq 0xFF -and
        $bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode.GetString($bytes)
    }

    # BOM なし UTF-8 として厳密デコードを試みる（throwOnInvalidBytes: $true）
    # Shift-JIS 固有バイトが含まれる場合は DecoderFallbackException が発生する
    try {
        $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
        return $utf8Strict.GetString($bytes)
    } catch {
        # UTF-8 として不正なバイト列 → Shift-JIS（CP932）で読み込む
        $sjis = [System.Text.Encoding]::GetEncoding(932)
        return $sjis.GetString($bytes)
    }
}

# --- 3. 実行処理 ---
try {
    # テキストの読み込み
    $text = Read-CsvText $CsvPath

    # 一時ファイルのパス作成（TEMP フォルダ + フル GUID）
    $tempDir  = [System.IO.Path]::GetTempPath()
    $guid     = [guid]::NewGuid().ToString("N")   # ハイフンなし 32 文字
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
    $tempPath = Join-Path $tempDir "${baseName}_${guid}.csv"

    # Excel が確実に認識する「UTF-8 BOM 付き」で保存
    $utf8bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($tempPath, $text, $utf8bom)

    # Excel で開く（-ErrorAction Stop で起動失敗を catch へ誘導）
    $excelProcess = Start-Process "excel.exe" `
        -ArgumentList "`"$tempPath`"" `
        -PassThru `
        -ErrorAction Stop

    if (-not $excelProcess) {
        throw "Excel の起動に失敗しました。Excel がインストールされているか確認してください。"
    }

    # --- 4. 一時ファイルの自動削除設定（独立した監視プロセス） ---
    # Excel が閉じられたことを検知して一時ファイルを削除する。
    # メインプロセス終了後も動作し続けるよう独立プロセスとして起動する。
    # 特殊文字・改行の問題を回避するため -EncodedCommand で渡す。
    $monitorScript = @"
        `$targetPath = '$tempPath'
        `$targetPid  = $($excelProcess.Id)
        while (Get-Process -Id `$targetPid -ErrorAction SilentlyContinue) {
            Start-Sleep -Seconds 10
        }
        `$retryCount = 0
        while ((Test-Path `$targetPath) -and (`$retryCount -lt 12)) {
            try {
                Remove-Item `$targetPath -Force -ErrorAction Stop
                break
            } catch {
                Start-Sleep -Seconds 10
                `$retryCount++
            }
        }
"@

    $encodedCommand = [Convert]::ToBase64String(
        [System.Text.Encoding]::Unicode.GetBytes($monitorScript)
    )

    Start-Process powershell.exe `
        -ArgumentList "-NoProfile -WindowStyle Hidden -EncodedCommand $encodedCommand"

} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "処理中にエラーが発生しました。`n$($_.Exception.Message)"
    )
    exit 1
}

exit 0