<#
    CsvToExcel_backend.ps1

    Purpose:
    - Open a CSV file in Excel.
    - Detect UTF-8 BOM, UTF-8, or Shift_JIS.
    - Parse CSV safely, including quoted commas.
    - Import all columns as text.
    - Avoid clipboard usage.
    - Write logs for troubleshooting.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[{0}] {1}" -f $timestamp, $Message

    if ($LogPath) {
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
}

function Show-Message {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Information", "Warning", "Error")]
        [string]$Icon = "Information"
    )

    $iconValue = [System.Windows.Forms.MessageBoxIcon]::$Icon

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        "CSV to Excel Converter",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $iconValue
    ) | Out-Null
}

function Get-CsvEncodingInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $bytes = [System.IO.File]::ReadAllBytes($Path)

    if ($bytes.Length -ge 3 -and
        $bytes[0] -eq 0xEF -and
        $bytes[1] -eq 0xBB -and
        $bytes[2] -eq 0xBF) {

        return [pscustomobject]@{
            Name     = "UTF-8 BOM"
            Encoding = New-Object System.Text.UTF8Encoding($true, $true)
        }
    }

    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
        [void]$utf8.GetString($bytes)

        return [pscustomobject]@{
            Name     = "UTF-8"
            Encoding = $utf8
        }
    }
    catch {
        return [pscustomobject]@{
            Name     = "Shift_JIS"
            Encoding = [System.Text.Encoding]::GetEncoding(932)
        }
    }
}

function Read-CsvRows {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [System.Text.Encoding]$Encoding
    )

    $parser = $null
    $rows = New-Object 'System.Collections.Generic.List[object[]]'
    $maxColumns = 0

    try {
        $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, $Encoding)
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(",")
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false

        while (-not $parser.EndOfData) {
            [object[]]$fields = $parser.ReadFields()

            if ($null -eq $fields) {
                [object[]]$fields = @("")
            }

            $rows.Add($fields)

            if ($fields.Count -gt $maxColumns) {
                $maxColumns = $fields.Count
            }
        }

        if ($rows.Count -eq 0 -or $maxColumns -eq 0) {
            throw "CSV file is empty or no columns were detected."
        }

        return [pscustomobject]@{
            Rows       = $rows
            RowCount   = $rows.Count
            MaxColumns = $maxColumns
        }
    }
    finally {
        if ($null -ne $parser) {
            $parser.Close()
            $parser.Dispose()
        }
    }
}

function Release-ComObjectSafely {
    param(
        [Parameter(Mandatory = $false)]
        [object]$ComObject
    )

    if ($null -ne $ComObject) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
        }
        catch {
            Write-Log ("Failed to release COM object: {0}" -f $_.Exception.Message)
        }
    }
}

$excel = $null
$workbooks = $null
$workbook = $null
$worksheets = $null
$worksheet = $null
$range = $null
$openedSuccessfully = $false

try {
    if (-not $LogPath) {
        $LogPath = Join-Path -Path $env:TEMP -ChildPath ("CsvToExcel_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    }

    Write-Log "===== Start CSV to Excel Converter ====="
    Write-Log ("CSV path: {0}" -f $CsvPath)
    Write-Log ("Log path: {0}" -f $LogPath)

    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
        throw "The specified CSV file was not found."
    }

    $extension = [System.IO.Path]::GetExtension($CsvPath)

    if ($extension.ToLowerInvariant() -ne ".csv") {
        throw "Only CSV files are supported."
    }

    $resolvedCsvPath = (Resolve-Path -LiteralPath $CsvPath).Path
    Write-Log ("Resolved CSV path: {0}" -f $resolvedCsvPath)

    $encodingInfo = Get-CsvEncodingInfo -Path $resolvedCsvPath
    Write-Log ("Detected encoding: {0}" -f $encodingInfo.Name)

    $csvData = Read-CsvRows -Path $resolvedCsvPath -Encoding $encodingInfo.Encoding

    $rowCount = [int]$csvData.RowCount
    $maxColumns = [int]$csvData.MaxColumns

    Write-Log ("Detected rows: {0}, max columns: {1}" -f $rowCount, $maxColumns)

    if ($rowCount -gt 1048576) {
        throw "CSV has more than 1,048,576 rows. Excel cannot display all rows."
    }

    if ($maxColumns -gt 16384) {
        throw "CSV has more than 16,384 columns. Excel cannot display all columns."
    }

    Write-Log "Creating two-dimensional array for Excel."

    $data = [object[,]]::new($rowCount, $maxColumns)

    for ($r = 0; $r -lt $rowCount; $r++) {
        $fields = $csvData.Rows[$r]

        for ($c = 0; $c -lt $maxColumns; $c++) {
            if ($c -lt $fields.Count -and $null -ne $fields[$c]) {
                $data[$r, $c] = [string]$fields[$c]
            }
            else {
                $data[$r, $c] = ""
            }
        }
    }

    Write-Log "Starting Excel."

    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        throw "Excel could not be started. Please confirm that Microsoft Excel is installed."
    }

    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Add()
    $worksheets = $workbook.Worksheets
    $worksheet = $worksheets.Item(1)

    $worksheet.Name = "CSV"

    Write-Log "Writing CSV data to Excel."

    $range = $worksheet.Range(
        $worksheet.Cells.Item(1, 1),
        $worksheet.Cells.Item($rowCount, $maxColumns)
    )

    # Import all cells as text to prevent leading zero loss and scientific notation.
    $range.NumberFormat = "@"
    $range.Value2 = $data

    $worksheet.Columns.AutoFit() | Out-Null

    $excel.DisplayAlerts = $true
    $excel.Visible = $true
    $openedSuccessfully = $true

    Write-Log "CSV file was opened successfully."
    Write-Log "===== End CSV to Excel Converter: Success ====="

    exit 0
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Log ("ERROR: {0}" -f $errorMessage)
    Write-Log "===== End CSV to Excel Converter: Failed ====="

    try {
        if ($null -ne $workbook -and -not $openedSuccessfully) {
            $workbook.Close($false) | Out-Null
        }
    }
    catch {
        Write-Log ("Failed to close workbook: {0}" -f $_.Exception.Message)
    }

    try {
        if ($null -ne $excel -and -not $openedSuccessfully) {
            $excel.Quit()
        }
    }
    catch {
        Write-Log ("Failed to quit Excel: {0}" -f $_.Exception.Message)
    }

    Show-Message -Icon "Error" -Message ("CSVファイルをExcelで開けませんでした。`n`n原因：{0}`n`nログ：{1}" -f $errorMessage, $LogPath)

    exit 1
}
finally {
    Release-ComObjectSafely -ComObject $range
    Release-ComObjectSafely -ComObject $worksheet
    Release-ComObjectSafely -ComObject $worksheets
    Release-ComObjectSafely -ComObject $workbook
    Release-ComObjectSafely -ComObject $workbooks
    Release-ComObjectSafely -ComObject $excel

    $range = $null
    $worksheet = $null
    $worksheets = $null
    $workbook = $null
    $workbooks = $null
    $excel = $null

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}