#requires -Version 5.1
<#
  Ouvre un CSV dans Excel avec détection automatique :
  encodage (BOM, UTF-8 ou ANSI), séparateur (; , tab |, ligne "sep=x"),
  séparateur décimal, colonnes à garder en texte (zéros en tête, numéros longs).
#>
param([Parameter(Mandatory = $true)][string]$Path)

$ErrorActionPreference = 'Stop'
$Candidates    = @(';', ',', "`t", '|')
$SampleBytes   = 1MB
$SampleRecords = 200

function Read-Sample([string]$File) {
    $fs = [IO.File]::OpenRead($File)
    try {
        $buf  = New-Object byte[] ([int][Math]::Min($SampleBytes, $fs.Length))
        $read = $fs.Read($buf, 0, $buf.Length)
        [pscustomobject]@{ Bytes = $buf; Length = $read; Truncated = $fs.Length -gt $read }
    } finally { $fs.Dispose() }
}

function Get-CodePage($Sample) {
    $b = $Sample.Bytes
    $n = $Sample.Length
    if ($n -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) { return 65001 }
    if ($n -ge 2 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE) { return 1200 }
    if ($n -ge 2 -and $b[0] -eq 0xFE -and $b[1] -eq 0xFF) { return 1201 }
    if ($n -eq 0) { return 65001 }
    if ($Sample.Truncated) {
        $nl = [Array]::LastIndexOf($b, [byte]10, $n - 1)
        if ($nl -gt 0) { $n = $nl }
    }
    try {
        [void](New-Object Text.UTF8Encoding($false, $true)).GetString($b, 0, $n)
        65001
    } catch {
        [Text.Encoding]::Default.CodePage
    }
}

function Get-Records([string]$Text, [bool]$Truncated) {
    $records  = New-Object Collections.Generic.List[string]
    $sb       = New-Object Text.StringBuilder
    $inQuotes = $false
    foreach ($c in $Text.ToCharArray()) {
        if ($c -eq '"') { $inQuotes = -not $inQuotes }
        if (-not $inQuotes -and ($c -eq "`n" -or $c -eq "`r")) {
            if ($sb.Length -gt 0) { $records.Add($sb.ToString()); [void]$sb.Clear() }
            if ($records.Count -ge $SampleRecords) { return , $records }
            continue
        }
        [void]$sb.Append($c)
    }
    if ($sb.Length -gt 0 -and -not $Truncated) { $records.Add($sb.ToString()) }
    , $records
}

function Measure-Char([string]$Record, [char]$Char) {
    $n = 0; $q = $false
    foreach ($c in $Record.ToCharArray()) {
        if ($c -eq '"') { $q = -not $q }
        elseif (-not $q -and $c -eq $Char) { $n++ }
    }
    $n
}

function Split-Record([string]$Record, [char]$Delim) {
    $fields = New-Object Collections.Generic.List[string]
    $sb = New-Object Text.StringBuilder; $q = $false
    foreach ($c in $Record.ToCharArray()) {
        if ($c -eq '"') { $q = -not $q; continue }
        if (-not $q -and $c -eq $Delim) { $fields.Add($sb.ToString()); [void]$sb.Clear(); continue }
        [void]$sb.Append($c)
    }
    $fields.Add($sb.ToString())
    , $fields
}

function Find-Delimiter($Records) {
    $best = [Globalization.CultureInfo]::CurrentCulture.TextInfo.ListSeparator
    $bestScore = -1
    foreach ($d in $Candidates) {
        $counts = @($Records | ForEach-Object { Measure-Char $_ $d })
        if ($counts.Count -eq 0 -or $counts[0] -eq 0) { continue }
        $consistent = @($counts | Where-Object { $_ -eq $counts[0] }).Count
        $score = $consistent * 1000 + $counts[0]
        if ($score -gt $bestScore) { $best = $d; $bestScore = $score }
    }
    $best
}

function Find-DecimalSeparator($Rows, [string]$Delim) {
    $dot = 0; $comma = 0
    foreach ($row in $Rows) {
        foreach ($f in $row) {
            $v = $f.Trim()
            if ($v -match '^[-+]?\d+\.\d+$') { $dot++ }
            elseif ($v -match '^[-+]?\d+,\d+$') { $comma++ }
        }
    }
    if ($dot -gt $comma) { return '.' }
    if ($comma -gt $dot) { return ',' }
    if ($Delim -eq ',') { '.' } else { ',' }
}

function Get-ColumnTypes($Rows) {
    $max = ($Rows | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    if (-not $max) { return $null }
    $types = New-Object object[] $max
    for ($i = 0; $i -lt $max; $i++) { $types[$i] = 1 }          # xlGeneralFormat
    foreach ($row in $Rows) {
        for ($i = 0; $i -lt $row.Count; $i++) {
            $v = $row[$i].Trim()
            if ($v -match '^0\d+$' -or $v -match '^\+?\d{16,}$' -or $v -match '^\+\d') {
                $types[$i] = 2                                    # xlTextFormat
            }
        }
    }
    , $types
}

function Get-CsvSettings([string]$File) {
    $sample = Read-Sample $File
    $cp     = Get-CodePage $sample
    $text   = [Text.Encoding]::GetEncoding($cp).GetString($sample.Bytes, 0, $sample.Length).TrimStart([char]0xFEFF)
    $records  = Get-Records $text $sample.Truncated
    $startRow = 1
    if ($records.Count -gt 0 -and $records[0] -match '^sep=(.)$') {
        $delim = $Matches[1]
        $records.RemoveAt(0)
        $startRow = 2
    } else {
        $delim = Find-Delimiter $records
    }
    $rows = @($records | ForEach-Object { , (Split-Record $_ $delim) })
    [pscustomobject]@{
        CodePage  = $cp
        Delimiter = $delim
        StartRow  = $startRow
        Decimal   = Find-DecimalSeparator $rows $delim
        Types     = Get-ColumnTypes $rows
    }
}

function Get-Excel {
    try { [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
    catch { New-Object -ComObject Excel.Application }
}

function Get-SheetName([string]$File) {
    $name = [IO.Path]::GetFileNameWithoutExtension($File) -replace '[\[\]:*?/\\]', '_'
    if (-not $name) { $name = 'CSV' }
    if ($name.Length -gt 31) { $name = $name.Substring(0, 31) }
    $name
}

function Open-InExcel([string]$File, $Cfg) {
    $xl = Get-Excel
    $xl.Visible = $true
    $xl.UserControl = $true
    $wb = $xl.Workbooks.Add(-4167)                               # xlWBATWorksheet
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = Get-SheetName $File

    $qt = $ws.QueryTables.Add("TEXT;$File", $ws.Range('A1'))
    $qt.TextFilePlatform             = $Cfg.CodePage
    $qt.TextFileStartRow             = $Cfg.StartRow
    $qt.TextFileParseType            = 1                         # xlDelimited
    $qt.TextFileTextQualifier        = 1                         # xlTextQualifierDoubleQuote
    $qt.TextFileConsecutiveDelimiter = $false
    $qt.TextFileTabDelimiter         = $Cfg.Delimiter -eq "`t"
    $qt.TextFileSemicolonDelimiter   = $Cfg.Delimiter -eq ';'
    $qt.TextFileCommaDelimiter       = $Cfg.Delimiter -eq ','
    $qt.TextFileSpaceDelimiter       = $false
    if ($Candidates[0..2] -notcontains $Cfg.Delimiter) { $qt.TextFileOtherDelimiter = $Cfg.Delimiter }
    $qt.TextFileDecimalSeparator     = $Cfg.Decimal
    $qt.TextFileThousandsSeparator   = $(if ($Cfg.Decimal -eq '.') { ',' } else { ' ' })
    if ($Cfg.Types) { $qt.TextFileColumnDataTypes = $Cfg.Types }
    $qt.AdjustColumnWidth            = $true
    [void]$qt.Refresh($false)

    # Données figées : plus de lien vers le fichier source
    $qt.Delete()
    for ($i = $wb.Connections.Count; $i -ge 1; $i--) { $wb.Connections.Item($i).Delete() }
    for ($i = $wb.Names.Count; $i -ge 1; $i--) { $wb.Names.Item($i).Delete() }

    $wb.Windows.Item(1).Caption = [IO.Path]::GetFileName($File)
    $ws.Range('A1').Select() | Out-Null
    [void](New-Object -ComObject WScript.Shell).AppActivate($xl.Caption)
}

try {
    $file = (Resolve-Path -LiteralPath $Path).ProviderPath
    Open-InExcel $file (Get-CsvSettings $file)
} catch {
    Add-Type -AssemblyName PresentationFramework
    [void][Windows.MessageBox]::Show("Impossible d'ouvrir le CSV :`n$Path`n`n$($_.Exception.Message)", 'CSV vers Excel', 'OK', 'Error')
    exit 1
}
