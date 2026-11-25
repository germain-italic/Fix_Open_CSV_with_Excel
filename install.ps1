#requires -Version 5.1
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$AppId   = 'CsvExcel.Import'
$Label   = 'Excel (import CSV propre)'
$Classes = 'HKCU:\Software\Classes'
$Dest    = Join-Path $env:LOCALAPPDATA 'CsvExcel'
$Script  = Join-Path $Dest 'Open-CsvInExcel.ps1'
$Verb    = "$Classes\SystemFileAssociations\.csv\shell\CsvExcel"
$OpenWith = "$Classes\.csv\OpenWithProgids"

function Set-RegValue([string]$Key, [string]$Name, [string]$Value) {
    if (-not (Test-Path -LiteralPath $Key)) { New-Item -Path $Key -Force | Out-Null }
    Set-ItemProperty -LiteralPath $Key -Name $Name -Value $Value
}

if ($Uninstall) {
    foreach ($k in "$Classes\$AppId", $Verb) {
        if (Test-Path -LiteralPath $k) { Remove-Item -LiteralPath $k -Recurse -Force }
    }
    Remove-ItemProperty -LiteralPath $OpenWith -Name $AppId -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $Dest) { Remove-Item -LiteralPath $Dest -Recurse -Force }
    Write-Host 'Désinstallé. Choisis à nouveau Excel pour les .csv si Windows le demande.'
    return
}

$excel = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction SilentlyContinue).'(default)'
if (-not $excel) { throw 'Excel introuvable sur ce PC.' }

New-Item -ItemType Directory -Force -Path $Dest | Out-Null
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Open-CsvInExcel.ps1') -Destination $Script -Force
Unblock-File -LiteralPath $Script

$ps  = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$cmd = "`"$ps`" -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Script`" `"%1`""

Set-RegValue "$Classes\$AppId"                    '(default)'       'Fichier CSV'
Set-RegValue "$Classes\$AppId\DefaultIcon"        '(default)'       "`"$excel`",0"
Set-RegValue "$Classes\$AppId\Application"        'ApplicationName' $Label
Set-RegValue "$Classes\$AppId\shell\open"         'FriendlyAppName' $Label
Set-RegValue "$Classes\$AppId\shell\open\command" '(default)'       $cmd
Set-RegValue $OpenWith                            $AppId            ''

# Clic droit > "Ouvrir dans Excel (import propre)", quel que soit le programme par défaut
Set-RegValue $Verb          '(default)' 'Ouvrir dans Excel (import propre)'
Set-RegValue $Verb          'Icon'      "`"$excel`",0"
Set-RegValue "$Verb\command" '(default)' $cmd

# Windows interdit de forcer l'association : on ouvre le choix "Ouvrir avec"
$sample = Join-Path $Dest 'exemple.csv'
[IO.File]::WriteAllText($sample, "nom;ville;montant`r`nÉlodie;Saint-Étienne;12,50`r`n", (New-Object Text.UTF8Encoding($false)))
Write-Host "Installé dans $Dest"
Write-Host "Dans la fenêtre qui s'ouvre : choisis « $Label » puis « Toujours »."
Start-Process rundll32.exe -ArgumentList "shell32.dll,OpenAs_RunDLL `"$sample`""
