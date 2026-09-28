<#
.SYNOPSIS
    Désinstalle le complément Word « Tableaux Word vers Excel » de l'utilisateur courant.
#>
[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'WordTableToExcel'),
    [switch]$KeepSettings
)

$ErrorActionPreference = 'Continue'

$ClassId = '{03F63233-F2FE-4A75-AF7A-F99CBEDC8030}'
$ProgId  = 'WordTableToExcel.Connect'

function Remove-RegKey {
    param([string]$Key, [string]$View)
    $arguments = @('delete', $Key, '/f')
    if ($View) { $arguments += $View }
    & reg.exe @arguments 2>&1 | Out-Null
}

Write-Host 'Désinstallation de « Tableaux Word vers Excel »...' -ForegroundColor Cyan

if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    Write-Warning 'Word est ouvert : fermez-le pour que les fichiers puissent être supprimés.'
}

$is64 = ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') -or ($env:PROCESSOR_ARCHITEW6432 -eq 'AMD64') -or ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') -or ($env:PROCESSOR_ARCHITEW6432 -eq 'ARM64')
$views = @('/reg:32')
if ($is64) { $views += '/reg:64' }

foreach ($view in $views) {
    Remove-RegKey "HKCU\Software\Classes\CLSID\$ClassId" $view
    Remove-RegKey "HKCU\Software\Classes\$ProgId" $view
}
Remove-RegKey "HKCU\Software\Microsoft\Office\Word\Addins\$ProgId"
foreach ($officeVersion in '15.0', '16.0') {
    & reg.exe delete "HKCU\Software\Microsoft\Office\$officeVersion\Word\Resiliency\DoNotDisableAddinList" /v $ProgId /f 2>&1 | Out-Null
}
if (-not $KeepSettings) { Remove-RegKey 'HKCU\Software\WordTableToExcel' }

if (Test-Path -LiteralPath $InstallDir) {
    try {
        Remove-Item -LiteralPath $InstallDir -Recurse -Force -ErrorAction Stop
    } catch {
        Write-Warning "Impossible de supprimer $InstallDir (Word est-il ouvert ?). Supprimez-le manuellement après avoir fermé Word."
    }
}

Write-Host 'Désinstallation terminée.' -ForegroundColor Green
