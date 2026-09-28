<#
.SYNOPSIS
    Désinstalle le complément Word « Tableaux Word vers Excel ».
    Nécessite les droits administrateur : le script demande l'élévation (UAC) si besoin.
#>
[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:ProgramFiles 'WordTableToExcel'),
    [switch]$KeepSettings,
    [switch]$PauseAtEnd
)

$ErrorActionPreference = 'Continue'
$ClassId = '{03F63233-F2FE-4A75-AF7A-F99CBEDC8030}'
$ProgId  = 'WordTableToExcel.Connect'

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host 'Droits administrateur nécessaires : confirmation demandée...' -ForegroundColor Yellow
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $MyInvocation.MyCommand.Path + '"'),
                   '-InstallDir', ('"' + $InstallDir + '"'), '-PauseAtEnd')
    if ($KeepSettings) { $arguments += '-KeepSettings' }
    try {
        $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -Verb RunAs -Wait -PassThru
        exit $process.ExitCode
    } catch {
        Write-Error "Désinstallation annulée : les droits administrateur ont été refusés. ($_)"
        exit 1
    }
}

Write-Host 'Désinstallation de « Tableaux Word vers Excel »...' -ForegroundColor Cyan

if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    Write-Warning 'Word est ouvert : fermez-le pour que les fichiers puissent être supprimés.'
}

# 1. Désenregistrement COM (retire aussi la déclaration auprès de Word).
$dll = Join-Path $InstallDir 'WordTableToExcel.dll'
foreach ($framework in 'Framework64', 'Framework') {
    $regasm = Join-Path $env:WINDIR "Microsoft.NET\$framework\v4.0.30319\RegAsm.exe"
    if ((Test-Path -LiteralPath $regasm) -and (Test-Path -LiteralPath $dll)) {
        & $regasm /unregister /nologo $dll 2>&1 | Out-Null
    }
}

# 2. Nettoyage de sécurité des clés (au cas où la DLL aurait déjà été supprimée).
$is64 = ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') -or ($env:PROCESSOR_ARCHITEW6432 -eq 'AMD64') -or ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') -or ($env:PROCESSOR_ARCHITEW6432 -eq 'ARM64')
$views = @('/reg:32')
if ($is64) { $views += '/reg:64' }
foreach ($view in $views) {
    foreach ($key in "HKLM\Software\Classes\CLSID\$ClassId", "HKLM\Software\Classes\$ProgId",
                     "HKLM\Software\Microsoft\Office\Word\Addins\$ProgId",
                     "HKCU\Software\Classes\CLSID\$ClassId", "HKCU\Software\Classes\$ProgId",
                     "HKCU\Software\Microsoft\Office\Word\Addins\$ProgId") {
        & reg.exe delete $key /f $view 2>&1 | Out-Null
    }
}
if (-not $KeepSettings) { & reg.exe delete 'HKCU\Software\WordTableToExcel' /f 2>&1 | Out-Null }

# 3. Fichiers.
$exitCode = 0
if (Test-Path -LiteralPath $InstallDir) {
    try {
        Remove-Item -LiteralPath $InstallDir -Recurse -Force -ErrorAction Stop
    } catch {
        Write-Warning "Impossible de supprimer $InstallDir (Word est-il ouvert ?). Supprimez-le manuellement après avoir fermé Word."
        $exitCode = 1
    }
}

Write-Host 'Désinstallation terminée.' -ForegroundColor Green
if ($PauseAtEnd) {
    Write-Host ''
    Read-Host 'Appuyez sur Entrée pour fermer cette fenêtre' | Out-Null
}
exit $exitCode
