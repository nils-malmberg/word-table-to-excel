<#
.SYNOPSIS
    Installe le complément Word « Tableaux Word vers Excel » pour tous les utilisateurs du poste.

.DESCRIPTION
    - copie WordTableToExcel.dll dans %ProgramFiles%\WordTableToExcel ;
    - l'enregistre avec RegAsm 64 bits et 32 bits : fonctionne avec Office 64 bits comme 32 bits,
      y compris les installations « Démarrer en un clic » (Microsoft 365) ;
    - déclare le complément auprès de Word (chargement au démarrage).
    Nécessite les droits administrateur : le script demande l'élévation (UAC) si besoin.
    Compatible Windows PowerShell 2.0 et suivants.

.PARAMETER SourceDll
    Chemin de WordTableToExcel.dll (par défaut : à côté de ce script).

.PARAMETER InstallDir
    Dossier d'installation (par défaut : %ProgramFiles%\WordTableToExcel).
#>
[CmdletBinding()]
param(
    [string]$SourceDll = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'WordTableToExcel.dll'),
    [string]$InstallDir = (Join-Path $env:ProgramFiles 'WordTableToExcel'),
    [switch]$PauseAtEnd
)

$ErrorActionPreference = 'Stop'
$FriendlyName = 'Tableaux Word vers Excel'

function Test-Administrator {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-RegAsmPaths {
    $paths = @()
    foreach ($framework in 'Framework64', 'Framework') {
        $regasm = Join-Path $env:WINDIR "Microsoft.NET\$framework\v4.0.30319\RegAsm.exe"
        if (Test-Path -LiteralPath $regasm) { $paths += $regasm }
    }
    return $paths
}

# Relance élevée si nécessaire (une fenêtre de confirmation Windows s'affiche).
if (-not (Test-Administrator)) {
    Write-Host 'Droits administrateur nécessaires : confirmation demandée...' -ForegroundColor Yellow
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $MyInvocation.MyCommand.Path + '"'),
                   '-SourceDll', ('"' + $SourceDll + '"'), '-InstallDir', ('"' + $InstallDir + '"'), '-PauseAtEnd')
    try {
        $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -Verb RunAs -Wait -PassThru
        exit $process.ExitCode
    } catch {
        Write-Error "Installation annulée : les droits administrateur ont été refusés. ($_)"
        exit 1
    }
}

$exitCode = 0
try {
    Write-Host "Installation de « $FriendlyName »..." -ForegroundColor Cyan

    if (-not (Test-Path -LiteralPath $SourceDll)) {
        throw "Fichier introuvable : $SourceDll`nCompilez d'abord le projet (build.cmd) ou placez ce script à côté de WordTableToExcel.dll."
    }

    $regasms = @(Get-RegAsmPaths)
    if ($regasms.Count -eq 0) {
        throw ".NET Framework 4 est introuvable. Installez .NET Framework 4.8 (gratuit, Microsoft) puis relancez l'installation."
    }

    if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
        Write-Warning 'Word est ouvert : le complément sera disponible au prochain démarrage de Word.'
    }

    # 1. Copie de la DLL.
    New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
    $dll = Join-Path $InstallDir 'WordTableToExcel.dll'
    $sourceFull = (Resolve-Path -LiteralPath $SourceDll).Path
    if ($sourceFull -ne $dll) {
        try {
            Copy-Item -LiteralPath $sourceFull -Destination $dll -Force
        } catch {
            throw "Impossible de remplacer $dll (Word est probablement ouvert). Fermez Word et relancez l'installation."
        }
    }

    # Fichier téléchargé : on retire la marque « provient d'Internet », sinon .NET refuse de le charger.
    try {
        Unblock-File -LiteralPath $dll -ErrorAction Stop
    } catch {
        try { cmd.exe /c "type nul > `"$dll`":Zone.Identifier" | Out-Null } catch { }
    }

    # 2. Enregistrement COM (64 bits puis 32 bits). RegAsm appelle aussi la fonction d'enregistrement du
    #    complément, qui le déclare auprès de Word dans la vue correspondante du registre.
    foreach ($regasm in $regasms) {
        Write-Host "Enregistrement : $regasm"
        $ErrorActionPreference = 'Continue'   # la sortie d'erreur de RegAsm est lue comme du texte
        $output = & $regasm /codebase /nologo $dll 2>&1 | Out-String
        $code = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        if ($code -ne 0) { throw "RegAsm a échoué ($regasm, code $code) :`n$output" }
    }

    Write-Host ''
    Write-Host "Installation terminée : $dll" -ForegroundColor Green
    Write-Host 'Démarrez Word : le bouton « Tableaux vers Excel » se trouve dans les onglets Accueil et Références'
    Write-Host '(Word 2000-2003 : barre d''outils Standard).'
} catch {
    Write-Host ''
    Write-Host "ÉCHEC : $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 1
}

if ($PauseAtEnd) {
    Write-Host ''
    Read-Host 'Appuyez sur Entrée pour fermer cette fenêtre' | Out-Null
}
exit $exitCode
