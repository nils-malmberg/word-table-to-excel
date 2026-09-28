<#
.SYNOPSIS
    Installe le complément Word « Tableaux Word vers Excel » pour l'utilisateur courant.

.DESCRIPTION
    - copie WordTableToExcel.dll dans %LOCALAPPDATA%\WordTableToExcel ;
    - enregistre la classe COM dans HKCU (vues 32 et 64 bits : fonctionne avec Office 32 ou 64 bits) ;
    - déclare le complément auprès de Word (chargement au démarrage).
    Aucun droit administrateur n'est nécessaire. Compatible PowerShell 2.0 et suivants.

.PARAMETER SourceDll
    Chemin de WordTableToExcel.dll (par défaut : à côté de ce script).

.PARAMETER InstallDir
    Dossier d'installation (par défaut : %LOCALAPPDATA%\WordTableToExcel).
#>
[CmdletBinding()]
param(
    [string]$SourceDll = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'WordTableToExcel.dll'),
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'WordTableToExcel')
)

$ErrorActionPreference = 'Stop'

# Doivent correspondre à AddIn\Connect.cs et au .csproj.
$ClassId      = '{03F63233-F2FE-4A75-AF7A-F99CBEDC8030}'
$ProgId       = 'WordTableToExcel.Connect'
$ClassName    = 'WordTableToExcel.AddIn.Connect'
$Version      = '1.0.0.0'
$AssemblyName = "WordTableToExcel, Version=$Version, Culture=neutral, PublicKeyToken=null"
$FriendlyName = 'Tableaux Word vers Excel'
$Description  = 'Exporte les tableaux du document (avec ou sans légende) vers un classeur Excel, en conservant la mise en forme.'

function Invoke-Reg {
    param([string[]]$Arguments)
    $output = & reg.exe @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) { throw "reg.exe $($Arguments -join ' ') : $output" }
}

function Set-RegValue {
    param([string]$Key, [string]$Name, [string]$Data, [string]$Type = 'REG_SZ', [string]$View)
    $arguments = @('add', $Key)
    if ($Name) { $arguments += @('/v', $Name) } else { $arguments += '/ve' }
    $arguments += @('/t', $Type, '/d', $Data, '/f')
    if ($View) { $arguments += $View }
    Invoke-Reg $arguments
}

function Get-LongPath {
    param([string]$Path)
    try {
        if (-not ('WordTableToExcel.Setup.NativePath' -as [type])) {
            Add-Type -Namespace WordTableToExcel.Setup -Name NativePath -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
public static extern int GetLongPathName(string shortPath, System.Text.StringBuilder longPath, int bufferSize);
'@
        }
        $buffer = New-Object System.Text.StringBuilder 1024
        if ([WordTableToExcel.Setup.NativePath]::GetLongPathName($Path, $buffer, $buffer.Capacity) -gt 0) { return $buffer.ToString() }
    } catch {
        Write-Verbose "GetLongPathName indisponible : $_"
    }
    return $Path
}

Write-Host "Installation de « $FriendlyName »..." -ForegroundColor Cyan

if (-not (Test-Path -LiteralPath $SourceDll)) {
    throw "Fichier introuvable : $SourceDll`nCompilez d'abord le projet (build.cmd) ou placez ce script à côté de WordTableToExcel.dll."
}

# .NET Framework 4.x requis (présent d'office sur Windows 8 et suivants).
$net4 = $false
foreach ($key in 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Client') {
    if (Test-Path $key) { $net4 = $true }
}
if (-not $net4) {
    Write-Warning ".NET Framework 4 n'a pas été détecté. Installez .NET Framework 4.8 (gratuit, Microsoft) pour que le complément se charge."
}

if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    Write-Warning 'Word est ouvert : le complément sera disponible au prochain démarrage de Word.'
}

# 1. Copie des fichiers.
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
    cmd.exe /c "type nul > `"$dll`":Zone.Identifier" 2>$null | Out-Null
}

# 2. Enregistrement COM par utilisateur (équivalent de « regasm /codebase », sans droits administrateur).
# Le CodeBase doit être un chemin long : .NET ne retrouve pas la DLL via un nom court 8.3 (ex. « JEAN-P~1 »).
$dll = Get-LongPath $dll
$codeBase = 'file:///' + ($dll -replace '\\', '/')
$is64 = ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') -or ($env:PROCESSOR_ARCHITEW6432 -eq 'AMD64') -or ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') -or ($env:PROCESSOR_ARCHITEW6432 -eq 'ARM64')
$views = @('/reg:32')
if ($is64) { $views += '/reg:64' }

foreach ($view in $views) {
    $classes = 'HKCU\Software\Classes'
    Set-RegValue "$classes\$ProgId" '' $ProgId -View $view
    Set-RegValue "$classes\$ProgId\CLSID" '' $ClassId -View $view

    $clsid = "$classes\CLSID\$ClassId"
    Set-RegValue $clsid '' $ProgId -View $view
    Set-RegValue "$clsid\ProgId" '' $ProgId -View $view
    Invoke-Reg @('add', "$clsid\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}", '/f', $view)

    foreach ($inproc in "$clsid\InprocServer32", "$clsid\InprocServer32\$Version") {
        if ($inproc -eq "$clsid\InprocServer32") {
            Set-RegValue $inproc '' 'mscoree.dll' -View $view
            Set-RegValue $inproc 'ThreadingModel' 'Both' -View $view
        }
        Set-RegValue $inproc 'Class' $ClassName -View $view
        Set-RegValue $inproc 'Assembly' $AssemblyName -View $view
        Set-RegValue $inproc 'RuntimeVersion' 'v4.0.30319' -View $view
        Set-RegValue $inproc 'CodeBase' $codeBase -View $view
    }
}

# 3. Déclaration du complément auprès de Word (toutes versions, 32/64 bits, Click-to-Run compris).
$addin = "HKCU\Software\Microsoft\Office\Word\Addins\$ProgId"
Set-RegValue $addin 'FriendlyName' $FriendlyName
Set-RegValue $addin 'Description' $Description
Set-RegValue $addin 'LoadBehavior' '3' 'REG_DWORD'

# Office 2013+ : ne pas désactiver le complément pour « lenteur » au démarrage.
foreach ($officeVersion in '15.0', '16.0') {
    if (Test-Path "HKCU:\Software\Microsoft\Office\$officeVersion\Word") {
        Set-RegValue "HKCU\Software\Microsoft\Office\$officeVersion\Word\Resiliency\DoNotDisableAddinList" $ProgId '1' 'REG_DWORD'
    }
}

Write-Host ''
Write-Host "Installation terminée : $dll" -ForegroundColor Green
Write-Host 'Démarrez Word : le bouton « Tableaux vers Excel » se trouve dans les onglets Accueil et Références'
Write-Host '(Word 2000-2003 : barre d''outils Standard).'
