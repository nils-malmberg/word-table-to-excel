<#
    Test de fumée exécuté par l'intégration continue (Windows, sans Word).

    1. install.ps1 (enregistrement par utilisateur, HKCU) puis activation COM du complément et lecture
       du ruban par un client COM natif (cscript, comme Word), 64 et 32 bits, depuis un processus
       NON élevé lancé via « runas /trustlevel » (l'agent CI est administrateur).
    2. uninstall.ps1 et vérification du nettoyage.
    3. Enregistrement machine (regasm /codebase, HKLM) et activation depuis un processus élevé, 64 et 32 bits.
#>
param([Parameter(Mandatory = $true)][string]$Dist)

$ErrorActionPreference = 'Stop'
$dist = (Resolve-Path $Dist).Path
$activate = Join-Path $PSScriptRoot 'activate.vbs'
$hosts = @("$env:WINDIR\System32\cscript.exe", "$env:WINDIR\SysWOW64\cscript.exe")

function Invoke-Activation([string]$cscript, [bool]$restricted) {
    $out = Join-Path $env:TEMP ('activation-' + [Guid]::NewGuid().ToString('N') + '.txt')
    if ($restricted) {
        & runas.exe /trustlevel:0x20000 "$cscript //nologo $activate $out" | Out-Null
        for ($i = 0; $i -lt 120 -and -not (Test-Path $out); $i++) { Start-Sleep -Milliseconds 500 }
        if (-not (Test-Path $out)) { return $null }
        Start-Sleep -Milliseconds 300
    } else {
        & $cscript //nologo $activate $out | Out-Host
    }
    $result = [IO.File]::ReadAllText($out)
    Remove-Item $out -ErrorAction SilentlyContinue
    return $result
}

$failures = @()

# --- 1. Installation par utilisateur (HKCU), activation non élevée.
$installDir = Join-Path $env:LOCALAPPDATA 'WordTableToExcel'
& (Join-Path $dist 'install.ps1') -SourceDll (Join-Path $dist 'WordTableToExcel.dll') -InstallDir $installDir

$addin = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\Word\Addins\WordTableToExcel.Connect'
if ($addin.LoadBehavior -ne 3) { throw "LoadBehavior inattendu : $($addin.LoadBehavior)" }
$inproc = Get-ItemProperty 'HKCU:\Software\Classes\CLSID\{03F63233-F2FE-4A75-AF7A-F99CBEDC8030}\InprocServer32'
if ($inproc.CodeBase -notlike 'file:///*WordTableToExcel.dll' -or $inproc.CodeBase -like '*~*') { throw "CodeBase inattendu : $($inproc.CodeBase)" }

foreach ($ps in $hosts) {
    if (-not (Test-Path $ps)) { continue }
    $result = Invoke-Activation $ps $true
    if ($result -eq $null) {
        Write-Warning "Impossible de lancer un processus non élevé (runas) sur cet agent : activation HKCU non vérifiée pour $ps."
        continue
    }
    Write-Host "HKCU, $ps -> $result"
    if ($result -notlike 'OK*') { $failures += "HKCU $ps : $result" }
}

# --- 2. Désinstallation.
& (Join-Path $dist 'uninstall.ps1') -InstallDir $installDir
if (Test-Path 'HKCU:\Software\Microsoft\Office\Word\Addins\WordTableToExcel.Connect') { throw 'Désinstallation incomplète (clé Addins).' }
if (Test-Path 'HKCU:\Software\Classes\WordTableToExcel.Connect') { throw 'Désinstallation incomplète (ProgID).' }
if (Test-Path $installDir) { throw "Désinstallation incomplète (dossier $installDir)." }

# Un chemin d'installation donné sous forme courte (8.3) est enregistré sous sa forme longue.
$shortDir = Join-Path $env:TEMP 'WordTableToExcel-ci'
& (Join-Path $dist 'install.ps1') -SourceDll (Join-Path $dist 'WordTableToExcel.dll') -InstallDir $shortDir | Out-Null
$codeBase = (Get-ItemProperty 'HKCU:\Software\Classes\CLSID\{03F63233-F2FE-4A75-AF7A-F99CBEDC8030}\InprocServer32').CodeBase
Write-Host "Installation depuis $shortDir -> CodeBase $codeBase"
if ($codeBase -like '*~*') { throw "Le CodeBase contient encore un nom court : $codeBase" }
& (Join-Path $dist 'uninstall.ps1') -InstallDir $shortDir | Out-Null

# --- 3. Enregistrement machine (regasm) et activation élevée.
$dll = Join-Path $env:TEMP 'wtte-regasm\WordTableToExcel.dll'
New-Item -ItemType Directory -Force (Split-Path $dll) | Out-Null
Copy-Item (Join-Path $dist 'WordTableToExcel.dll') $dll -Force
$regasms = @{
    "$env:WINDIR\System32\cscript.exe" = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe";
    "$env:WINDIR\SysWOW64\cscript.exe" = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
}
foreach ($ps in $hosts) {
    if (-not (Test-Path $ps)) { continue }
    $regasm = $regasms[$ps]
    & $regasm /codebase /nologo $dll | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "regasm en échec ($regasm)" }
    try {
        $result = Invoke-Activation $ps $false
        Write-Host "HKLM, $ps -> $result"
        if ($result -notlike 'OK*') { $failures += "HKLM $ps : $result" }
    } finally {
        & $regasm /unregister /nologo $dll | Out-Host
    }
}
Remove-Item 'HKCU:\Software\Microsoft\Office\Word\Addins\WordTableToExcel.Connect' -ErrorAction SilentlyContinue

if ($failures.Count -gt 0) { throw ("Activation COM en échec :`n" + ($failures -join "`n")) }
Write-Host 'Test de fumée réussi.'
