<#
    Test de fumée exécuté par l'intégration continue (Windows, sans Word) :
    installation par install.ps1, activation COM du complément depuis PowerShell 64 bits et 32 bits,
    lecture du ruban (GetCustomUI), puis désinstallation. Deux emplacements sont testés :
    le dossier par défaut (%LOCALAPPDATA%) et un chemin donné sous forme courte 8.3 (%TEMP% de l'agent).
#>
param([Parameter(Mandatory = $true)][string]$Dist)

$ErrorActionPreference = 'Stop'
$dist = (Resolve-Path $Dist).Path

# Journal de chargement .NET (Fusion) pour diagnostiquer un échec ; nécessite les droits administrateur de l'agent.
$fusionLogs = Join-Path $env:TEMP 'fusion-logs'
try {
    New-Item -ItemType Directory -Force $fusionLogs | Out-Null
    $fusion = 'HKLM:\SOFTWARE\Microsoft\Fusion'
    if (-not (Test-Path $fusion)) { New-Item $fusion | Out-Null }
    Set-ItemProperty $fusion -Name LogFailures -Value 1 -Type DWord
    Set-ItemProperty $fusion -Name LogPath -Value ($fusionLogs + '\') -Type String
} catch {
    Write-Host "Journal Fusion non activé : $_"
}

function Test-Location([string]$installDir) {
    Write-Host "=== Installation dans $installDir ==="

    & (Join-Path $dist 'install.ps1') -SourceDll (Join-Path $dist 'WordTableToExcel.dll') -InstallDir $installDir

    # Contrôle préalable : l'assembly et la classe se chargent (lecture en mémoire, le fichier n'est pas verrouillé).
    $assembly = [Reflection.Assembly]::Load([IO.File]::ReadAllBytes((Join-Path $installDir 'WordTableToExcel.dll')))
    $type = $assembly.GetType('WordTableToExcel.AddIn.Connect', $true)
    Write-Host ("Assembly : {0} ; GUID de la classe : {1}" -f $assembly.FullName, $type.GUID)

    $addin = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\Word\Addins\WordTableToExcel.Connect'
    if ($addin.LoadBehavior -ne 3) { throw "LoadBehavior inattendu : $($addin.LoadBehavior)" }

    $activate = Join-Path $PSScriptRoot 'activate.ps1'
    $hosts = @("$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe", "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe")
    foreach ($ps in $hosts) {
        if (-not (Test-Path $ps)) { continue }
        # Les erreurs du processus enfant sont lues comme du texte, sans interrompre ce script.
        $ErrorActionPreference = 'Continue'
        $output = & $ps -NoProfile -ExecutionPolicy Bypass -File $activate 2>&1 | Out-String
        $code = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        Write-Host "$ps ->"
        Write-Host $output
        if ($code -ne 0) { throw "Activation COM en échec avec $ps (code $code)." }
    }

    & (Join-Path $dist 'uninstall.ps1') -InstallDir $installDir
    if (Test-Path 'HKCU:\Software\Microsoft\Office\Word\Addins\WordTableToExcel.Connect') { throw 'Désinstallation incomplète (clé Addins).' }
    if (Test-Path 'HKCU:\Software\Classes\WordTableToExcel.Connect') { throw 'Désinstallation incomplète (ProgID).' }
    if (Test-Path $installDir) { throw "Désinstallation incomplète (dossier $installDir)." }
}

try {
    Test-Location (Join-Path $env:LOCALAPPDATA 'WordTableToExcel')
    Test-Location (Join-Path $env:TEMP 'WordTableToExcel-ci')
} catch {
    Get-ChildItem $fusionLogs -Recurse -Filter *.htm -ErrorAction SilentlyContinue | Select-Object -First 3 | ForEach-Object {
        Write-Host "--- $($_.FullName)"
        (Get-Content $_.FullName -Raw) -replace '<[^>]+>', '' | Write-Host
    }
    throw
}
Write-Host 'Test de fumée réussi.'
