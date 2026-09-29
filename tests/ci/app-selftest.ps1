<#
    Vérification de l'application autonome TableauxWordExcel.exe sous Windows (sans Word) :
    - exécutable unique, AnyCPU, manifeste (sans élévation, mise à l'échelle GDI), icône ;
    - autotest intégré (--selftest) : .NET Framework, filtre de messages OLE, recherche de Word,
      construction des fenêtres (captures PNG), lecture du classeur de test ;
    - démarrage réel : la fenêtre s'ouvre sans Word, une seconde instance se contente de réactiver
      la première, fermeture normale.
#>
param(
    [Parameter(Mandatory = $true)][string]$Exe,
    [Parameter(Mandatory = $true)][string]$Fixture,
    [Parameter(Mandatory = $true)][string]$Output,
    # Affiche les captures dans le journal (base64), pour les examiner sans télécharger d'artefact.
    [switch]$PrintScreenshots
)

$ErrorActionPreference = 'Stop'
$exePath = (Resolve-Path $Exe).Path
$fixturePath = (Resolve-Path $Fixture).Path
New-Item -ItemType Directory -Force $Output | Out-Null
$outPath = (Resolve-Path $Output).Path

# --- Fichier unique : ni .config ni DLL nécessaires à côté.
$assembly = [Reflection.AssemblyName]::GetAssemblyName($exePath)
Write-Host "Assemblage : $($assembly.FullName) ($($assembly.ProcessorArchitecture))"
if ($assembly.ProcessorArchitecture -ne 'MSIL') { throw "L'exécutable doit être AnyCPU." }
$info = (Get-Item $exePath).VersionInfo
Write-Host "Produit : $($info.ProductName) $($info.FileVersion) ; taille : $([math]::Round((Get-Item $exePath).Length / 1KB)) Ko"

$raw = [Text.Encoding]::UTF8.GetString([IO.File]::ReadAllBytes($exePath))
if ($raw -notmatch 'level="asInvoker"') { throw 'Manifeste : exécution sans élévation (asInvoker) absente.' }
if ($raw -notmatch 'gdiScaling') { throw 'Manifeste : mise à l''échelle GDI absente.' }
Add-Type -AssemblyName System.Drawing
$icon = [Drawing.Icon]::ExtractAssociatedIcon($exePath)
if ($icon -eq $null) { throw 'Icône absente.' }
Write-Host 'OK  manifeste et icône'

# --- Autotest intégré.
$selftest = Start-Process -FilePath $exePath -ArgumentList @('--selftest', "`"$outPath`"", "`"$fixturePath`"") -Wait -PassThru
Get-Content (Join-Path $outPath 'selftest.txt') -Encoding UTF8 | ForEach-Object { Write-Host $_ }
if ($selftest.ExitCode -ne 0) { throw "Autotest en échec (code $($selftest.ExitCode))." }
foreach ($png in 'fenetre.png', 'fenetre-sans-word.png', 'export.png', 'import.png', 'import-tout.png') {
    $file = Join-Path $outPath $png
    if (-not (Test-Path $file)) { throw "Capture manquante : $png" }
    if ($PrintScreenshots) {
        $b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($file))
        for ($i = 0; $i -lt $b64.Length; $i += 3000) {
            Write-Host ("PNG64 {0} {1}" -f $png, $b64.Substring($i, [math]::Min(3000, $b64.Length - $i)))
        }
    }
}

# --- Démarrage réel, sans Word.
$first = Start-Process -FilePath $exePath -PassThru
Start-Sleep -Seconds 5
$first.Refresh()
if ($first.HasExited) { throw "L'application s'est fermée au démarrage (code $($first.ExitCode))." }
Write-Host "OK  application démarrée : « $($first.MainWindowTitle) »"

$second = Start-Process -FilePath $exePath -PassThru
if (-not $second.WaitForExit(15000)) { $second.Kill(); throw 'La seconde instance ne s''est pas arrêtée.' }
if ($second.ExitCode -ne 0) { throw "Seconde instance : code $($second.ExitCode)." }
Write-Host 'OK  instance unique'

$closed = $false
if ($first.MainWindowHandle -ne [IntPtr]::Zero) {
    [void]$first.CloseMainWindow()
    $closed = $first.WaitForExit(15000)
}
if (-not $closed) {
    Write-Host 'Fermeture normale impossible sur ce poste (session sans bureau) : arrêt forcé.'
    $first.Kill()
} else {
    Write-Host "OK  fermeture normale (code $($first.ExitCode))"
}

$log = Join-Path $env:LOCALAPPDATA 'WordTableToExcel\WordTableToExcel.log'
if (Test-Path $log) {
    Write-Host '--- Journal'
    Get-Content $log -Encoding UTF8 -Tail 30 | ForEach-Object { Write-Host $_ }
    if (Select-String -Path $log -Pattern ' ERROR ' -SimpleMatch -Quiet) { throw 'Le journal contient des erreurs.' }
}
Write-Host 'Application vérifiée.'
exit 0
