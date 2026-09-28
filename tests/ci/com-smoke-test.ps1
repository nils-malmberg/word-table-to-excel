<#
    Test de fumée exécuté par l'intégration continue (Windows, sans Word), sur un dossier
    « Installation » tel que l'utilisateur l'obtient en extrayant le .zip de GitHub :

    1. les fichiers sont marqués « provenant d'Internet » (Zone.Identifier), comme après un téléchargement ;
    2. install.cmd (utilisateur courant, sans droits administrateur) ;
    3. vérification des clés HKCU (classe COM dans les vues 32 et 64 bits, déclaration auprès de Word) ;
    4. activation du complément et lecture du ruban par un client COM natif (cscript, comme Word),
       en 64 et 32 bits, depuis un processus utilisateur standard (« runas /trustlevel ») ;
    5. uninstall.cmd et vérification du nettoyage.
#>
param([Parameter(Mandatory = $true)][string]$Folder)

$ErrorActionPreference = 'Stop'
$source = (Resolve-Path $Folder).Path
$clsid = '{03F63233-F2FE-4A75-AF7A-F99CBEDC8030}'
$activate = Join-Path $PSScriptRoot 'activate.vbs'
$installDir = Join-Path $env:LOCALAPPDATA 'WordTableToExcel'
$env:WTTE_NO_PAUSE = '1'

function Test-Key([string]$key, [string]$view) {
    $ErrorActionPreference = 'Continue'   # « clé introuvable » est une réponse attendue, pas une erreur
    if ($view) { & reg.exe query $key $view 2>&1 | Out-Null } else { & reg.exe query $key 2>&1 | Out-Null }
    return $LASTEXITCODE -eq 0
}

function Invoke-Activation([string]$cscript, [bool]$standardUser) {
    $out = Join-Path $env:TEMP ('activation-' + [Guid]::NewGuid().ToString('N') + '.txt')
    if ($standardUser) {
        & runas.exe /trustlevel:0x20000 "$cscript //nologo $activate $out" | Out-Null
        for ($i = 0; $i -lt 120 -and -not (Test-Path $out); $i++) { Start-Sleep -Milliseconds 500 }
        if (-not (Test-Path $out)) { return '(aucun résultat : runas indisponible)' }
        Start-Sleep -Milliseconds 300
    } else {
        & $cscript //nologo $activate $out | Out-Host
    }
    $result = [IO.File]::ReadAllText($out)
    Remove-Item $out -ErrorAction SilentlyContinue
    return $result
}

# --- 1. Copie du dossier, fichiers marqués « téléchargés ».
$folder = Join-Path $env:TEMP ('Installation-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
Copy-Item $source $folder -Recurse
Get-ChildItem $folder -File | ForEach-Object {
    Set-Content -LiteralPath $_.FullName -Stream Zone.Identifier -Value "[ZoneTransfer]`r`nZoneId=3"
}

# --- 2. Installation.
& cmd.exe /c "`"$folder\install.cmd`""
if ($LASTEXITCODE -ne 0) { throw "install.cmd a échoué (code $LASTEXITCODE)." }

# --- 3. Registre.
foreach ($view in '/reg:64', '/reg:32') {
    if (-not (Test-Key "HKCU\Software\Classes\CLSID\$clsid\InprocServer32" $view)) { throw "Classe COM absente ($view)." }
}
if (-not (Test-Key 'HKCU\Software\Microsoft\Office\Word\Addins\WordTableToExcel.Connect' '')) { throw 'Déclaration Word absente.' }
$zone = Get-Content -LiteralPath (Join-Path $installDir 'WordTableToExcel.dll') -Stream Zone.Identifier -ErrorAction SilentlyContinue
if ($zone -match 'ZoneId=3') { throw 'La marque « provenant d''Internet » n''a pas été retirée.' }

# --- 4. Activation COM (client natif, utilisateur standard ; administrateur pour information).
$failures = @()
foreach ($cscript in "$env:WINDIR\System32\cscript.exe", "$env:WINDIR\SysWOW64\cscript.exe") {
    if (-not (Test-Path $cscript)) { continue }
    $result = Invoke-Activation $cscript $true
    Write-Host "$cscript (utilisateur standard) -> $result"
    if ($result -notlike 'OK*') { $failures += "$cscript : $result" }
    $elevated = Invoke-Activation $cscript $false
    Write-Host "$cscript (administrateur, pour information) -> $elevated"
}
if ($failures.Count -gt 0) {
    $log = Join-Path $installDir 'WordTableToExcel.log'
    if (Test-Path $log) { Get-Content $log | Write-Host }
    throw ("Activation COM en échec :`n" + ($failures -join "`n"))
}

# --- 5. Désinstallation.
& cmd.exe /c "`"$folder\uninstall.cmd`""
if ($LASTEXITCODE -ne 0) { throw "uninstall.cmd a échoué (code $LASTEXITCODE)." }
foreach ($view in '/reg:64', '/reg:32') {
    if (Test-Key "HKCU\Software\Classes\CLSID\$clsid" $view) { throw "Désinstallation incomplète ($view)." }
}
if (Test-Key 'HKCU\Software\Microsoft\Office\Word\Addins\WordTableToExcel.Connect' '') { throw 'Déclaration Word non supprimée.' }
if (Test-Path $installDir) { throw "Désinstallation incomplète : $installDir" }
Remove-Item $folder -Recurse -Force

Write-Host "Test de fumée réussi pour $source."
exit 0
