<#
    Test de fumée exécuté par l'intégration continue (Windows, sans Word) :

    1. install.ps1 : copie dans Program Files et enregistrement COM 64 et 32 bits (RegAsm) ;
    2. vérification des clés (classe COM et déclaration du complément Word, dans les deux vues du registre) ;
    3. activation du complément et lecture du ruban par un client COM natif (cscript, comme Word),
       en 64 et 32 bits, depuis un processus élevé et depuis un processus utilisateur standard
       (« runas /trustlevel ») ;
    4. uninstall.ps1 et vérification du nettoyage.
#>
param([Parameter(Mandatory = $true)][string]$Dist)

$ErrorActionPreference = 'Stop'
$dist = (Resolve-Path $Dist).Path
$clsid = '{03F63233-F2FE-4A75-AF7A-F99CBEDC8030}'
$activate = Join-Path $PSScriptRoot 'activate.vbs'
$installDir = Join-Path $env:ProgramFiles 'WordTableToExcel'
$hosts = @("$env:WINDIR\System32\cscript.exe", "$env:WINDIR\SysWOW64\cscript.exe")

function Test-Key([string]$key, [string]$view) {
    $ErrorActionPreference = 'Continue'   # « clé introuvable » est une réponse attendue, pas une erreur
    & reg.exe query $key $view 2>&1 | Out-Null
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

# --- 1. Installation.
& (Join-Path $dist 'install.ps1') -SourceDll (Join-Path $dist 'WordTableToExcel.dll')
if ($LASTEXITCODE -ne 0) { throw "install.ps1 a échoué (code $LASTEXITCODE)." }

# --- 2. Registre.
foreach ($view in '/reg:64', '/reg:32') {
    foreach ($key in "HKLM\Software\Classes\CLSID\$clsid\InprocServer32", 'HKLM\Software\Microsoft\Office\Word\Addins\WordTableToExcel.Connect') {
        if (-not (Test-Key $key $view)) { throw "Clé absente ($view) : $key" }
    }
}
$loadBehavior = (& reg.exe query 'HKLM\Software\Microsoft\Office\Word\Addins\WordTableToExcel.Connect' /v LoadBehavior /reg:64 | Out-String)
if ($loadBehavior -notmatch '0x3') { throw "LoadBehavior inattendu : $loadBehavior" }

# --- 3. Activation COM.
$failures = @()
foreach ($cscript in $hosts) {
    if (-not (Test-Path $cscript)) { continue }
    foreach ($standardUser in $false, $true) {
        $result = Invoke-Activation $cscript $standardUser
        $label = if ($standardUser) { 'utilisateur standard' } else { 'administrateur' }
        Write-Host "$cscript ($label) -> $result"
        if ($result -notlike 'OK*') { $failures += "$cscript ($label) : $result" }
    }
}
if ($failures.Count -gt 0) { throw ("Activation COM en échec :`n" + ($failures -join "`n")) }

# --- 4. Désinstallation.
& (Join-Path $dist 'uninstall.ps1')
foreach ($view in '/reg:64', '/reg:32') {
    foreach ($key in "HKLM\Software\Classes\CLSID\$clsid", 'HKLM\Software\Microsoft\Office\Word\Addins\WordTableToExcel.Connect') {
        if (Test-Key $key $view) { throw "Désinstallation incomplète ($view) : $key" }
    }
}
if (Test-Path $installDir) { throw "Désinstallation incomplète : $installDir" }

Write-Host 'Test de fumée réussi.'
exit 0   # sinon le code du dernier « reg query » (clé absente, attendu) serait renvoyé
