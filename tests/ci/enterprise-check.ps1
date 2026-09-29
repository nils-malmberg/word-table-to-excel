<#
    Postes d'entreprise (intégration continue, Windows, sans Word) :
    1. signer-les-dll.ps1 signe les trois DLL (certificat de test auto-signé) et le complément signé
       s'installe et s'active toujours (test de fumée complet sur le dossier signé) ;
    2. install.cmd détecte les règles de Word qui bloqueraient un complément non signé
       (stratégie de groupe ou réglage de l'utilisateur) et l'explique.
#>
param([Parameter(Mandatory = $true)][string]$Folder)

$ErrorActionPreference = 'Stop'
$env:WTTE_NO_PAUSE = '1'
$source = (Resolve-Path $Folder).Path

# --- 1. Signature.
$signed = Join-Path $env:TEMP ('Signe-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
Copy-Item $source $signed -Recurse
$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=WordTableToExcel CI' -CertStoreLocation Cert:\CurrentUser\My
try {
    & (Join-Path $signed 'signer-les-dll.ps1') -Empreinte $cert.Thumbprint -Horodatage '' -Dossier $signed 3>$null
    foreach ($name in 'WordTableToExcel.Shim32.dll', 'WordTableToExcel.Shim64.dll', 'WordTableToExcel.dll') {
        $signature = Get-AuthenticodeSignature (Join-Path $signed $name)
        if (-not $signature.SignerCertificate -or $signature.SignerCertificate.Thumbprint -ne $cert.Thumbprint) {
            throw "$name n'est pas signé ($($signature.Status))."
        }
        Write-Host "$name signé ($($signature.Status))"
    }
    & (Join-Path $PSScriptRoot 'com-smoke-test.ps1') -Folder $signed
    if ($LASTEXITCODE -ne 0) { throw 'Le complément signé ne s''installe pas ou ne s''active pas.' }
}
finally {
    Remove-Item "Cert:\CurrentUser\My\$($cert.Thumbprint)" -ErrorAction SilentlyContinue
    Remove-Item $signed -Recurse -Force -ErrorAction SilentlyContinue
}

# --- 2. Détection des règles de Word par install.cmd.
function Invoke-Install([string]$folder) {
    $ErrorActionPreference = 'Continue'
    $output = & cmd.exe /c "`"$folder\install.cmd`"" 2>&1 | Out-String
    Write-Host $output
    if ($LASTEXITCODE -ne 0) { throw "install.cmd a échoué (code $LASTEXITCODE)." }
    & cmd.exe /c "`"$folder\uninstall.cmd`"" | Out-Null
    return $output
}

$cases = @(
    @{ Key = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Word\Security'; Value = 'requireaddinsig'; Expected = 'SIGNES par un editeur approuve' },
    @{ Key = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'; Value = 'requireaddinsig'; Expected = 'decochez' },
    @{ Key = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'; Value = 'disablealladdins'; Expected = 'decochez' }
)
$output = Invoke-Install $source
if ($output -match 'MAIS') { throw 'Avertissement affiché alors qu''aucune règle ne bloque le complément.' }
foreach ($case in $cases) {
    $created = -not (Test-Path $case.Key)
    New-Item -Path $case.Key -Force | Out-Null
    New-ItemProperty -Path $case.Key -Name $case.Value -Value 1 -PropertyType DWord -Force | Out-Null
    try {
        $output = Invoke-Install $source
        if ($output -notmatch [regex]::Escape($case.Expected)) { throw "Règle $($case.Key)\$($case.Value) non signalée par install.cmd." }
        Write-Host "OK : $($case.Key)\$($case.Value) signalée."
    }
    finally {
        Remove-ItemProperty -Path $case.Key -Name $case.Value -ErrorAction SilentlyContinue
        if ($created) { Remove-Item $case.Key -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
Write-Host 'Signature et détection des règles de Word vérifiées.'
exit 0
