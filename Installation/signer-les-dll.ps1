<#
    Signe les DLL du complément « Tableaux Word vers Excel » (signature Authenticode SHA-256 horodatée),
    pour les postes où Word exige des compléments signés par un éditeur approuvé.
    Usage (dans le dossier Installation) :
        powershell -ExecutionPolicy Bypass -File .\signer-les-dll.ps1 -Empreinte <empreinte du certificat>
    Voir INFORMATIQUE.txt.
#>
param(
    [Parameter(Mandatory = $true)][string]$Empreinte,
    [string]$Horodatage = 'http://timestamp.digicert.com',
    [string]$Dossier = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
$Empreinte = ($Empreinte -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
$certificat = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -CodeSigningCert -ErrorAction SilentlyContinue |
    Where-Object { $_.Thumbprint -eq $Empreinte -and $_.HasPrivateKey } | Select-Object -First 1
if (-not $certificat) {
    throw "Certificat de signature de code introuvable (empreinte $Empreinte, avec clé privée, magasins personnels utilisateur et ordinateur)."
}
Write-Host "Certificat : $($certificat.Subject) (expire le $($certificat.NotAfter.ToString('d')))"

foreach ($nom in 'WordTableToExcel.Shim32.dll', 'WordTableToExcel.Shim64.dll', 'WordTableToExcel.dll') {
    $fichier = Join-Path $Dossier $nom
    if (-not (Test-Path $fichier)) { throw "Fichier absent : $fichier" }
    $parametres = @{ FilePath = $fichier; Certificate = $certificat; HashAlgorithm = 'SHA256' }
    if ($Horodatage) { $parametres.TimestampServer = $Horodatage }
    $resultat = Set-AuthenticodeSignature @parametres
    if (-not $resultat.SignerCertificate) { throw "Échec de la signature de $nom : $($resultat.StatusMessage)" }
    $etat = (Get-AuthenticodeSignature $fichier).Status
    Write-Host ("{0,-30} signé ({1})" -f $nom, $etat)
    if ($etat -ne 'Valid') {
        Write-Warning "$nom : la chaîne du certificat n'est pas approuvée sur ce poste ($etat). Sur les postes des utilisateurs, le certificat doit être reconnu comme éditeur approuvé."
    }
}
Write-Host 'Terminé : diffusez ce dossier, les utilisateurs lancent install.cmd.'
