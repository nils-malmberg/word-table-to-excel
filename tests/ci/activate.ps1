# Crée l'objet COM du complément comme le fait Word, puis demande la définition du ruban.
$ErrorActionPreference = 'Stop'
$addin = New-Object -ComObject WordTableToExcel.Connect
$ribbon = $addin.GetCustomUI('Microsoft.Word.Document')
if ($ribbon -notmatch 'OnExportClick') { Write-Output "Ruban invalide : $ribbon"; exit 2 }
[xml]$ribbon | Out-Null
Write-Output ("OK ({0} bits)" -f ([IntPtr]::Size * 8))
exit 0
