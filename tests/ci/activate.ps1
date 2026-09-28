# Crée l'objet COM du complément comme le fait Word, puis demande la définition du ruban.
$ErrorActionPreference = 'Stop'
$bits = [IntPtr]::Size * 8
try {
    $addin = New-Object -ComObject WordTableToExcel.Connect
    $ribbon = $addin.GetCustomUI('Microsoft.Word.Document')
} catch {
    $e = $_.Exception
    while ($e.InnerException) { $e = $e.InnerException }
    Write-Output ("ÉCHEC ({0} bits) : {1} HRESULT=0x{2:X8} {3}" -f $bits, $e.GetType().FullName, $e.HResult, $e.Message)
    Write-Output (& reg.exe query 'HKCU\Software\Classes\CLSID\{03F63233-F2FE-4A75-AF7A-F99CBEDC8030}' /s 2>&1 | Out-String)
    exit 3
}
if ($ribbon -notmatch 'OnExportClick') { Write-Output "Ruban invalide : $ribbon"; exit 2 }
[xml]$ribbon | Out-Null
Write-Output ("OK ({0} bits)" -f $bits)
exit 0
