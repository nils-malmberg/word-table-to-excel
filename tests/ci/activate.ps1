# Crée l'objet COM du complément comme le fait Word, puis demande la définition du ruban.
# Le résultat (« OK … » ou « ÉCHEC … ») est écrit dans -OutFile (utile quand le script est lancé via runas).
param([Parameter(Mandatory = $true)][string]$OutFile)

$bits = [IntPtr]::Size * 8
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$elevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$context = "{0} bits, élevé : {1}" -f $bits, $elevated

try {
    $addin = New-Object -ComObject WordTableToExcel.Connect
    $ribbon = $addin.GetCustomUI('Microsoft.Word.Document')
    if ($ribbon -notmatch 'OnExportClick') { throw "ruban invalide : $ribbon" }
    [xml]$ribbon | Out-Null
    $result = "OK ($context)"
} catch {
    $e = $_.Exception
    while ($e.InnerException) { $e = $e.InnerException }
    $result = "ÉCHEC ({0}) : {1} HRESULT=0x{2:X8} {3}" -f $context, $e.GetType().FullName, $e.HResult, $e.Message
}
[IO.File]::WriteAllText($OutFile, $result)
