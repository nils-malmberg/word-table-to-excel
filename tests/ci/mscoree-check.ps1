<#
    Expérience : enregistrement « .NET standard » (comme regasm et comme l'extension Outlook RBLclass) :
    InprocServer32 = mscoree.dll (chargeur .NET de Microsoft) + CodeBase vers WordTableToExcel.dll,
    dans HKCU (vues 64 et 32 bits). Activation par un client COM natif (cscript) 64 et 32 bits,
    en utilisateur standard puis administrateur. Affiche les résultats, puis nettoie.
#>
param([Parameter(Mandatory = $true)][string]$Folder)

$ErrorActionPreference = 'Stop'
$clsid = '{03F63233-F2FE-4A75-AF7A-F99CBEDC8030}'
$progId = 'WordTableToExcel.Connect'
$activate = Join-Path $PSScriptRoot 'activate.vbs'
$dir = Join-Path $env:LOCALAPPDATA 'WordTableToExcel'
New-Item -ItemType Directory -Force $dir | Out-Null
Copy-Item (Join-Path $Folder 'WordTableToExcel.dll') $dir -Force
$dll = Join-Path $dir 'WordTableToExcel.dll'
$identity = [Reflection.AssemblyName]::GetAssemblyName($dll)
$assembly = $identity.FullName
$version = $identity.Version.ToString()
$codeBase = 'file:///' + $dll.Replace('\', '/')
Write-Host "Assembly = $assembly"
Write-Host "CodeBase = $codeBase"

foreach ($root in 'HKCU:\Software\Classes\CLSID', 'HKCU:\Software\Classes\Wow6432Node\CLSID') {
    $key = "$root\$clsid"
    Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue
    New-Item $key -Force | Set-ItemProperty -Name '(default)' -Value $progId
    foreach ($server in "$key\InprocServer32", "$key\InprocServer32\$version") {
        New-Item $server -Force | Out-Null
        if ($server -eq "$key\InprocServer32") {
            Set-ItemProperty $server -Name '(default)' -Value 'mscoree.dll'
            Set-ItemProperty $server -Name 'ThreadingModel' -Value 'Both'
        }
        Set-ItemProperty $server -Name 'Class' -Value 'WordTableToExcel.AddIn.Connect'
        Set-ItemProperty $server -Name 'Assembly' -Value $assembly
        Set-ItemProperty $server -Name 'RuntimeVersion' -Value 'v4.0.30319'
        Set-ItemProperty $server -Name 'CodeBase' -Value $codeBase
    }
    New-Item "$key\ProgId" -Force | Set-ItemProperty -Name '(default)' -Value $progId
    New-Item "$key\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" -Force | Out-Null
}
New-Item "HKCU:\Software\Classes\$progId\CLSID" -Force | Set-ItemProperty -Name '(default)' -Value $clsid

function Invoke-Activation([string]$cscript, [bool]$standardUser) {
    $out = Join-Path $env:TEMP ('mscoree-' + [Guid]::NewGuid().ToString('N') + '.txt')
    if ($standardUser) {
        & runas.exe /trustlevel:0x20000 "$cscript //nologo $activate $out" | Out-Null
        for ($i = 0; $i -lt 120 -and -not (Test-Path $out); $i++) { Start-Sleep -Milliseconds 500 }
        Start-Sleep -Milliseconds 300
    } else {
        & $cscript //nologo $activate $out | Out-Host
    }
    if (-not (Test-Path $out)) { return '(aucun résultat)' }
    $r = [IO.File]::ReadAllText($out)
    Remove-Item $out -ErrorAction SilentlyContinue
    return $r
}

$results = @()
foreach ($cscript in "$env:WINDIR\System32\cscript.exe", "$env:WINDIR\SysWOW64\cscript.exe") {
    foreach ($standard in $true, $false) {
        $r = Invoke-Activation $cscript $standard
        $line = "MSCOREE | $cscript | utilisateur standard=$standard | $r"
        Write-Host $line
        $results += $line
    }
}

foreach ($root in 'HKCU:\Software\Classes\CLSID', 'HKCU:\Software\Classes\Wow6432Node\CLSID') {
    Remove-Item "$root\$clsid" -Recurse -Force -ErrorAction SilentlyContinue
}
Remove-Item "HKCU:\Software\Classes\$progId" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
if ($results | Where-Object { $_ -notmatch '\| OK' }) { Write-Host 'Au moins une activation a échoué.'; exit 1 }
Write-Host 'Toutes les activations ont réussi.'
exit 0
