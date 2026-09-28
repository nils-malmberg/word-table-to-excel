# Diagnostic temporaire : quelles variantes d'enregistrement COM par utilisateur (HKCU) sont activables ?
param([Parameter(Mandatory = $true)][string]$Dist)
$ErrorActionPreference = 'Continue'
$dist = (Resolve-Path $Dist).Path
$clsid = '{03F63233-F2FE-4A75-AF7A-F99CBEDC8030}'
$activate = Join-Path $PSScriptRoot 'activate.vbs'
$dir = Join-Path $env:LOCALAPPDATA 'wtte-exp'
New-Item -ItemType Directory -Force $dir | Out-Null
$dll = Join-Path $dir 'WordTableToExcel.dll'
Copy-Item (Join-Path $dist 'WordTableToExcel.dll') $dll -Force

function Clean {
    foreach ($v in '/reg:32', '/reg:64') {
        & reg.exe delete "HKCU\Software\Classes\CLSID\$clsid" /f $v 2>&1 | Out-Null
        & reg.exe delete "HKCU\Software\Classes\WordTableToExcel.Connect" /f $v 2>&1 | Out-Null
        & reg.exe delete "HKCU\Software\Classes\Record" /f $v 2>&1 | Out-Null
    }
}

function Try-Activate([string]$label) {
    foreach ($cscript in "$env:WINDIR\System32\cscript.exe", "$env:WINDIR\SysWOW64\cscript.exe") {
        $out = Join-Path $env:TEMP ('exp-' + [Guid]::NewGuid().ToString('N') + '.txt')
        & runas.exe /trustlevel:0x20000 "$cscript //nologo $activate $out" | Out-Null
        for ($i = 0; $i -lt 60 -and -not (Test-Path $out); $i++) { Start-Sleep -Milliseconds 500 }
        Start-Sleep -Milliseconds 300
        $r = if (Test-Path $out) { [IO.File]::ReadAllText($out) } else { '(pas de résultat)' }
        Write-Host ("[{0}] {1} (non élevé) -> {2}" -f $label, (Split-Path (Split-Path $cscript) -Leaf), $r)
        $out2 = Join-Path $env:TEMP ('exp-' + [Guid]::NewGuid().ToString('N') + '.txt')
        & $cscript //nologo $activate $out2 | Out-Null
        Write-Host ("[{0}] {1} (élevé)     -> {2}" -f $label, (Split-Path (Split-Path $cscript) -Leaf), [IO.File]::ReadAllText($out2))
    }
}

# A. Script d'installation actuel.
Clean
& (Join-Path $dist 'install.ps1') -SourceDll (Join-Path $dist 'WordTableToExcel.dll') -InstallDir $dir | Out-Null
Try-Activate 'A install.ps1'
& reg.exe export "HKCU\Software\Classes\CLSID\$clsid" (Join-Path $env:TEMP 'a.reg') /y /reg:64 | Out-Null
Get-Content (Join-Path $env:TEMP 'a.reg') | Write-Host

# B. Même chose, InprocServer32 = chemin complet de mscoree.dll.
foreach ($v in '/reg:32', '/reg:64') { & reg.exe add "HKCU\Software\Classes\CLSID\$clsid\InprocServer32" /ve /t REG_SZ /d "$env:WINDIR\System32\mscoree.dll" /f $v | Out-Null }
Try-Activate 'B chemin complet'

# C. Fichier .reg produit par regasm, importé dans HKCU\Software\Classes.
Clean
$regasm = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$regfile = Join-Path $env:TEMP 'regasm.reg'
& $regasm /codebase /nologo $dll "/regfile:$regfile" | Out-Null
$content = (Get-Content $regfile -Raw) -replace 'HKEY_CLASSES_ROOT', 'HKEY_CURRENT_USER\Software\Classes'
Write-Host '--- regasm /regfile ---'
Write-Host $content
$hkcu = Join-Path $env:TEMP 'regasm-hkcu.reg'
Set-Content -Path $hkcu -Value $content -Encoding Unicode
& reg.exe import $hkcu /reg:64 2>&1 | Out-Host
& reg.exe import $hkcu /reg:32 2>&1 | Out-Host
Try-Activate 'C regasm vers HKCU'

# D. Clés écrites via le fournisseur Registry de PowerShell (sans reg.exe), vue native uniquement.
Clean
$base = "HKCU:\Software\Classes\CLSID\$clsid"
New-Item -Force "$base\InprocServer32\1.0.0.0" | Out-Null
New-Item -Force "$base\ProgId" | Out-Null
Set-ItemProperty $base '(default)' 'WordTableToExcel.AddIn.Connect'
Set-ItemProperty "$base\ProgId" '(default)' 'WordTableToExcel.Connect'
foreach ($k in "$base\InprocServer32", "$base\InprocServer32\1.0.0.0") {
    Set-ItemProperty $k 'Class' 'WordTableToExcel.AddIn.Connect'
    Set-ItemProperty $k 'Assembly' 'WordTableToExcel, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
    Set-ItemProperty $k 'RuntimeVersion' 'v4.0.30319'
    Set-ItemProperty $k 'CodeBase' ('file:///' + ($dll -replace '\\', '/'))
}
Set-ItemProperty "$base\InprocServer32" '(default)' 'mscoree.dll'
Set-ItemProperty "$base\InprocServer32" 'ThreadingModel' 'Both'
New-Item -Force 'HKCU:\Software\Classes\WordTableToExcel.Connect\CLSID' | Out-Null
Set-ItemProperty 'HKCU:\Software\Classes\WordTableToExcel.Connect' '(default)' 'WordTableToExcel.AddIn.Connect'
Set-ItemProperty 'HKCU:\Software\Classes\WordTableToExcel.Connect\CLSID' '(default)' $clsid
Try-Activate 'D fournisseur PowerShell'

Clean
& reg.exe delete 'HKCU\Software\Microsoft\Office\Word\Addins\WordTableToExcel.Connect' /f 2>&1 | Out-Null
