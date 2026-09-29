<#
    Vérification de l'import Excel → Word sur le .NET Framework réellement utilisé par Word
    (Windows PowerShell 5.1 fonctionne sur .NET Framework 4.x ; les tests unitaires, eux, tournent sur .NET 8).
    Charge la DLL compilée pour .NET Framework 4.0 et contrôle les valeurs affichées, la lecture du classeur
    de test et la génération du tableau Word.
#>
param(
    [Parameter(Mandatory = $true)][string]$Dll,
    [Parameter(Mandatory = $true)][string]$Fixture
)

$ErrorActionPreference = 'Stop'
Add-Type -Path (Resolve-Path $Dll).Path
Write-Host ".NET Framework : $([Environment]::Version)"

$failures = New-Object System.Collections.Generic.List[string]
function Check([string]$name, [string]$actual, [string]$expected) {
    if ($actual -ne $expected) { $failures.Add("$name : obtenu « $actual », attendu « $expected »") }
    else { Write-Host "OK  $name = $actual" }
}

# --- Formats de nombre (culture en-US : séparateurs connus).
$enUs = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$settings = New-Object WordTableToExcel.Core.ExcelImport.ExcelFormatSettings($enUs, $enUs, $false, $null)
$cases = @(
    @(1.005, '0.00', '1.01'),
    @(2.675, '0.00', '2.68'),
    @(1234567.891, '#,##0.00', '1,234,567.89'),
    @(0.30000000000000004, 'General', '0.3'),
    @(123456789012, 'General', '1.23457E+11'),
    @(0.333333333333333, 'General', '0.333333333'),
    @(12345, '0.00E+00', '1.23E+04'),
    @(3.14159265, '# ??/??', '3 14/99'),
    @(45366.75, 'dd/mm/yyyy hh:mm', '15/03/2024 18:00'),
    @(45366.75, 'dddd, mmmm d, yyyy', 'Friday, March 15, 2024'),
    @(1.52, '[h]:mm', '36:28'),
    @(-1234.567, '#,##0.00;(#,##0.00)', '(1,234.57)'),
    @(0.256, '0.0%', '25.6%')
)
foreach ($c in $cases) {
    $format = $settings.Get([string]$c[1])
    Check ("{0} au format {1}" -f $c[0], $c[1]) $format.FormatNumber([double]$c[0], $settings).Text ([string]$c[2])
}

# --- Classeur de test (culture fr-FR).
$frFr = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
$workbook = [WordTableToExcel.Core.ExcelImport.XlsxWorkbook]::Load((Resolve-Path $Fixture).Path)
$format = New-Object WordTableToExcel.Core.ExcelImport.ExcelFormatSettings($frFr, $frFr, $workbook.Date1904, $workbook.Styles.Colors)
$options = New-Object WordTableToExcel.Core.ExcelImport.ExcelImportOptions
$options.AvailableWidthPt = 453.5
$converter = New-Object WordTableToExcel.Core.ExcelImport.SheetConverter($workbook, $format, $options, $null)

function CellText($table, [int]$row, [int]$column) {
    foreach ($cell in $table.Cells) {
        if ($cell.Row -eq $row -and $cell.Column -eq $column) {
            return $cell.PlainText.Replace([string][char]0x202F, ' ').Replace([string][char]0x00A0, ' ')
        }
    }
    return '(absente)'
}

$sheet = $workbook.ReadSheet($workbook.Sheets[0])
$result = $converter.Convert($sheet, $null)
if (-not $result.Succeeded) { $failures.Add("Conversion « Ventes 2024 » : $($result.Error)") }
else {
    Check 'Légende' $result.Caption 'Tableau 1 : Ventes trimestrielles par région'
    Check 'Taille' ("{0}x{1}" -f $result.Table.RowCount, $result.Table.ColumnCount) '9x7'
    Check 'B4' (CellText $result.Table 1 1) '125 430,50 €'
    Check 'G5' (CellText $result.Table 2 6) '-1,1%'
    Check 'Total' (CellText $result.Table 6 5) '1 671 156,50 €'
    $xml = [WordTableToExcel.Core.ExcelImport.WordTableXmlWriter]::BuildFlatOpc($result.Table)
    [xml]$parsed = $xml
    Write-Host ("OK  XML Word : {0} caractères" -f $xml.Length)
}

$sheet = $workbook.ReadSheet($workbook.Sheets[1])
$result = $converter.Convert($sheet, $null)
Check 'Date longue' (CellText $result.Table 1 2) 'lundi 8 janvier 2024'
Check 'Date et heure' (CellText $result.Table 1 3) '19/01/2024 17:30'
Check 'Durée' (CellText $result.Table 1 4) '36:00'
Check 'Booléen' (CellText $result.Table 1 5) 'VRAI'
Check 'Scientifique' (CellText $result.Table 7 2) '1,23E-04'
Check 'Standard' (CellText $result.Table 7 4) '1,23457E+12'

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Host "ÉCHEC $_" }
    throw "$($failures.Count) vérification(s) en échec sur .NET Framework."
}
Write-Host 'Import vérifié sur .NET Framework.'
exit 0
