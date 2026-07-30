param(
    [string]$WorkbookPath = "working\AI_YouTube_Content_System_Working.xlsx",
    [string]$OutputPath = "working\AI_YouTube_Content_System_Working.xlsm"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$sourceWorkbook = Join-Path $root $WorkbookPath
$outputWorkbook = Join-Path $root $OutputPath
$moduleDirectory = Join-Path $root "src\vba"

if (-not (Test-Path -LiteralPath $sourceWorkbook)) { throw "Workbook was not found: $sourceWorkbook" }
if (-not (Test-Path -LiteralPath $moduleDirectory)) { throw "VBA module directory was not found: $moduleDirectory" }

$securityKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$accessVBOM = (Get-ItemProperty -LiteralPath $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
if ($accessVBOM -ne 1) {
    throw "Excel Trust Center must enable 'Trust access to the VBA project object model' before module import. The script does not modify this security setting."
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $workbook = $excel.Workbooks.Open($sourceWorkbook)
    foreach ($module in Get-ChildItem -LiteralPath $moduleDirectory -Filter "*.bas") {
        $workbook.VBProject.VBComponents.Import($module.FullName) | Out-Null
    }
    $workbook.SaveAs($outputWorkbook, 52)
    $workbook.Close($true)
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
