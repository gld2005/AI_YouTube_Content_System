param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"
$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Set-WorkbookName {
    param(
        [Parameter(Mandatory = $true)][object]$Workbook,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$RefersTo
    )

    try {
        $Workbook.Names.Item($Name).Delete()
    }
    catch {
        # The name does not yet exist.
    }
    $Workbook.Names.Add($Name, $RefersTo) | Out-Null
}

try {
    $workbook = $excel.Workbooks.Open($resolvedWorkbookPath)
    $settingsSheet = $workbook.Worksheets.Item(21)
    foreach ($technicalSheetIndex in 24..27) {
        $workbook.Worksheets.Item($technicalSheetIndex).Visible = 2
    }

    Set-WorkbookName -Workbook $workbook -Name "nrDesignTokens" -RefersTo "='Design Tokens'!`$A`$1:`$C`$10"
    Set-WorkbookName -Workbook $workbook -Name "nrParagraphTypeList" -RefersTo "=_Lists!`$A`$2:`$A`$7"
    Set-WorkbookName -Workbook $workbook -Name "nrParagraphStatusList" -RefersTo "=_Lists!`$B`$2:`$B`$8"
    Set-WorkbookName -Workbook $workbook -Name "nrProjectRoot" -RefersTo ("='" + $settingsSheet.Name + "'!`$B`$4")
    Set-WorkbookName -Workbook $workbook -Name "nrAssistantBaseUrl" -RefersTo ("='" + $settingsSheet.Name + "'!`$B`$5")
    Set-WorkbookName -Workbook $workbook -Name "nrAssistantVersion" -RefersTo ("='" + $settingsSheet.Name + "'!`$B`$6")

    $workbook.Save()
    $workbook.Close($true)
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
