param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,
    [Parameter(Mandatory = $true)]
    [string]$NodePath
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$builderPath = Join-Path $root "scripts\build_r2_workbook.mjs"
$finalizerPath = Join-Path $root "scripts\finalize_r2_workbook.ps1"

& $NodePath $builderPath $InputPath $OutputPath
if ($LASTEXITCODE -ne 0) {
    throw "The R2 workbook builder failed."
}

& $finalizerPath -WorkbookPath $OutputPath
