param(
    [string]$ExcelPath = "",
    [switch]$AllowStale
)

$ErrorActionPreference = "Stop"

$ProjectDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$RefreshExcelScript = Join-Path $ProjectDir "scripts\refresh_input_excel.ps1"
$BuildScript = Join-Path $ProjectDir "scripts\build_dashboard.py"
$PublishScript = Join-Path $ProjectDir "scripts\publish_to_docs.ps1"

$buildArgs = @()
if (-not [string]::IsNullOrWhiteSpace($ExcelPath)) {
    $buildArgs += @("--input", $ExcelPath)
}
if ($AllowStale) {
    $buildArgs += "--allow-stale"
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $RefreshExcelScript -ExcelPath $ExcelPath

if ($LASTEXITCODE -ne 0) {
    throw "Excel refresh script failed with exit code $LASTEXITCODE."
}

& python $BuildScript @buildArgs

if ($LASTEXITCODE -ne 0) {
    throw "Build dashboard script failed with exit code $LASTEXITCODE."
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $PublishScript

if ($LASTEXITCODE -ne 0) {
    throw "Publish to docs script failed with exit code $LASTEXITCODE."
}

Write-Host "Hotovo."
