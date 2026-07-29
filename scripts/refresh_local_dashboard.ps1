param(
    [string]$ExcelPath = ""
)

$ErrorActionPreference = "Stop"

$ProjectDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$BuildScript = Join-Path $ProjectDir "scripts\build_dashboard.py"
$PublishScript = Join-Path $ProjectDir "scripts\publish_to_docs.ps1"

if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
    & python $BuildScript
} else {
    & python $BuildScript --input $ExcelPath
}

if ($LASTEXITCODE -ne 0) {
    throw "Build dashboard script failed with exit code $LASTEXITCODE."
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $PublishScript

if ($LASTEXITCODE -ne 0) {
    throw "Publish to docs script failed with exit code $LASTEXITCODE."
}

Write-Host "Hotovo."
