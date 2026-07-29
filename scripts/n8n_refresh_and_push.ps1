param(
    [string]$ExcelPath = "",
    [string]$CommitMessage = ""
)

$ErrorActionPreference = "Stop"

$ProjectDir = Resolve-Path (Join-Path $PSScriptRoot "..")

if ([string]::IsNullOrWhiteSpace($CommitMessage)) {
    $dateStamp = Get-Date -Format "yyyy-MM-dd HH:mm"
    $CommitMessage = "Auto-refresh dashboard $dateStamp"
}

if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
    $inputDir = Join-Path $ProjectDir "input"
    $ExcelPath = Get-ChildItem -LiteralPath $inputDir -File |
        Where-Object { $_.Extension -in ".xlsx", ".xls", ".xlsm" -and -not $_.Name.StartsWith("~$") } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1 -ExpandProperty FullName
}

& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ProjectDir "scripts\refresh_local_dashboard.ps1") -ExcelPath $ExcelPath

if ($LASTEXITCODE -ne 0) {
    throw "Local dashboard refresh failed with exit code $LASTEXITCODE."
}

& git -C $ProjectDir add --all -- "input" "output" "docs" "n8n" "scripts"

if ($LASTEXITCODE -ne 0) {
    throw "git add failed with exit code $LASTEXITCODE."
}

$status = git -C $ProjectDir status --porcelain
if ([string]::IsNullOrWhiteSpace($status)) {
    Write-Host "Ziadne zmeny na commit."
    exit 0
}

& git -C $ProjectDir commit -m $CommitMessage
if ($LASTEXITCODE -ne 0) {
    throw "git commit failed with exit code $LASTEXITCODE."
}

& git -C $ProjectDir push origin HEAD:main
if ($LASTEXITCODE -ne 0) {
    throw "git push failed with exit code $LASTEXITCODE."
}

Write-Host "Hotovo."
