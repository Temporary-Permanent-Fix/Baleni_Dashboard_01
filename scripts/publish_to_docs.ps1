param(
    [string]$SourceFile = ""
)

$ErrorActionPreference = "Stop"

$ProjectDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$OutputDir = Join-Path $ProjectDir "output"
$DocsDir = Join-Path $ProjectDir "docs"
$SnapshotFile = Join-Path $DocsDir "vyvoj-balenia.html"

if ([string]::IsNullOrWhiteSpace($SourceFile)) {
    $SourceFile = Join-Path $OutputDir "packaging_dashboard.html"
}

if (-not (Test-Path -LiteralPath $SourceFile)) {
    throw "Zdrojovy HTML subor neexistuje: $SourceFile"
}

New-Item -ItemType Directory -Force -Path $DocsDir | Out-Null
Copy-Item -LiteralPath $SourceFile -Destination $SnapshotFile -Force
New-Item -ItemType File -Force -Path (Join-Path $DocsDir ".nojekyll") | Out-Null

Write-Host "Hotovo."
Write-Host "GitHub Pages snapshot: $SnapshotFile"
