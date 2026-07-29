param(
    [string]$DailyKpiPath = "",
    [string]$To = "peter.kadlec@alza.sk",
    [string]$From = ""
)

$ErrorActionPreference = "Stop"

$ProjectDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$LogPath = Join-Path $env:TEMP "baleni-dashboard-outlook-mail.log"

function Write-Log {
    param([string]$Message)
    Add-Content -LiteralPath $LogPath -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

function Strip-LeadingEquals {
    param([AllowNull()][object]$Value)

    return ([string]$Value) -replace '^\s*=+\s*', ''
}

Set-Content -LiteralPath $LogPath -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] start"

if ([string]::IsNullOrWhiteSpace($DailyKpiPath)) {
    $DailyKpiPath = Join-Path $ProjectDir "output\daily_kpi.json"
}
if (-not [System.IO.Path]::IsPathRooted($DailyKpiPath)) {
    $DailyKpiPath = Join-Path $ProjectDir $DailyKpiPath
}
if (-not (Test-Path -LiteralPath $DailyKpiPath)) {
    throw "Daily KPI file not found: $DailyKpiPath"
}

Write-Log "Loading daily KPI from $DailyKpiPath"
$summary = Get-Content -LiteralPath $DailyKpiPath -Raw -Encoding UTF8 | ConvertFrom-Json
Write-Log "Daily KPI loaded"

$targetDay = [datetime]::ParseExact([string]$summary.target_day, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
$dateText = $targetDay.ToString('d.M.yyyy', [System.Globalization.CultureInfo]::GetCultureInfo('sk-SK'))

$bodyLines = @()
foreach ($row in @($summary.sheet_rows)) {
    $ratioText = if ($null -eq $row.ratio) { 'bez dat' } else { ([math]::Round(([double]$row.ratio) * 100, 1)).ToString('N1', [System.Globalization.CultureInfo]::GetCultureInfo('sk-SK')) + ' %' }
    $bodyLines += "$($row.sheet): $ratioText eliminace z geosize = SPO, doprava = alzabox"
}

$bodyLine = $bodyLines -join "`r`n"
$htmlBody = @"
<!doctype html>
<html lang="sk">
  <body style="font-family: Arial, sans-serif; color: #15202b;">
    <p style="margin: 0; white-space: pre-line;">$([System.Security.SecurityElement]::Escape($bodyLine))</p>
  </body>
</html>
"@

Write-Log "Connecting to Outlook"
$outlook = New-Object -ComObject Outlook.Application
Write-Log "Outlook connected"

Write-Log "Creating mail item"
$mail = $outlook.CreateItem(0)
if (-not [string]::IsNullOrWhiteSpace($From)) {
    try {
        $mail.SentOnBehalfOfName = (Strip-LeadingEquals $From)
    } catch {
        Write-Warning "Could not set SentOnBehalfOfName: $($_.Exception.Message)"
    }
}

$mail.To = Strip-LeadingEquals $To
$mail.Subject = Strip-LeadingEquals "Balenie dashboard - $dateText"
$mail.Body = $bodyLine
$mail.HTMLBody = $htmlBody

Write-Log "Sending mail"
$mail.Send()
Write-Log "Done"
