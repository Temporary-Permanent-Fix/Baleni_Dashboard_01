param(
    [string]$ExcelPath = "",
    [int]$TimeoutMinutes = 15
)

$ErrorActionPreference = "Stop"

$ProjectDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$InputDir = Join-Path $ProjectDir "input"

function Get-LatestPackagingExcel {
    param([string]$Folder)

    $files = Get-ChildItem -LiteralPath $Folder -File |
        Where-Object { $_.Extension -in ".xlsx", ".xls", ".xlsm" -and -not $_.Name.StartsWith("~$") } |
        Sort-Object LastWriteTime -Descending

    if (-not $files) {
        throw "V zlozke input nie je ziaden Excel subor."
    }

    $preferred = $files | Where-Object { $_.Name -match "Data_pro" -and $_.Name -notmatch "balikovka" } | Select-Object -First 1
    if ($preferred) {
        return $preferred.FullName
    }

    return $files[0].FullName
}

if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
    $ExcelPath = Get-LatestPackagingExcel -Folder $InputDir
}

$ExcelPath = (Resolve-Path -LiteralPath $ExcelPath).Path

if (-not (Test-Path -LiteralPath $ExcelPath)) {
    throw "Excel subor neexistuje: $ExcelPath"
}

Write-Host "Otváram Excel workbook: $ExcelPath"

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    try {
        $excel.AskToUpdateLinks = $false
    } catch {
        # Not available in every Excel version; safe to ignore.
    }

    $openStart = Get-Date
    $workbook = $excel.Workbooks.Open($ExcelPath, 0, $false)
    Write-Host "Spúšťam Refresh All..."
    $workbook.RefreshAll()

    try {
        $excel.CalculateUntilAsyncQueriesDone()
    } catch {
        # Some Excel versions do not expose this method; we will still poll below.
    }

    $done = $false
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while (-not $done -and (Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 2
        $done = $true

        try {
            if ($excel.CalculationState -ne 0) {
                $done = $false
            }
        } catch {
            # Ignore and rely on time limit.
        }

        try {
            foreach ($connection in @($workbook.Connections)) {
                try {
                    if ($connection.ODBCConnection -and $connection.ODBCConnection.Refreshing) {
                        $done = $false
                        break
                    }
                } catch {
                }
                try {
                    if ($connection.OLEDBConnection -and $connection.OLEDBConnection.Refreshing) {
                        $done = $false
                        break
                    }
                } catch {
                }
            }
        } catch {
            # If connection inspection fails, fall back to calculation state only.
        }
    }

    if (-not $done) {
        throw "Refresh v Exceli neskončil do $TimeoutMinutes minút."
    }

    $workbook.Save()

    $newWriteTime = (Get-Item -LiteralPath $ExcelPath).LastWriteTime
    Write-Host ("Excel uložený. Pôvodne otvorený o {0}, súbor má teraz LastWriteTime {1:yyyy-MM-dd HH:mm:ss}." -f $openStart.ToString("yyyy-MM-dd HH:mm:ss"), $newWriteTime)
} finally {
    if ($workbook) {
        try { $workbook.Close($true) } catch {}
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
    }

    foreach ($obj in @($workbook, $excel)) {
        if ($obj) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
