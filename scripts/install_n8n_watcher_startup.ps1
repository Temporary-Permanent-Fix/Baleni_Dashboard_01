param(
    [string]$TaskName = "Balení dashboard n8n watcher",
    [int]$PollSeconds = 60
)

$ErrorActionPreference = "Stop"

$ProjectDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$PowerShellExe = (Get-Command powershell.exe).Source
$WatcherScript = Join-Path $ProjectDir "scripts\watch_n8n_trigger.ps1"
$Arguments = @(
    "-NoProfile"
    "-ExecutionPolicy"
    "Bypass"
    "-WindowStyle"
    "Hidden"
    "-File"
    "`"$WatcherScript`""
    "-PollSeconds"
    $PollSeconds
)

if (-not (Test-Path -LiteralPath $WatcherScript)) {
    throw "Nenasiel som watcher script: $WatcherScript"
}

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

$action = New-ScheduledTaskAction -Execute $PowerShellExe -Argument ($Arguments -join " ")
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description "Keeps the packaging dashboard watcher running in the background." | Out-Null

Write-Host "Hotovo."
Write-Host "Task: $TaskName"
Write-Host "Watcher script: $WatcherScript"
