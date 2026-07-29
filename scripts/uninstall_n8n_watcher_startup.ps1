param(
    [string]$TaskName = "Balení dashboard n8n watcher"
)

$ErrorActionPreference = "Stop"

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Odstranene."
} else {
    Write-Host "Task neexistuje."
}
