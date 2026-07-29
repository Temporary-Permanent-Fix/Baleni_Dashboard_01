param(
    [int]$PollSeconds = 60,
    [switch]$RunOnce
)

$ErrorActionPreference = "Stop"

$ProjectDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$WatcherRoot = Join-Path $env:TEMP "baleni-dashboard-n8n-watcher"
$CloneDir = Join-Path $WatcherRoot "repo"
$TriggerFile = Join-Path $CloneDir "n8n\refresh.request.json"
$RefreshScript = Join-Path $CloneDir "scripts\n8n_refresh_and_push.ps1"
$HelperScripts = @(
    "scripts\n8n_refresh_and_push.ps1",
    "scripts\refresh_local_dashboard.ps1",
    "scripts\send_dashboard_email.py",
    "scripts\send_dashboard_email_outlook.ps1"
)

function Get-OriginUrl {
    $url = git -C $ProjectDir remote get-url origin 2>$null
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($url)) {
        throw "Nenasiel som remote origin v hlavnom repozitari."
    }
    return $url.Trim()
}

function Ensure-CloneWorkspace {
    if (Test-Path -LiteralPath (Join-Path $CloneDir ".git")) {
        Ensure-HelperScripts
        return
    }

    New-Item -ItemType Directory -Force -Path $WatcherRoot | Out-Null
    $originUrl = Get-OriginUrl
    $cloneArgs = @(
        "clone"
        "--quiet"
        "--branch"
        "main"
        "--single-branch"
        $originUrl
        $CloneDir
    )

    & git @cloneArgs 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "git clone failed."
    }

    Ensure-HelperScripts
}

function Ensure-HelperScripts {
    foreach ($relativePath in $HelperScripts) {
        $sourcePath = Join-Path $ProjectDir $relativePath
        $targetPath = Join-Path $CloneDir $relativePath
        $targetDir = Split-Path -Parent $targetPath
        if (-not (Test-Path -LiteralPath $sourcePath)) {
            throw "Nenasiel som helper script v hlavnom repozitari: $sourcePath"
        }

        New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
        Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
    }
}

function Get-FileSignature {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return ""
    }

    return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash
}

function Remove-UntrackedHelperScripts {
    foreach ($relativePath in $HelperScripts) {
        $targetPath = Join-Path $CloneDir $relativePath
        if (-not (Test-Path -LiteralPath $targetPath)) {
            continue
        }

        & git -C $CloneDir ls-files --error-unmatch -- $relativePath 2>$null | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Remove-Item -LiteralPath $targetPath -Force
        }
    }
}

function Test-RepositoryClean {
    $status = git -C $CloneDir status --porcelain
    if ($LASTEXITCODE -ne 0) {
        return $false
    }

    $status = $status | Where-Object {
        $_ -notmatch '^\?\?\s+scripts/n8n_refresh_and_push\.ps1$' -and
        $_ -notmatch '^\?\?\s+scripts/refresh_local_dashboard\.ps1$' -and
        $_ -notmatch '^\?\?\s+scripts/send_dashboard_email\.py$' -and
        $_ -notmatch '^\?\?\s+scripts/send_dashboard_email_outlook\.ps1$'
    }
    return [string]::IsNullOrWhiteSpace($status)
}

function Sync-Repository {
    Remove-UntrackedHelperScripts

    if (-not (Test-RepositoryClean)) {
        Write-Warning "Repository is not clean. Skipping pull so local changes are not overwritten."
        return $false
    }

    & git -C $CloneDir pull --ff-only --quiet 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "git pull failed."
        return $false
    }

    return $true
}

Ensure-CloneWorkspace

if (-not (Test-Path -LiteralPath $TriggerFile)) {
    throw "Trigger file not found in clone workspace: $TriggerFile"
}

if (-not (Test-Path -LiteralPath $RefreshScript)) {
    throw "Could not find refresh script: $RefreshScript"
}

$script:lastHandledSignature = Get-FileSignature -Path $TriggerFile

if (-not (Sync-Repository)) {
    Write-Warning "Repository sync failed. I will try again on the next cycle."
}

while ($true) {
    if (Sync-Repository) {
        $currentSignature = Get-FileSignature -Path $TriggerFile
        if ($currentSignature -and $currentSignature -ne $script:lastHandledSignature) {
            Write-Host "New trigger detected from n8n. Running refresh..."
            & powershell -NoProfile -ExecutionPolicy Bypass -File $RefreshScript
            if ($LASTEXITCODE -ne 0) {
                throw "Refresh script failed."
            }

            $script:lastHandledSignature = $currentSignature
            Write-Host "Refresh finished. Waiting for the next trigger."
        }
    }

    if ($RunOnce) {
        break
    }

    Start-Sleep -Seconds $PollSeconds
}
