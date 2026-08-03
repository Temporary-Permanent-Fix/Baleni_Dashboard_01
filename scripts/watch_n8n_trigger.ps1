param(
    [int]$PollSeconds = 60,
    [switch]$RunOnce,
    [string]$LogPath = ""
)

$ErrorActionPreference = "Stop"

$ProjectDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$WatcherRoot = Join-Path $env:TEMP "baleni-dashboard-n8n-watcher"
$CloneDir = Join-Path $WatcherRoot "repo"
$LogDir = Join-Path $WatcherRoot "logs"
$ResolvedLogPath = if ([string]::IsNullOrWhiteSpace($LogPath)) {
    Join-Path $LogDir "watcher.log"
} else {
    $LogPath
}
$TriggerFile = Join-Path $CloneDir "n8n\refresh.request.json"
$RefreshScript = Join-Path $ProjectDir "scripts\n8n_refresh_and_push.ps1"

function Get-OriginUrl {
    $url = git -C $ProjectDir remote get-url origin 2>$null
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($url)) {
        throw "Nenasiel som remote origin v hlavnom repozitari."
    }

    return $url.Trim()
}

function Ensure-CloneWorkspace {
    if (Test-Path -LiteralPath (Join-Path $CloneDir ".git")) {
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
}

function Reset-CloneWorkspace {
    & git -C $CloneDir reset --hard --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "git reset failed."
        return $false
    }

    & git -C $CloneDir clean -fd --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "git clean failed."
        return $false
    }

    return $true
}

function Get-FileSignature {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return ""
    }

    return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash
}

function Sync-Repository {
    if (-not (Reset-CloneWorkspace)) {
        return $false
    }

    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "SilentlyContinue"
    try {
        & git -C $CloneDir pull --ff-only --quiet 2>$null | Out-Null
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Warning "git pull failed."
        return $false
    }

    return $true
}

Ensure-CloneWorkspace
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

try {
    try {
        Start-Transcript -Path $ResolvedLogPath -Append | Out-Null
    } catch {
        Write-Warning "Nepodarilo sa zapnut transcript logovanie: $($_.Exception.Message)"
    }

    Write-Host "Watcher log: $ResolvedLogPath"

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
} catch {
    throw
} finally {
    try {
        Stop-Transcript | Out-Null
    } catch {
        # Transcript may already be stopped or unavailable; ignore on exit.
    }
}
