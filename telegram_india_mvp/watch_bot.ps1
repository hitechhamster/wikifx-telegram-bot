param(
    [int]$CheckIntervalSeconds = 5,
    [int]$PendingThreshold = 2
)

$ErrorActionPreference = "Stop"
$projectPath = $PSScriptRoot
$projectParent = Split-Path -Parent $projectPath
$runnerPath = Join-Path $projectPath "run_bot_now.ps1"
$watchdogLog = Join-Path $projectPath "bot_watchdog.log"
$staleChecks = 0

function Write-WatchdogLog([string]$Message) {
    "$(Get-Date -Format s) $Message" | Out-File -LiteralPath $watchdogLog -Encoding utf8 -Append
}

function Get-BotProcesses {
    return @(Get-CimInstance Win32_Process | Where-Object {
        $_.Name -match '^python(w)?\.exe$' -and
        $_.CommandLine -like "*$projectPath*" -and
        $_.CommandLine -like '*bot.py*'
    })
}

function Start-BotProcess {
    Start-Process -FilePath "powershell.exe" `
        -ArgumentList @(
            "-NoLogo",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            ('"' + $runnerPath + '"')
        ) `
        -WindowStyle Hidden | Out-Null
    Write-WatchdogLog "Bot launcher started."
}

function Restart-BotProcess([string]$Reason) {
    $botProcesses = Get-BotProcesses
    $launcherIds = @($botProcesses | ForEach-Object { $_.ParentProcessId } | Select-Object -Unique)
    foreach ($botProcess in $botProcesses) {
        Stop-Process -Id $botProcess.ProcessId -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2
    foreach ($launcherId in $launcherIds) {
        $launcher = Get-CimInstance Win32_Process -Filter "ProcessId = $launcherId" -ErrorAction SilentlyContinue
        if ($launcher -and $launcher.Name -eq "powershell.exe" -and $launcher.CommandLine -like '*run_bot_now.ps1*') {
            Stop-Process -Id $launcherId -Force -ErrorAction SilentlyContinue
        }
    }
    Write-WatchdogLog "Restarting bot: $Reason"
    Start-BotProcess
}

Write-WatchdogLog "Watchdog started. interval=$CheckIntervalSeconds threshold=$PendingThreshold"

while ($true) {
    try {
        $tokenFile = Get-ChildItem -LiteralPath $projectParent -File -Filter "bot api*token.txt" | Select-Object -First 1
        if (-not $tokenFile) {
            Write-WatchdogLog "Token TXT not found."
            Start-Sleep -Seconds $CheckIntervalSeconds
            continue
        }

        $tokenText = Get-Content -LiteralPath $tokenFile.FullName -Raw -Encoding UTF8
        $tokenMatch = [regex]::Match($tokenText, '\b\d{5,}:[A-Za-z0-9_-]{20,}\b')
        if (-not $tokenMatch.Success) {
            Write-WatchdogLog "Valid BotFather token not found."
            Start-Sleep -Seconds $CheckIntervalSeconds
            continue
        }

        $botProcesses = Get-BotProcesses
        if ($botProcesses.Count -eq 0) {
            Restart-BotProcess "no bot process"
            $staleChecks = 0
            Start-Sleep -Seconds $CheckIntervalSeconds
            continue
        }

        $webhookInfo = Invoke-RestMethod -Method Get `
            -Uri ("https://api.telegram.org/bot" + $tokenMatch.Value + "/getWebhookInfo") `
            -TimeoutSec 8
        $pendingUpdates = [int]$webhookInfo.result.pending_update_count

        if ($pendingUpdates -gt 0) {
            $staleChecks += 1
            Write-WatchdogLog "Pending updates detected: $pendingUpdates; stale check $staleChecks/$PendingThreshold."
        }
        else {
            $staleChecks = 0
        }

        if ($staleChecks -ge $PendingThreshold) {
            Restart-BotProcess "pending updates remained at $pendingUpdates"
            $staleChecks = 0
        }
    }
    catch {
        Write-WatchdogLog "Health check failed: $($_.Exception.GetType().Name)"
    }
    finally {
        $tokenText = $null
        $tokenMatch = $null
        $webhookInfo = $null
    }

    Start-Sleep -Seconds $CheckIntervalSeconds
}
