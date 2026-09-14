$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot
$runtimeLog = Join-Path $PSScriptRoot "bot_runtime.log"
$runtimeOutLog = Join-Path $PSScriptRoot "bot_stdout.log"
$runtimeErrorLog = Join-Path $PSScriptRoot "bot_stderr.log"

$existingBot = Get-CimInstance Win32_Process | Where-Object {
    $_.Name -match '^python(w)?\.exe$' -and
    $_.CommandLine -match 'telegram_india_mvp.*bot\.py|bot\.py'
}

if ($existingBot) {
    Write-Host "The Telegram bot is already running. PID: $($existingBot.ProcessId -join ', ')" -ForegroundColor Yellow
    Read-Host "Press Enter to close"
    exit 0
}

$pythonPath = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $pythonPath)) {
    throw "Virtual environment Python was not found: $pythonPath"
}

try {
    $tokenPtr = [IntPtr]::Zero
    $tokenFile = Get-ChildItem -LiteralPath (Split-Path -Parent $PSScriptRoot) -File -Filter "bot api*token.txt" | Select-Object -First 1

    try {
        if ($tokenFile) {
            $tokenText = Get-Content -LiteralPath $tokenFile.FullName -Raw -Encoding UTF8
            $tokenMatch = [regex]::Match($tokenText, '\b\d{5,}:[A-Za-z0-9_-]{20,}\b')
            if (-not $tokenMatch.Success) {
                throw "No valid BotFather token was found in the TXT file."
            }
            $env:TELEGRAM_BOT_TOKEN = $tokenMatch.Value
            Write-Host "Bot token loaded from the project TXT file." -ForegroundColor Cyan
        }
        else {
            Write-Host "Paste the BotFather token, then press Enter. Input will not be displayed." -ForegroundColor Cyan
            $secureToken = Read-Host "Bot Token" -AsSecureString
            $tokenPtr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureToken)
            $env:TELEGRAM_BOT_TOKEN = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($tokenPtr)
        }

        if ([string]::IsNullOrWhiteSpace($env:TELEGRAM_BOT_TOKEN)) {
            throw "The BotFather token cannot be empty."
        }

        $adminId = & $pythonPath -c "import sqlite3; c=sqlite3.connect('telegram_mvp.db'); ids=[str(r[0]) for r in c.execute('SELECT DISTINCT telegram_user_id FROM events WHERE telegram_user_id IS NOT NULL')]; print(ids[0] if len(ids)==1 else '')"
        if ([string]::IsNullOrWhiteSpace($adminId)) {
            $adminId = Read-Host "Enter your numeric Telegram user ID"
        }

        $env:TELEGRAM_ADMIN_IDS = $adminId.Trim()
        $env:WIKIFX_APP_ONELINK = "https://fxeye.onelink.me/Vm4A/intgbot"
        $env:BOT_MARKET = "India"
        $env:SOURCE_TAG = "india_own_group_mvp"
        $env:RETENTION_DAYS = "30"
        $env:PYTHONUNBUFFERED = "1"

        Write-Host "Starting the WikiFX Telegram bot. Keep this window open." -ForegroundColor Green
        "$(Get-Date -Format s) Starting bot process." | Out-File -LiteralPath $runtimeLog -Encoding utf8 -Append
        $botProcess = Start-Process -FilePath $pythonPath `
            -ArgumentList @(".\bot.py") `
            -NoNewWindow `
            -Wait `
            -PassThru `
            -RedirectStandardOutput $runtimeOutLog `
            -RedirectStandardError $runtimeErrorLog
        "$(Get-Date -Format s) Bot process exited with code $($botProcess.ExitCode)." |
            Out-File -LiteralPath $runtimeLog -Encoding utf8 -Append
    }
    finally {
        $env:TELEGRAM_BOT_TOKEN = $null
        if ($tokenPtr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($tokenPtr)
        }
    }
}
catch {
    $errorMessage = "$(Get-Date -Format s) Startup failed: $($_.Exception.Message)"
    $errorMessage | Tee-Object -FilePath (Join-Path $PSScriptRoot "bot_startup_error.log") -Append
    $errorMessage | Out-File -LiteralPath $runtimeLog -Encoding utf8 -Append
    Write-Host "Keep this window open and send me a screenshot of the error above." -ForegroundColor Red
}

Write-Host "The bot is not running." -ForegroundColor Yellow
Read-Host "Press Enter to close"
