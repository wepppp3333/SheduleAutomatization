$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$requiredVariables = @(
    "TELEGRAM_BOT_TOKEN",
    "LUKOYANOV_API_TOKEN"
)

foreach ($variable in $requiredVariables) {
    if (-not (Get-Item "Env:$variable" -ErrorAction SilentlyContinue).Value) {
        Write-Error "Set $variable before starting the Telegram bot."
    }
}

py telegram_bot.py
