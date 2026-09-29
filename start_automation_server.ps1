$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not $env:BARCO_API_TOKEN) {
    Write-Error "Set BARCO_API_TOKEN before starting the server."
}

$tailscaleCommand = Get-Command "tailscale.exe" -ErrorAction SilentlyContinue
$tailscaleExe = if ($tailscaleCommand) {
    $tailscaleCommand.Source
} else {
    "C:\Program Files\Tailscale\tailscale.exe"
}

if (-not (Test-Path $tailscaleExe)) {
    Write-Error "tailscale.exe was not found. Install or start Tailscale first."
}

$tailscaleIp = (& $tailscaleExe ip -4 | Select-Object -First 1).Trim()
if (-not $tailscaleIp) {
    Write-Error "Tailscale IPv4 address was not found. Start Tailscale first."
}

Write-Host "Barco automation API: http://${tailscaleIp}:8080"
py -m uvicorn automation_server:app --host $tailscaleIp --port 8080
