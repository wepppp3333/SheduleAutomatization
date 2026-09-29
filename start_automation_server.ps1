$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not $env:BARCO_API_TOKEN) {
    Write-Error "Set BARCO_API_TOKEN before starting the server."
}

$tailscaleIp = (& tailscale ip -4 | Select-Object -First 1).Trim()
if (-not $tailscaleIp) {
    Write-Error "Tailscale IPv4 address was not found. Start Tailscale first."
}

Write-Host "Barco automation API: http://${tailscaleIp}:8080"
py -m uvicorn automation_server:app --host $tailscaleIp --port 8080
