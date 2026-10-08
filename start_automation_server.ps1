$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not $env:BARCO_API_TOKEN) {
    $env:BARCO_API_TOKEN = [Environment]::GetEnvironmentVariable("BARCO_API_TOKEN", "User")
}
if (-not $env:BARCO_API_TOKEN) {
    Write-Error "Set BARCO_API_TOKEN before starting the server."
}
if (-not $env:BARCO_TAILSCALE_KEEPALIVE_IP) {
    $env:BARCO_TAILSCALE_KEEPALIVE_IP = [Environment]::GetEnvironmentVariable(
        "BARCO_TAILSCALE_KEEPALIVE_IP", "User"
    )
}

$listener = Get-NetTCPConnection -LocalPort 8080 -State Listen -ErrorAction SilentlyContinue |
    Select-Object -First 1
if ($listener) {
    Write-Host "Port 8080 is already in use (PID $($listener.OwningProcess))."
    Write-Host "If the Barco API is already running, use its existing window."
    exit 0
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

$tailscaleIp = $null
for ($attempt = 1; $attempt -le 24; $attempt++) {
    try {
        $address = & $tailscaleExe ip -4 2>$null | Select-Object -First 1
        if ($address) {
            $tailscaleIp = $address.Trim()
            break
        }
    } catch {
        # Tailscale may still be starting during Windows sign-in.
    }

    if ($attempt -lt 24) {
        if ($attempt -eq 1) {
            Write-Host "Waiting for Tailscale to connect..."
        }
        Start-Sleep -Seconds 5
    }
}
if (-not $tailscaleIp) {
    Write-Error "Tailscale IPv4 address was not found after 2 minutes. Check Tailscale."
}

Write-Host "Barco automation API: http://${tailscaleIp}:8080"
$keepaliveJob = $null
if ($env:BARCO_TAILSCALE_KEEPALIVE_IP) {
    $keepaliveIp = $env:BARCO_TAILSCALE_KEEPALIVE_IP
    if ($keepaliveIp -notmatch '^100\.(6[4-9]|[7-9][0-9]|1[01][0-9]|12[0-7])\.(\d{1,3})\.(\d{1,3})$') {
        Write-Error "BARCO_TAILSCALE_KEEPALIVE_IP must be a Tailscale IPv4 address."
    }

    $keepaliveJob = Start-Job -ArgumentList $tailscaleExe, $keepaliveIp -ScriptBlock {
        param($tailscalePath, $peerIp)
        while ($true) {
            & $tailscalePath ping --timeout 5s --c 1 $peerIp *> $null
            Start-Sleep -Seconds 60
        }
    }
    Write-Host "Tailscale keepalive enabled for $keepaliveIp (every 60 seconds)."
}

# Listen on every local interface so Windows accepts connections arriving
# through the Tailscale adapter as well as local health checks.
try {
    py -m uvicorn automation_server:app --host 0.0.0.0 --port 8080
} finally {
    if ($keepaliveJob) {
        $keepaliveJob | Stop-Job
        $keepaliveJob | Remove-Job
    }
}
