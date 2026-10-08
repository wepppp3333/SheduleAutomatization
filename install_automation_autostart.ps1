$ErrorActionPreference = "Stop"

$launcher = Join-Path $PSScriptRoot "Start Barco API.cmd"
if (-not (Test-Path $launcher)) {
    throw "Start Barco API.cmd was not found in the project directory."
}

$startup = [Environment]::GetFolderPath([Environment+SpecialFolder]::Startup)
if (-not $startup -or -not (Test-Path $startup)) {
    throw "Windows Startup folder was not found for the current user."
}

$shortcutPath = Join-Path $startup "Barco Automation API.lnk"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)

if ((Test-Path $shortcutPath) -and $shortcut.TargetPath -ne $launcher) {
    throw "A different Barco Automation API shortcut already exists in Startup: $shortcutPath"
}

$shortcut.TargetPath = $launcher
$shortcut.WorkingDirectory = $PSScriptRoot
$shortcut.Description = "Start the Barco automation API after Windows sign-in"
$shortcut.Save()

Write-Host "Autostart installed for the current Windows user: $shortcutPath"
Write-Host "The API will start after the next sign-in. To start it now, run Start Barco API.cmd."
