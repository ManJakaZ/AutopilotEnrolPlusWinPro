# Elevate Execution Policy
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Check Windows Version
$winver = (Get-WmiObject -Class Win32_OperatingSystem).Caption
if ($winver -notlike "*Professional*") {
    Write-Host "ERROR: The laptop is not running Windows 10/11 Pro. Exiting..." -ForegroundColor Red
    exit
}
Write-Host "Windows version check passed: $winver" -ForegroundColor Green

# Install Windows Autopilot Info Script
if (-not (Get-Command -Name Get-WindowsAutopilotInfo -ErrorAction SilentlyContinue)) {
    Write-Host "Installing Get-WindowsAutopilotInfo script..." -ForegroundColor Yellow
    Install-Script -Name Get-WindowsAutopilotInfo -Force -Scope CurrentUser
}

# Upload Device to Autopilot
Write-Host "Running Get-WindowsAutopilotInfo.ps1 to upload device to Autopilot..." -ForegroundColor Yellow
Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"
    Get-WindowsAutopilotInfo -Online;`"" -Wait

# Prompt for Manual Steps
Write-Host "`n--- Manual Steps Required ---" -ForegroundColor Cyan
Write-Host "1. Sign in with Global Admin credentials when prompted."
Write-Host "2. After '1 device imported successfully', open Entra Admin Center."
Write-Host "3. Navigate to 'Devices > All Devices', search for the serial number."
Write-Host "4. Ensure 'Enabled = Yes'. If not, click 'Enable'."
Write-Host "Press Enter to continue once manual steps are complete." -ForegroundColor Magenta
Read-Host

# Restart the Laptop
Write-Host "Restarting the laptop now..." -ForegroundColor Green
shutdown /r /t 0
