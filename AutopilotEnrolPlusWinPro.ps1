# Elevate Execution Policy
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# 1. Check Windows Version
$winver = (Get-WmiObject -Class Win32_OperatingSystem).Caption
Write-Host "Checking Windows version..." -ForegroundColor Yellow
if ($winver -like "*Windows 11*") {
    Write-Host "Windows version detected: $winver" -ForegroundColor Green
} elseif ($winver -like "*Windows 10*") {
    Write-Host "Windows 10 detected. Continuing with the process..." -ForegroundColor Yellow
} else {
    Write-Host "ERROR: Unsupported version of Windows. Exiting..." -ForegroundColor Red
    exit
}

# 2. Check for Home Edition and Force Upgrade to Pro
if ($winver -like "*Home*") {
    Write-Host "Windows Home detected. Forcing upgrade to Windows Pro..." -ForegroundColor Red

    # Clear Activation Data
    Write-Host "Clearing activation data..." -ForegroundColor Yellow
    Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c slmgr.vbs /upk"
    Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c slmgr.vbs /cpky"
    Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c slmgr.vbs /ckms"

    # Force the Upgrade to Windows Pro
    Write-Host "Starting upgrade to Windows Pro. This may take some time..." -ForegroundColor Yellow
    Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c sc config LicenseManager start= auto & net start LicenseManager"
    Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c sc config wuauserv start= auto & net start wuauserv"
    Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c changepk.exe /productkey VK7JG-NPHTM-C97JM-9MPGT-3V66T"

    Write-Host "Rebooting system to finalize upgrade. Please wait during updates..." -ForegroundColor Cyan
    shutdown /r /t 0
    exit
}

# 3. Activate Windows Pro
Write-Host "Activating Windows Pro..." -ForegroundColor Yellow
Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX"
Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c slmgr.vbs /skms kms8.msguides.com"
Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c slmgr.vbs /ato"

Write-Host "Windows Pro has been successfully activated!" -ForegroundColor Green

# 4. Prepare for Autopilot Enrollment
Write-Host "Preparing for Autopilot enrollment..." -ForegroundColor Yellow

# Install Get-WindowsAutopilotInfo if not installed
if (-not (Get-Command -Name Get-WindowsAutopilotInfo -ErrorAction SilentlyContinue)) {
    Write-Host "Installing Get-WindowsAutopilotInfo script..." -ForegroundColor Yellow
    Install-Script -Name Get-WindowsAutopilotInfo -Force -Scope CurrentUser
}

# Upload Device to Autopilot
Write-Host "Running Get-WindowsAutopilotInfo.ps1 to upload device to Autopilot..." -ForegroundColor Yellow
Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"
    Get-WindowsAutopilotInfo -Online;`"" -Wait

Write-Host "`n--- Manual Steps Required ---" -ForegroundColor Cyan
Write-Host "1. Sign in with Global Admin credentials when prompted."
Write-Host "2. After '1 device imported successfully', open Entra Admin Center."
Write-Host "3. Navigate to 'Devices > All Devices', search for the serial number."
Write-Host "4. Ensure 'Enabled = Yes'. If not, click 'Enable'."
Write-Host "Press Enter to continue once manual steps are complete." -ForegroundColor Magenta
Read-Host

# 5. Restart the Laptop
Write-Host "Restarting the laptop now..." -ForegroundColor Green
shutdown /r /t 0
