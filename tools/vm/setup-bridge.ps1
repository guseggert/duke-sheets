# PowerShell script to set up the Excel COM bridge server inside the Windows VM.
#
# Run this once after installing Windows and Office.
# Can be run via WinRM from the Linux host:
#   python3 -c "import winrm; s = winrm.Session('localhost:5985', auth=('user','pass'), transport='ntlm'); print(s.run_ps(open('tools/vm/setup-bridge.ps1').read()).std_out.decode())"

param(
    [string]$BridgePath = "C:\tools\ExcelBridgeServer.exe",
    [int]$Port = 9876
)

$ErrorActionPreference = "Stop"

Write-Host "=== Excel COM Bridge Server Setup ===" -ForegroundColor Cyan

# 1. Create tools directory
$toolsDir = Split-Path $BridgePath -Parent
if (-not (Test-Path $toolsDir)) {
    New-Item -ItemType Directory -Path $toolsDir -Force | Out-Null
    Write-Host "Created $toolsDir"
}

# 2. Check if bridge exe exists
if (-not (Test-Path $BridgePath)) {
    Write-Host "WARNING: Bridge executable not found at $BridgePath" -ForegroundColor Yellow
    Write-Host "Copy ExcelBridgeServer.exe to $BridgePath before continuing."
    Write-Host ""
    Write-Host "From the Linux host (via SMB share):"
    Write-Host "  The exe should be available at \\10.0.2.4\qemu\ExcelBridgeServer.exe"
    Write-Host "  Copy it to $BridgePath"
}

# 3. Configure Windows Firewall to allow bridge port
Write-Host "Configuring firewall rule for port $Port..."
$ruleName = "ExcelBridgeServer"
$existing = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
if ($existing) {
    Remove-NetFirewallRule -DisplayName $ruleName
}
New-NetFirewallRule -DisplayName $ruleName `
    -Direction Inbound -Protocol TCP -LocalPort $Port `
    -Action Allow -Profile Any | Out-Null
Write-Host "  Firewall rule created: allow TCP $Port inbound"

# 4. Create a scheduled task to auto-start the bridge on login
Write-Host "Creating auto-start scheduled task..."
$taskName = "ExcelBridgeServer"
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

$action = New-ScheduledTaskAction -Execute $BridgePath -Argument "--port $Port"
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
    -Settings $settings -RunLevel Highest -Force | Out-Null
Write-Host "  Scheduled task '$taskName' created (runs at login)"

# 5. Enable WinRM for remote management (if not already enabled)
Write-Host "Ensuring WinRM is configured..."
try {
    $winrm = Get-Service WinRM -ErrorAction Stop
    if ($winrm.Status -ne "Running") {
        winrm quickconfig -force 2>&1 | Out-Null
    }
    winrm set winrm/config/service/auth '@{Basic="true"}' 2>&1 | Out-Null
    winrm set winrm/config/service '@{AllowUnencrypted="true"}' 2>&1 | Out-Null
    Write-Host "  WinRM configured (Basic auth, unencrypted for dev)"
} catch {
    Write-Host "  WinRM setup skipped: $($_.Exception.Message)" -ForegroundColor Yellow
}

# 6. Enable OpenSSH Server (optional, for convenience)
Write-Host "Checking OpenSSH Server..."
$sshCapability = Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH.Server*'
if ($sshCapability -and $sshCapability.State -ne "Installed") {
    Write-Host "  Installing OpenSSH Server..."
    Add-WindowsCapability -Online -Name $sshCapability.Name | Out-Null
    Start-Service sshd
    Set-Service -Name sshd -StartupType Automatic
    Write-Host "  OpenSSH Server installed and started"
} elseif ($sshCapability) {
    Write-Host "  OpenSSH Server already installed"
} else {
    Write-Host "  OpenSSH Server capability not available" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Setup complete ===" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Copy ExcelBridgeServer.exe to $BridgePath"
Write-Host "  2. Restart the VM (or run: Start-ScheduledTask -TaskName '$taskName')"
Write-Host "  3. From Linux host: nc -z localhost $Port  (should succeed)"
