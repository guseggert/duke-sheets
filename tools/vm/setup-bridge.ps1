# PowerShell script to set up the Excel COM bridge server inside the Windows VM.
#
# IMPORTANT: The bridge must run in the interactive desktop session (Session 1),
# NOT from WinRM (Session 0). Excel COM automation requires an interactive
# desktop - it will crash with APPCRASH if launched from a non-interactive session.
#
# This script creates a scheduled task with LogonType Interactive to ensure
# the bridge always runs in the correct session.
#
# Run this once after installing Windows and Office. Can be run via WinRM:
#   python3 /tmp/winrm-exec.py -ps "$(cat tools/vm/setup-bridge.ps1)"

param(
    [string]$BridgePath = "C:\tools\ExcelBridgeServer.exe",
    [int]$Port = 9876
)

$ErrorActionPreference = "Stop"

Write-Host "=== Excel COM Bridge Server Setup ===" -ForegroundColor Cyan

# 1. Create tools directory and copy bridge exe from SMB share
$toolsDir = Split-Path $BridgePath -Parent
if (-not (Test-Path $toolsDir)) {
    New-Item -ItemType Directory -Path $toolsDir -Force | Out-Null
    Write-Host "Created $toolsDir"
}

# Try to copy from QEMU SMB share if not already present
$smbPath = "\\10.0.2.4\qemu\ExcelBridgeServer.exe"
if (-not (Test-Path $BridgePath)) {
    if (Test-Path $smbPath) {
        Copy-Item $smbPath $BridgePath -Force
        Write-Host "Copied bridge exe from SMB share to $BridgePath"
    } else {
        Write-Host "WARNING: Bridge exe not found at $BridgePath or $smbPath" -ForegroundColor Yellow
        Write-Host "Place ExcelBridgeServer.exe in /tmp/duke-sheets-excel/ on the Linux host"
        Write-Host "and it will be accessible at $smbPath via QEMU SMB."
    }
}

# 2. Disable Windows Firewall entirely (simplest for test VM)
#    Individual firewall rules are unreliable - Windows may still block connections
#    even with rules added, depending on network profile detection.
Write-Host "Disabling Windows Firewall (test VM only)..."
Set-NetFirewallProfile -All -Enabled False
Write-Host "  Firewall disabled on all profiles"

# 3. Create a scheduled task with LogonType Interactive
#    This is CRITICAL: Excel COM needs the interactive desktop (Session 1).
#    A normal scheduled task or WinRM-launched process runs in Session 0,
#    where COM activation of Excel.Application will crash.
Write-Host "Creating scheduled task with Interactive logon..."
$taskName = "ExcelBridgeServer"

# Remove existing task if present
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "  Removed existing task"
}

$action = New-ScheduledTaskAction -Execute $BridgePath -Argument "--port $Port"
$trigger = New-ScheduledTaskTrigger -AtLogOn -User "user"
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -ExecutionTimeLimit (New-TimeSpan -Days 365)

# LogonType Interactive ensures the task runs in the user's desktop session
$principal = New-ScheduledTaskPrincipal -UserId "user" -LogonType Interactive -RunLevel Highest

Register-ScheduledTask -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Force | Out-Null

Write-Host "  Scheduled task '$taskName' created (LogonType Interactive, runs at login)"

# 4. Start the bridge now (if user is logged in interactively)
Write-Host "Starting bridge server..."
Start-ScheduledTask -TaskName $taskName
Start-Sleep -Seconds 2

$task = Get-ScheduledTask -TaskName $taskName
if ($task.State -eq "Running") {
    Write-Host "  Bridge server is running!" -ForegroundColor Green
} else {
    Write-Host "  Task state: $($task.State)" -ForegroundColor Yellow
    Write-Host "  The bridge will start automatically at next interactive login."
    Write-Host "  If running via WinRM, this is expected - the task needs Session 1."
}

Write-Host ""
Write-Host "=== Setup complete ===" -ForegroundColor Green
Write-Host ""
Write-Host "Verify from Linux host:"
Write-Host "  echo '{\"id\":1,\"command\":\"Init\",\"data\":{\"prog_id\":\"Excel.Application\"}}' | nc -q1 localhost $Port"
Write-Host ""
Write-Host "If running via WinRM and the bridge didn't start, restart the VM"
Write-Host "or have the user log in interactively (the AtLogOn trigger will fire)."
