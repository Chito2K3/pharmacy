<#
.SYNOPSIS
    Registers a Windows Scheduled Task for ARMMC Pharmacy IMS Live Sync.
.DESCRIPTION
    Creates a recurring Windows Scheduled Task that runs tools\sync_from_google_sheet.py
    automatically on a configurable daily schedule (default: 6:00 AM) or interval.
.PARAMETER Uninstall
    Removes the scheduled task if present.
.PARAMETER Time
    Daily start time for synchronization (default: "06:00").
#>

param (
    [switch]$Uninstall,
    [string]$Time = "06:00"
)

$TaskName = "ARMMC_Pharmacy_IMS_Sync"
$WorkspaceDir = (Get-Item $PSScriptRoot).Parent.FullName
$PythonExe = (Get-Command python -ErrorAction SilentlyContinue)?.Source
$ScriptPath = Join-Path $PSScriptRoot "sync_from_google_sheet.py"

if (-not $PythonExe) {
    Write-Error "Python executable not found in PATH. Please install Python 3 or add it to system PATH."
    exit 1
}

if ($Uninstall) {
    Write-Host "Unregistering scheduled task '$TaskName'..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "[OK] Task '$TaskName' removed successfully." -ForegroundColor Green
    exit 0
}

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " ARMMC Pharmacy IMS -- Windows Scheduled Task Installer" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Workspace Dir: $WorkspaceDir"
Write-Host "Python Path:   $PythonExe"
Write-Host "Sync Script:   $ScriptPath"
Write-Host "Daily Schedule Time: $Time"
Write-Host ""

$Action = New-ScheduledTaskAction -Execute $PythonExe -Argument "`"$ScriptPath`"" -WorkingDirectory $WorkspaceDir
$Trigger = New-ScheduledTaskTrigger -Daily -At $Time
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

try {
    # Remove existing task if already present
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Description "Automated Google Sheet sync for ARMMC Pharmacy Inventory Monitoring System" | Out-Null
    Write-Host "[OK] Successfully registered scheduled task '$TaskName' to run daily at $Time!" -ForegroundColor Green
    Write-Host "     You can test-run it anytime in PowerShell with:" -ForegroundColor Gray
    Write-Host "     Start-ScheduledTask -TaskName '$TaskName'" -ForegroundColor White
} catch {
    Write-Error "Failed to register scheduled task: $_"
    exit 1
}
