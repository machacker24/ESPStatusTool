#Requires -Version 5.1
<#
.SYNOPSIS
  Boot-persistence guard for Autopilot ESP troubleshooting toolkit.
.DESCRIPTION
  Registered as a scheduled task (AutopilotESPTools) by deploy-troubleshooting-tools.ps1.
  Runs at every startup, checks whether the tool should relaunch, and self-removes
  the task when provisioning is complete or the device has been resealed.
.NOTES
  Guard 1 - Reseal: Sysprep clears HKLM:\...\Enrollments. If the GUIDs saved
            at deploy time are gone, the device was resealed. Task self-removes.
  Guard 2 - Complete: Win32App count >= 2 means ESP finished. Task self-removes.
  Guard 3 - Profile: AutopilotDDSZTDFile.json absent = not an Autopilot device.
  Otherwise waits up to 3 min for CloudExperienceHost then launches via ServiceUI.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

$ToolsFolder = 'C:\ProgramData\ServiceUI'
$LogFile     = Join-Path $ToolsFolder 'startup.log'
$TaskName    = 'AutopilotESPTools'

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts][$Level] $Message"
    Add-Content -Path $LogFile -Value $line -ErrorAction SilentlyContinue
}

function Remove-SelfAndExit {
    param([string]$Reason)
    Write-Log "Removing startup task: $Reason"
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    exit 0
}

Write-Log ("startup.ps1 triggered at " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))

# Guard 1: Reseal detection
# Sysprep (reseal) clears HKLM:\SOFTWARE\Microsoft\Enrollments.
# If none of the enrollment GUIDs saved by the deploy script still exist,
# the device has been resealed and we must not relaunch the tool.
$sessionFile = Join-Path $ToolsFolder 'session.id'
if (Test-Path $sessionFile) {
    $storedIds = ((Get-Content $sessionFile -Raw -ErrorAction SilentlyContinue) -split ',') |
                 ForEach-Object { $_.Trim() } |
                 Where-Object   { $_ -ne '' }

    if ($storedIds.Count -gt 0) {
        $anyExists = $false
        foreach ($id in $storedIds) {
            if (Test-Path "HKLM:\SOFTWARE\Microsoft\Enrollments\$id") {
                $anyExists = $true
                break
            }
        }
        if (-not $anyExists) {
            Remove-SelfAndExit 'Reseal detected: stored enrollment GUIDs no longer present'
        }
    }
}

# Guard 2: Provisioning complete
$intunePath = 'HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps'
$appCount   = @(Get-ChildItem $intunePath -ErrorAction SilentlyContinue).Count
if ($appCount -ge 2) {
    Remove-SelfAndExit "Provisioning complete ($appCount Win32App entries)"
}

# Guard 3: Autopilot profile must be present
$autopilotJson = Join-Path $env:WINDIR 'ServiceState\wmansvc\AutopilotDDSZTDFile.json'
if (-not (Test-Path $autopilotJson)) {
    Remove-SelfAndExit 'AutopilotDDSZTDFile.json absent - not an Autopilot device'
}

# Wait for CloudExperienceHost (OOBE/ESP UI) to be running - up to 3 minutes
Write-Log 'Waiting for CloudExperienceHost...'
$waited = 0
while ($waited -lt 180) {
    if (Get-Process -Name 'CloudExperienceHost' -ErrorAction SilentlyContinue) { break }
    Start-Sleep -Seconds 10
    $waited += 10
}

if (-not (Get-Process -Name 'CloudExperienceHost' -ErrorAction SilentlyContinue)) {
    Write-Log "CloudExperienceHost not detected after ${waited}s - not in OOBE, skipping launch" 'WARN'
    exit 0
}

# Brief settle time so ServiceUI can find the interactive session
Start-Sleep -Seconds 15

# Launch tools via ServiceUI
$serviceUI = Join-Path $ToolsFolder 'serviceui.exe'
$psExe     = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$shiftf10  = Join-Path $ToolsFolder 'shiftf10.ps1'

if (-not (Test-Path $serviceUI)) { Write-Log 'serviceui.exe not found' 'ERROR'; exit 1 }
if (-not (Test-Path $shiftf10))  { Write-Log 'shiftf10.ps1 not found'  'ERROR'; exit 1 }

$psArgs = '-NoLogo -NoProfile -ExecutionPolicy Bypass -File "' + $shiftf10 + '" -WindowStyle Hidden'
Write-Log 'Launching tools via ServiceUI...'
Start-Process $serviceUI -ArgumentList @('-process:explorer.exe', "$psExe $psArgs")
Write-Log 'Launch complete.'
