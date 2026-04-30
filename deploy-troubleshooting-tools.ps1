#Requires -Version 5.1
#Requires -RunAsAdministrator
<#PSScriptInfo
.VERSION 2.1.0
.DESCRIPTION Downloads and deploys Autopilot ESP troubleshooting toolkit.
.TAGS Intune Autopilot
.RELEASENOTES
  2.1.0: Adds startup persistence via scheduled task. Tool relaunches after
         reboots during provisioning, but self-removes after reseal (Sysprep
         clears enrollment keys) or when ESP completes.
  2.0.0: Rewritten to support new tools.ps1 (DataGridView + toolbar layout).
         Script dropdown now auto-discovers any .ps1 files in the tools folder.
         NuGet + Get-AutopilotDiagnostics still installed for dropdown availability.
#>
<#
.SYNOPSIS
  Downloads and deploys the Autopilot ESP troubleshooting toolkit.
.DESCRIPTION
  Runs as a SYSTEM-context Intune script during Autopilot. Drops all required
  binaries and scripts into C:\ProgramData\ServiceUI, then uses ServiceUI.exe
  to surface the tools GUI in the interactive user session during OOBE/ESP.
  Exits silently if provisioning is already complete.
.NOTES
  Set $ToolsBaseUrl below to the raw-file URL of your own hosting location
  (GitHub raw, Azure Blob, SharePoint direct-download, etc.).
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Downloads ─────────────────────────────────────────────────────────────────
# Each entry is  'filename.ext' = 'https://...'
# The file is saved to $ToolsFolder\filename.ext.
# Every URL is independent — mix GitHub raw, Azure Blob, SharePoint, etc.
# Any .ps1 added here (other than tools.ps1 and shiftf10.ps1) will automatically
# appear in the Script dropdown inside the tools UI.
$Downloads = [ordered]@{

    # ── Required binaries ──────────────────────────────────────────────────────
    'serviceui.exe' = 'https://github.com/andrew-s-taylor/public/raw/main/Troubleshooting/ServiceUI.exe'
    'cmtrace.exe'   = 'https://github.com/andrew-s-taylor/public/raw/main/Troubleshooting/CMTrace.exe'

    # ── Main UI ────────────────────────────────────────────────────────────────
    'tools.ps1'     = 'https://raw.githubusercontent.com/machacker24/ESPStatusTool/main/tools.ps1'

    # ── Boot-persistence guard (runs at startup, self-removes when done) ──────
    'startup.ps1'   = 'https://raw.githubusercontent.com/machacker24/ESPStatusTool/main/startup.ps1'

    # ── Optional diagnostic scripts (each appears in the Script dropdown) ────────
    'Get-AutopilotDiagnosticsCommunity.ps1'        = 'https://raw.githubusercontent.com/machacker24/ESPStatusTool/main/Get-AutopilotDiagnosticsCommunity.ps1'
    'Get-IntuneManagementExtensionDiagnostics.ps1' = 'https://raw.githubusercontent.com/machacker24/ESPStatusTool/main/Get-IntuneManagementExtensionDiagnostics.ps1'
    # Add further custom tools here — one line per file:
    # 'My-CustomTool.ps1' = 'https://...'
}
# ──────────────────────────────────────────────────────────────────────────────

$ToolsFolder = 'C:\ProgramData\ServiceUI'
$LogFile     = Join-Path $ToolsFolder 'deploy.log'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')][$Level] $Message"
    Write-Output $line                                      # Intune portal output
    Add-Content -Path $LogFile -Value $line -ErrorAction SilentlyContinue  # local disk
}

# ── Ensure log folder exists before first write ───────────────────────────────
if (-not (Test-Path $ToolsFolder)) {
    New-Item -Path $ToolsFolder -ItemType Directory -Force | Out-Null
}
Write-Log "deploy-troubleshooting-tools.ps1 started (v2.0.0)"

# ── Guard: skip if provisioning already finished ──────────────────────────────
$intunePath     = 'HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps'
$intuneAppCount = if (Test-Path $intunePath) { @(Get-ChildItem $intunePath -ErrorAction SilentlyContinue).Count } else { 0 }
if ($intuneAppCount -ge 2) {
    Write-Log "Provisioning appears complete ($intuneAppCount Win32App entries). Exiting." 'WARN'
    exit 0
}

Write-Log "Tools folder: $ToolsFolder"

# ── Execution policy ──────────────────────────────────────────────────────────
try {
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force
    Write-Log "Execution policy set to RemoteSigned"
} catch {
    Write-Log "Set-ExecutionPolicy failed (may be GPO-controlled): $($_.Exception.Message)" 'WARN'
}

# ── NuGet + Get-AutopilotDiagnostics (makes it available in the script dropdown) ──
try {
    Write-Log "Installing NuGet provider..."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    Write-Log "NuGet installed."
} catch {
    Write-Log "NuGet install failed: $($_.Exception.Message)" 'WARN'
}

try {
    Write-Log "Installing Get-AutopilotDiagnostics from PSGallery..."
    Install-Script -Name Get-AutopilotDiagnostics -Force
    Write-Log "Get-AutopilotDiagnostics installed."
} catch {
    Write-Log "Get-AutopilotDiagnostics install failed: $($_.Exception.Message)" 'WARN'
}

# ── Download files ────────────────────────────────────────────────────────────
$webArgs = @{
    UseBasicParsing = $true
    Headers         = @{ 'Cache-Control' = 'no-cache' }
}

foreach ($fileName in $Downloads.Keys) {
    $dest = Join-Path $ToolsFolder $fileName
    $url  = $Downloads[$fileName]
    try {
        Write-Log "Downloading $fileName from $url"
        Invoke-WebRequest -Uri $url -OutFile $dest @webArgs
        Write-Log "Saved $fileName ($([Math]::Round((Get-Item $dest).Length / 1KB, 1)) KB)"
    } catch {
        Write-Log "Failed to download $fileName`: $($_.Exception.Message)" 'ERROR'
    }
}

# ── Create session-bridge script (shiftf10.ps1) ───────────────────────────────
# ServiceUI runs this in the user session. It sends Alt+Tab then Shift+F10 to
# provoke a shell in the OOBE UI, kills it immediately, then launches tools.ps1.
$shiftf10Content = @'
$shell = New-Object -ComObject WScript.Shell
$shell.SendKeys('%({TAB})')
Start-Sleep -Seconds 1
$shell.SendKeys('+({F10})')

$waited = 0
Do {
    Start-Sleep -Seconds 1
    $waited++
} While (-not (Get-Process cmd -ErrorAction SilentlyContinue) -and $waited -lt 15)

if (Get-Process cmd -ErrorAction SilentlyContinue) {
    Get-Process cmd | Stop-Process -Force
}

Start-Process powershell.exe -ArgumentList '-NoLogo -NoProfile -NoExit -ExecutionPolicy Bypass -File C:\ProgramData\ServiceUI\tools.ps1' -Wait
'@

$shiftf10Path = Join-Path $ToolsFolder 'shiftf10.ps1'
$shiftf10Content | Out-File -FilePath $shiftf10Path -Encoding utf8 -Force
Write-Log "Created shiftf10.ps1"

# ── Save enrollment session ID for reseal detection ──────────────────────────
# After reseal, Sysprep clears HKLM:\SOFTWARE\Microsoft\Enrollments.
# startup.ps1 checks these GUIDs on each boot; if none remain, the device
# was resealed and the task removes itself rather than relaunching the tool.
$enrollBase = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
$enrollIds  = @()
if (Test-Path $enrollBase) {
    $enrollIds = @(Get-ChildItem $enrollBase -ErrorAction SilentlyContinue |
                   Where-Object { $_.SubKeyCount -gt 0 } |
                   Select-Object -ExpandProperty PSChildName)
}
if ($enrollIds.Count -gt 0) {
    $enrollIds -join ',' | Out-File (Join-Path $ToolsFolder 'session.id') -Encoding utf8 -Force
    Write-Log "Session ID saved ($($enrollIds.Count) enrollment(s))"
} else {
    Write-Log "No enrollment keys found — session.id not written" 'WARN'
}

# ── Register scheduled task for boot persistence ──────────────────────────────
$startupPath = Join-Path $ToolsFolder 'startup.ps1'
try {
    $action    = New-ScheduledTaskAction `
                     -Execute  'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' `
                     -Argument "-NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$startupPath`""
    $trigger   = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet `
                     -AllowStartIfOnBatteries `
                     -DontStopIfGoingOnBatteries `
                     -StartWhenAvailable `
                     -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

    Register-ScheduledTask `
        -TaskName 'AutopilotESPTools' `
        -Action    $action `
        -Trigger   $trigger `
        -Principal $principal `
        -Settings  $settings `
        -Force | Out-Null

    Write-Log "Scheduled task 'AutopilotESPTools' registered for startup persistence."
} catch {
    Write-Log "Failed to register scheduled task: $($_.Exception.Message)" 'WARN'
}

# ── Launch via ServiceUI ──────────────────────────────────────────────────────
$serviceUI = Join-Path $ToolsFolder 'serviceui.exe'
$psExe     = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$psArgs    = "-ExecutionPolicy Bypass -File `"$shiftf10Path`" -WindowStyle Hidden"

if (-not (Test-Path $serviceUI)) {
    Write-Log "serviceui.exe not found at $serviceUI — cannot launch GUI." 'ERROR'
    exit 1
}

$toolsPs1 = Join-Path $ToolsFolder 'tools.ps1'
if (-not (Test-Path $toolsPs1)) {
    Write-Log "tools.ps1 not found at $toolsPs1 — GUI cannot start." 'ERROR'
    exit 1
}

Write-Log "Launching tools GUI via ServiceUI..."
Start-Process $serviceUI -ArgumentList @('-process:explorer.exe', "$psExe $psArgs")
Write-Log "Deploy complete."
