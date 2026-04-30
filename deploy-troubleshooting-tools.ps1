#Requires -Version 5.1
#Requires -RunAsAdministrator
<#PSScriptInfo
.VERSION 2.0.0
.DESCRIPTION Downloads and deploys Autopilot ESP troubleshooting toolkit.
.TAGS Intune Autopilot
.RELEASENOTES
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
    'tools.ps1'     = 'https://raw.githubusercontent.com/YOUR-ORG/YOUR-REPO/main/ESPStatusTool/tools.ps1'

    # ── Optional diagnostic scripts (uncomment and set URLs to enable) ─────────
    # Each will appear as a selectable entry in the Script dropdown.
    #
    # 'Get-AutopilotDiagnosticsCommunity.ps1'         = 'https://raw.githubusercontent.com/YOUR-ORG/YOUR-REPO/main/scripts/Get-AutopilotDiagnosticsCommunity.ps1'
    # 'Get-IntuneManagementExtensionDiagnostics.ps1'  = 'https://raw.githubusercontent.com/YOUR-ORG/YOUR-REPO/main/scripts/Get-IntuneManagementExtensionDiagnostics.ps1'
    # 'My-CustomTool.ps1'                             = 'https://mystorageaccount.blob.core.windows.net/scripts/My-CustomTool.ps1'
}
# ──────────────────────────────────────────────────────────────────────────────

$ToolsFolder = 'C:\ProgramData\ServiceUI'

function Write-Status([string]$Message) {
    Write-Output "[$(Get-Date -Format 'HH:mm:ss')] $Message"
}

# ── Guard: skip if provisioning already finished ──────────────────────────────
$intunePath     = 'HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps'
$intuneAppCount = if (Test-Path $intunePath) { @(Get-ChildItem $intunePath -ErrorAction SilentlyContinue).Count } else { 0 }
if ($intuneAppCount -ge 2) {
    Write-Status "Provisioning appears complete ($intuneAppCount Win32App entries). Exiting."
    exit 0
}

# ── Tools folder ──────────────────────────────────────────────────────────────
if (-not (Test-Path $ToolsFolder)) {
    New-Item -Path $ToolsFolder -ItemType Directory -Force | Out-Null
    Write-Status "Created $ToolsFolder"
} else {
    Write-Status "$ToolsFolder already exists."
}

# ── Execution policy ──────────────────────────────────────────────────────────
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force

# ── NuGet + Get-AutopilotDiagnostics (makes it available in the script dropdown) ──
Write-Status "Installing NuGet provider..."
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null

Write-Status "Installing Get-AutopilotDiagnostics from PSGallery..."
Install-Script -Name Get-AutopilotDiagnostics -Force

# ── Download files ────────────────────────────────────────────────────────────
$webArgs = @{
    UseBasicParsing = $true
    Headers         = @{ 'Cache-Control' = 'no-cache' }
}

foreach ($fileName in $Downloads.Keys) {
    $dest = Join-Path $ToolsFolder $fileName
    $url  = $Downloads[$fileName]
    Write-Status "Downloading $fileName..."
    Invoke-WebRequest -Uri $url -OutFile $dest @webArgs
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

$shiftf10Content | Out-File -FilePath (Join-Path $ToolsFolder 'shiftf10.ps1') -Encoding utf8 -Force
Write-Status "Created shiftf10.ps1"

# ── Launch via ServiceUI ──────────────────────────────────────────────────────
$serviceUI  = Join-Path $ToolsFolder 'serviceui.exe'
$psExe      = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$psArgs     = "-ExecutionPolicy Bypass -File `"$(Join-Path $ToolsFolder 'shiftf10.ps1')`" -WindowStyle Hidden"

Write-Status "Launching tools GUI via ServiceUI..."
Start-Process $serviceUI -ArgumentList @(
    '-process:explorer.exe',
    "$psExe $psArgs"
)
