#Requires -Version 5.1
<#
.SYNOPSIS
    Autopilot ESP troubleshooting toolkit — GUI launcher + live status view.
.DESCRIPTION
    Displays ESP-tracked apps/policies (Name, GUID, Status) read from the local
    registry. Tool buttons across the top give quick access to diagnostic utilities.
    The Script dropdown discovers any .ps1 files deployed alongside this script.
.NOTES
    Designed to run during Autopilot OOBE via ServiceUI.exe.
    No internet or Graph access required for core functionality.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ── Constants ─────────────────────────────────────────────────────────────────
$Script:ToolsFolder = 'C:\ProgramData\ServiceUI'
$Script:CMTrace     = Join-Path $Script:ToolsFolder 'cmtrace.exe'
$Script:SetupActLog = 'C:\Windows\Panther\setupact.log'
$Script:IMELog      = 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log'

# Scripts to exclude from the dropdown (internal/helper scripts)
$Script:ExcludedScripts = @('tools.ps1', 'shiftf10.ps1')

$Script:AutoRefreshIntervalMs = 10000   # 10 seconds; change to taste
$Script:RefreshBusy           = $false  # guard against overlapping timer ticks

# ── State maps (sourced from Get-AutopilotDiagnostics logic) ─────────────────
# Win32 ComplianceStateMessage.ComplianceState (sub-key structure, newer IME)
$Script:ComplianceStates = @{
    0 = 'Unknown'
    1 = 'Installed'
    2 = 'Not Installed'
    3 = 'Conflict'
    4 = 'Failed'
}

# Win32 / Sidecar ProvisioningProgress WorkloadState
$Script:WorkloadStates = @{
    0 = 'Not Started'
    1 = 'Installed'
    2 = 'Skipped'
    3 = 'Uninstalled'
    4 = 'Failed'
    5 = 'Installing'
    6 = 'Pending Reboot'
    7 = 'Cancelled'
}

# ESP tracking InstallationState (EnrollmentStatusTracking)
$Script:ESPStates = @{
    1 = 'Not Installed'
    2 = 'Installing'
    3 = 'Installed'
    4 = 'Failed'
}

# MSI app state (EnterpriseDesktopAppManagement) — uses same extended codes as OfficeCSP
$Script:MSIStates = @{
    0  = 'None'
    1  = 'Not Installed'
    2  = 'Downloading'
    3  = 'Installing'
    10 = 'Initialized'
    20 = 'Downloading'
    25 = 'Pending Retry'
    30 = 'Download Failed'
    40 = 'Download Complete'
    48 = 'Pending User Session'
    50 = 'Installing'
    55 = 'Pending Retry'
    60 = 'Failed'
    70 = 'Installed'
}

# Office CSP state
$Script:OfficeStates = @{
    0  = 'None'
    10 = 'Initialized'
    20 = 'Downloading'
    25 = 'Pending Retry'
    30 = 'Download Failed'
    40 = 'Download Complete'
    48 = 'Pending User Session'
    50 = 'Installing'
    55 = 'Pending Retry'
    60 = 'Failed'
    70 = 'Installed'
}

function Resolve-State([hashtable]$Map, $State) {
    if ($null -eq $State) { return 'Unknown' }
    $key = [int]$State
    if ($Map.ContainsKey($key)) { return $Map[$key] }
    return "State $key"
}

# ── Helpers ───────────────────────────────────────────────────────────────────
function Get-StatusColor([string]$Status) {
    switch -Wildcard ($Status) {
        'Installed'        { return [System.Drawing.Color]::FromArgb(0, 150, 0) }
        'Failed'           { return [System.Drawing.Color]::FromArgb(200, 0, 0) }
        'Download Failed'  { return [System.Drawing.Color]::FromArgb(200, 0, 0) }
        'Installing'       { return [System.Drawing.Color]::FromArgb(0, 100, 200) }
        'Downloading'      { return [System.Drawing.Color]::FromArgb(0, 100, 200) }
        'Pending*'         { return [System.Drawing.Color]::FromArgb(180, 100, 0) }
        'Not Installed'    { return [System.Drawing.Color]::FromArgb(120, 120, 120) }
        'Not Started'      { return [System.Drawing.Color]::FromArgb(120, 120, 120) }
        'Skipped'          { return [System.Drawing.Color]::FromArgb(120, 120, 120) }
        'Not Applicable'   { return [System.Drawing.Color]::FromArgb(160, 140, 180) }
        default            { return [System.Drawing.Color]::Black }
    }
}

# ── Device / profile data functions ──────────────────────────────────────────
function Get-AutopilotProfile {
    $result = @{ DeviceName = $env:COMPUTERNAME; Tenant = '-'; Mode = '-' }

    $jsonPath = "$env:WINDIR\ServiceState\wmansvc\AutopilotDDSZTDFile.json"
    if (-not (Test-Path $jsonPath)) { return $result }

    try {
        $p = Get-Content $jsonPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

        if ($p.CloudAssignedDeviceName) {
            $deviceName = $p.CloudAssignedDeviceName
            # Resolve %SERIAL% template token if present
            if ($deviceName -like '*%SERIAL%*') {
                try {
                    $serial = (Get-WmiObject -Class Win32_BIOS -ErrorAction Stop).SerialNumber.Trim()
                    $deviceName = $deviceName -replace '%SERIAL%', $serial
                } catch { }
            }
            $result.DeviceName = $deviceName
        }

        $result.Tenant = if ($p.CloudAssignedTenantDomain) { $p.CloudAssignedTenantDomain }
                         elseif ($p.CloudAssignedTenantId)  { $p.CloudAssignedTenantId }
                         else                               { '-' }

        $result.Mode = if ($p.CloudAssignedDomainJoinMethod -eq 1) { 'Hybrid Azure AD' }
                       elseif ($p.CloudAssignedForcedEnrollment -eq 1) { 'Self-Deploying' }
                       else { 'Azure AD Join' }
    } catch { }

    return $result
}

function Get-AssignedUser {
    # MDM enrollment UPN — set after enrollment completes
    $enrollBase = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (Test-Path $enrollBase) {
        foreach ($key in Get-ChildItem $enrollBase -ErrorAction SilentlyContinue) {
            $upn = (Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue).UPN
            if ($upn) {
                # Strip AAD object-ID suffix: "user@domain@{guid}" -> "user@domain"
                if ($upn -match '^(.+@[^@]+)@[0-9a-fA-F\-]{36}') { $upn = $Matches[1] }
                return $upn
            }
        }
    }

    # AAD identity cache — populated when an AAD user has signed in
    $cacheBase = 'HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache'
    if (Test-Path $cacheBase) {
        foreach ($sid in Get-ChildItem $cacheBase -ErrorAction SilentlyContinue) {
            $idCache = Join-Path $sid.PSPath 'IdentityCache'
            if (-not (Test-Path $idCache)) { continue }
            foreach ($entry in Get-ChildItem $idCache -ErrorAction SilentlyContinue) {
                $name = (Get-ItemProperty $entry.PSPath -ErrorAction SilentlyContinue).UserName
                if ($name -like '*@*') { return $name }
            }
        }
    }

    return 'Not assigned'
}

function Get-BitLockerStatus {
    try {
        $vol = Get-WmiObject `
            -Namespace 'root\CIMV2\Security\MicrosoftVolumeEncryption' `
            -Class Win32_EncryptableVolume `
            -Filter "DriveLetter='C:'" `
            -ErrorAction Stop

        $prot  = $vol.GetProtectionStatus().protectionStatus
        $conv  = $vol.GetConversionStatus()
        $pct   = [int]$conv.encryptionPercentage
        $state = [int]$conv.conversionStatus

        $text = switch ($state) {
            0 { if ($prot -eq 1) { 'Pending Encryption' } else { 'Not Encrypted' } }
            1 { 'Encrypted' }
            2 { "Encrypting $pct%" }
            3 { "Decrypting $pct%" }
            4 { "Paused (Encrypting) $pct%" }
            5 { "Paused (Decrypting) $pct%" }
            default { 'Unknown' }
        }

        return @{ Text = $text; Percentage = $pct; Protected = ($prot -eq 1); State = $state }
    } catch {
        return @{ Text = 'Unavailable'; Percentage = 0; Protected = $false; State = -1 }
    }
}

# ── IME log name resolution ───────────────────────────────────────────────────
function Get-IMENameMap {
    # Scans IME log files for "Get policies = [...]" entries containing GUID->Name map.
    # App entries moved to AppWorkload.log in IME ~Aug 2024; scans both with wildcards.
    # Uses FileShare.ReadWrite so locked files (IME actively writing) can still be read.
    # Handles multi-line CMTrace entries by buffering lines between <![LOG[ and ]LOG]!>.
    $nameMap   = @{}
    $logFolder = 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs'
    if (-not (Test-Path $logFolder)) { return $nameMap }

    # AppWorkload first (newer IME), then IntuneManagementExtension (older)
    $logFiles  = @()
    $logFiles += Get-ChildItem "$logFolder\*AppWorkload*.log"       -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    $logFiles += Get-ChildItem "$logFolder\*IntuneManagement*.log"  -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    if (-not $logFiles) { return $nameMap }

    foreach ($logFile in $logFiles) {
        try {
            $fs     = [System.IO.FileStream]::new(
                          $logFile.FullName,
                          [System.IO.FileMode]::Open,
                          [System.IO.FileAccess]::Read,
                          [System.IO.FileShare]::ReadWrite)
            $reader = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::UTF8, $true, 65536)
            $buffer = $null   # accumulates multi-line CMTrace entries
            try {
                while ($null -ne ($line = $reader.ReadLine())) {
                    # Start of a CMTrace entry
                    if ($line -match '^\<\!\[LOG\[') {
                        $buffer = $line
                        # Single-line entry ends with ]LOG]!> on the same line
                        if ($line -match '\]LOG\]\!>') {
                            # Extract message: between <![LOG[ and ]LOG]
                            $msg = ($line -replace '^\<\!\[LOG\[', '') -replace '\]LOG\]\!>.*$', ''
                        } else {
                            continue   # wait for the closing line
                        }
                    } elseif ($null -ne $buffer) {
                        # Continuation of a multi-line entry
                        $buffer += "`n" + $line
                        if ($line -notmatch '\]LOG\]\!>') { continue }
                        # End of multi-line: extract full message
                        $msg = ($buffer -replace '^\<\!\[LOG\[', '') -replace '\]LOG\]\!>.*$', ''
                        $buffer = $null
                    } else {
                        continue
                    }

                    $idx = $msg.IndexOf('Get policies = [')
                    if ($idx -lt 0) { continue }

                    $jsonPart = $msg.Substring($idx + 'Get policies = '.Length)
                    try {
                        $policies = $jsonPart | ConvertFrom-Json -ErrorAction Stop
                    } catch { continue }

                    foreach ($policy in $policies) {
                        if ($policy.Id -and $policy.Name -and -not $nameMap.ContainsKey($policy.Id)) {
                            $nameMap[$policy.Id] = $policy.Name
                        }
                    }
                }
            } finally {
                $reader.Dispose()
                $fs.Dispose()
            }
        } catch { }
    }

    return $nameMap
}

# ── Data loading ──────────────────────────────────────────────────────────────
function Get-TrackedItems {
    $items = [System.Collections.Generic.List[PSCustomObject]]::new()
    $seen  = [System.Collections.Generic.HashSet[string]]::new()

    # ── Win32 apps (IntuneManagementExtension\Win32Apps) ─────────────────────
    # Two possible structures depending on IME version / provisioning phase:
    #   A) ProvisioningProgress sub-key — GUID-named JSON values with FriendlyName+WorkloadState
    #   B) {guid}_{revision} sub-keys — each has ComplianceStateMessage\ComplianceStateMessage JSON
    # We read A first (richer data); B fills in anything A didn't cover.
    $imePath = 'HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps'
    if (Test-Path $imePath) {
        foreach ($userKey in Get-ChildItem $imePath -ErrorAction SilentlyContinue) {

            # ── Structure A: ProvisioningProgress ────────────────────────────
            $provPath = Join-Path $userKey.PSPath 'ProvisioningProgress'
            if (Test-Path $provPath) {
                $provVals = Get-ItemProperty $provPath -ErrorAction SilentlyContinue
                if ($provVals) {
                    foreach ($prop in $provVals.PSObject.Properties) {
                        if ($prop.Name -like 'PS*') { continue }
                        if ($prop.Name -notmatch '^[{(]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[})]?$') { continue }
                        try { $entry = $prop.Value | ConvertFrom-Json -ErrorAction Stop } catch { continue }

                        $guid = $prop.Name.Trim('{}()')
                        if (-not $seen.Add($guid)) { continue }

                        $items.Add([PSCustomObject]@{
                            Name   = if ($entry.FriendlyName) { $entry.FriendlyName } else { $guid }
                            GUID   = $guid
                            Status = Resolve-State $Script:WorkloadStates $entry.WorkloadState
                            Source = 'Win32App'
                        })
                    }
                }
            }

            # ── Structure B: {guid}_{revision} sub-keys ───────────────────────
            foreach ($appKey in Get-ChildItem $userKey.PSPath -ErrorAction SilentlyContinue) {
                # Key name format: "{guid}_{revision}" — strip the version suffix
                $guid = $appKey.PSChildName -replace '_\d+$', ''
                if ($guid -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') { continue }
                if (-not $seen.Add($guid)) { continue }

                # ComplianceStateMessage is a sub-key; its value has the same name
                $csmKeyPath = Join-Path $appKey.PSPath 'ComplianceStateMessage'
                $csmJson    = (Get-ItemProperty $csmKeyPath -ErrorAction SilentlyContinue).ComplianceStateMessage

                $status = 'Unknown'
                if ($csmJson) {
                    try {
                        $csm = $csmJson | ConvertFrom-Json -ErrorAction Stop
                        # Applicability 1 = requirements not met — distinct from genuinely not installed
                        $status = if ($csm.Applicability -eq 1) {
                            'Not Applicable'
                        } else {
                            Resolve-State $Script:ComplianceStates $csm.ComplianceState
                        }
                    } catch { }
                }

                $items.Add([PSCustomObject]@{
                    Name   = $guid   # enriched by IME log pass below
                    GUID   = $guid
                    Status = $status
                    Source = 'Win32App'
                })
            }
        }
    }

    # ── MSI apps (EnterpriseDesktopAppManagement) ─────────────────────────────
    $msiBase = 'HKLM:\SOFTWARE\Microsoft\EnterpriseDesktopAppManagement'
    if (Test-Path $msiBase) {
        foreach ($userKey in Get-ChildItem $msiBase -ErrorAction SilentlyContinue) {
            $msiPath = Join-Path $userKey.PSPath 'MSI'
            if (-not (Test-Path $msiPath)) { continue }

            foreach ($appKey in Get-ChildItem $msiPath -ErrorAction SilentlyContinue) {
                $productCode = $appKey.PSChildName
                # Normalize: strip braces so "{guid}" deduplicates against "guid" from ProvisioningProgress
                $guid = $productCode.Trim('{}')
                if (-not $seen.Add($guid)) { continue }

                $vals   = Get-ItemProperty $appKey.PSPath -ErrorAction SilentlyContinue
                $status = Resolve-State $Script:MSIStates $vals.Status

                # Try to resolve display name from Uninstall registry (try both braced and bare)
                $uninstallPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$productCode"
                $displayName   = (Get-ItemProperty $uninstallPath -ErrorAction SilentlyContinue).DisplayName
                $name          = if ($displayName) { $displayName } else { $guid }

                $items.Add([PSCustomObject]@{
                    Name   = $name
                    GUID   = $guid
                    Status = $status
                    Source = 'MSI'
                })
            }
        }
    }

    # ── Office CSP ────────────────────────────────────────────────────────────
    $officePath = 'HKLM:\SOFTWARE\Microsoft\OfficeCSP'
    if (Test-Path $officePath) {
        foreach ($appKey in Get-ChildItem $officePath -ErrorAction SilentlyContinue) {
            $guid = $appKey.PSChildName
            if (-not $seen.Add("office_$guid")) { continue }

            $vals   = Get-ItemProperty $appKey.PSPath -ErrorAction SilentlyContinue
            $status = Resolve-State $Script:OfficeStates $vals.Status
            $name   = if ($vals.Name) { $vals.Name } else { "Office ($guid)" }

            $items.Add([PSCustomObject]@{
                Name   = $name
                GUID   = $guid
                Status = $status
                Source = 'Office'
            })
        }
    }

    # ── ESP tracking (fallback for anything not covered above) ─────────────────
    # EnrollmentStatusTracking\ESPTrackingInfo\Diagnostics is the authoritative
    # list of what ESP is blocking on. Use it to catch anything not in IME/MSI.
    $diagBase = 'HKLM:\SOFTWARE\Microsoft\Windows\Autopilot\EnrollmentStatusTracking\ESPTrackingInfo\Diagnostics'
    $trackedSections = @('ExpectedMSIAppPackages', 'ExpectedModernAppPackages', 'Sidecar', 'ExpectedSCEPCerts', 'ExpectedPolicies')

    foreach ($section in $trackedSections) {
        $sectionPath = Join-Path $diagBase $section
        if (-not (Test-Path $sectionPath)) { continue }

        foreach ($appKey in Get-ChildItem $sectionPath -ErrorAction SilentlyContinue) {
            $guid = $appKey.PSChildName.Trim('{}()')
            if ($guid -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') { continue }
            if (-not $seen.Add("esp_$guid")) { continue }

            $vals   = Get-ItemProperty $appKey.PSPath -ErrorAction SilentlyContinue
            $status = Resolve-State $Script:ESPStates $vals.InstallationState

            $items.Add([PSCustomObject]@{
                Name   = $guid
                GUID   = $guid
                Status = $status
                Source = $section
            })
        }
    }

    # ── IME log enrichment pass ───────────────────────────────────────────────
    # Upgrades any item whose Name is still a raw GUID (i.e. ProvisioningProgress
    # had no FriendlyName, or it came from the ESP fallback section) using names
    # parsed from the IME log's "Get policies" JSON entry.
    $nameMap = Get-IMENameMap
    if ($nameMap.Count -gt 0) {
        foreach ($item in $items) {
            if ($item.Name -ne $item.GUID) { continue }  # already has a friendly name
            if ($nameMap.ContainsKey($item.GUID)) {
                $item.Name = $nameMap[$item.GUID]
            }
        }
    }

    return $items
}

# ── Script discovery ──────────────────────────────────────────────────────────
function Get-AvailableScripts {
    $map = [ordered]@{}

    # Gallery-installed Get-AutopilotDiagnostics
    $galleryPath = "$env:ProgramFiles\WindowsPowerShell\Scripts\Get-AutopilotDiagnostics.ps1"
    if (Test-Path $galleryPath) {
        $map['Get-AutopilotDiagnostics'] = $galleryPath
    } elseif (Get-Command Get-AutopilotDiagnostics -ErrorAction SilentlyContinue) {
        $map['Get-AutopilotDiagnostics'] = 'Get-AutopilotDiagnostics'
    }

    # Any .ps1 files in the tools folder (community scripts deployed alongside this one)
    if (Test-Path $Script:ToolsFolder) {
        Get-ChildItem "$Script:ToolsFolder\*.ps1" -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notin $Script:ExcludedScripts } |
            ForEach-Object {
                $label = $_.BaseName -replace '-', ' ' -replace '_', ' '
                $map[$label] = $_.FullName
            }
    }

    return $map
}

# ── Form ──────────────────────────────────────────────────────────────────────
$form                  = New-Object System.Windows.Forms.Form
$form.Text             = 'Autopilot ESP Tools'
$form.Size             = New-Object System.Drawing.Size(960, 580)
$form.MinimumSize      = New-Object System.Drawing.Size(700, 400)
$form.StartPosition    = 'CenterScreen'
$form.BackColor        = [System.Drawing.Color]::White
$form.Font             = New-Object System.Drawing.Font('Segoe UI', 9)

# ── Toolbar ───────────────────────────────────────────────────────────────────
$toolbar               = New-Object System.Windows.Forms.Panel
$toolbar.Dock          = 'Top'
$toolbar.Height        = 48
$toolbar.BackColor     = [System.Drawing.Color]::FromArgb(242, 242, 242)
$toolbar.Padding       = New-Object System.Windows.Forms.Padding(6, 6, 6, 6)

$flow                  = New-Object System.Windows.Forms.FlowLayoutPanel
$flow.Dock             = 'Fill'
$flow.FlowDirection    = 'LeftToRight'
$flow.WrapContents     = $false
$flow.BackColor        = [System.Drawing.Color]::Transparent

function New-ToolButton([string]$Text) {
    $b              = New-Object System.Windows.Forms.Button
    $b.Text         = $Text
    $b.AutoSize     = $true
    $b.Height       = 30
    $b.Padding      = New-Object System.Windows.Forms.Padding(6, 0, 6, 0)
    $b.Margin       = New-Object System.Windows.Forms.Padding(0, 0, 4, 0)
    $b.FlatStyle    = 'System'
    return $b
}

function New-Separator {
    $lbl            = New-Object System.Windows.Forms.Label
    $lbl.Text       = '|'
    $lbl.AutoSize   = $false
    $lbl.Width      = 12
    $lbl.Height     = 30
    $lbl.TextAlign  = 'MiddleCenter'
    $lbl.ForeColor  = [System.Drawing.Color]::Silver
    $lbl.Margin     = New-Object System.Windows.Forms.Padding(2, 0, 2, 0)
    return $lbl
}

$btnEventViewer = New-ToolButton 'Event Viewer'
$btnRegedit     = New-ToolButton 'Regedit'
$btnExplorer    = New-ToolButton 'File Explorer'
$btnSetupAct    = New-ToolButton 'SetupAct Log'
$btnIMELog      = New-ToolButton 'IME Log'

$sep1           = New-Separator

$scriptDropdown              = New-Object System.Windows.Forms.ComboBox
$scriptDropdown.DropDownStyle = 'DropDownList'
$scriptDropdown.Width        = 240
$scriptDropdown.Height       = 30
$scriptDropdown.Margin       = New-Object System.Windows.Forms.Padding(0, 0, 4, 0)

$btnRun         = New-ToolButton '> Run Script'

$sep2           = New-Separator
$btnRefresh     = New-ToolButton 'Refresh'

$sep3           = New-Separator

$chkAuto              = New-Object System.Windows.Forms.CheckBox
$chkAuto.Text         = 'Auto'
$chkAuto.AutoSize     = $true
$chkAuto.Height       = 30
$chkAuto.TextAlign    = 'MiddleLeft'
$chkAuto.Margin       = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
$chkAuto.Checked      = $true   # on by default during ESP

$flow.Controls.AddRange(@(
    $btnEventViewer, $btnRegedit, $btnExplorer, $btnSetupAct, $btnIMELog,
    $sep1,
    $scriptDropdown, $btnRun,
    $sep2,
    $btnRefresh, $sep3, $chkAuto
))
$toolbar.Controls.Add($flow)

# ── Info panel ────────────────────────────────────────────────────────────────
$infoPanel            = New-Object System.Windows.Forms.Panel
$infoPanel.Dock       = 'Top'
$infoPanel.Height     = 64
$infoPanel.BackColor  = [System.Drawing.Color]::FromArgb(235, 243, 252)
$infoPanel.Padding    = New-Object System.Windows.Forms.Padding(10, 0, 10, 0)

# Top row — profile labels
$infoRow1             = New-Object System.Windows.Forms.TableLayoutPanel
$infoRow1.Dock        = 'Top'
$infoRow1.Height      = 30
$infoRow1.ColumnCount = 4
$infoRow1.RowCount    = 1
$infoRow1.GrowStyle   = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
$infoRow1.BackColor   = [System.Drawing.Color]::Transparent
$infoRow1.Padding     = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
@(25, 30, 22, 23) | ForEach-Object {
    $cs = New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, $_)
    $infoRow1.ColumnStyles.Add($cs) | Out-Null
}

function New-InfoLabel([string]$Text) {
    $l           = New-Object System.Windows.Forms.Label
    $l.Text      = $Text
    $l.AutoSize  = $false
    $l.Dock      = 'Fill'
    $l.TextAlign = 'MiddleLeft'
    $l.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    return $l
}

$lblDevice = New-InfoLabel 'Device: -'
$lblTenant = New-InfoLabel 'Tenant: -'
$lblMode   = New-InfoLabel 'Mode: -'
$lblUser   = New-InfoLabel 'User: -'
$infoRow1.Controls.Add($lblDevice, 0, 0)
$infoRow1.Controls.Add($lblTenant, 1, 0)
$infoRow1.Controls.Add($lblMode,   2, 0)
$infoRow1.Controls.Add($lblUser,   3, 0)

# Bottom row — BitLocker
$infoRow2            = New-Object System.Windows.Forms.FlowLayoutPanel
$infoRow2.Dock       = 'Fill'
$infoRow2.FlowDirection = 'LeftToRight'
$infoRow2.BackColor  = [System.Drawing.Color]::Transparent
$infoRow2.Padding    = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$infoRow2.WrapContents = $false

$lblBLKey            = New-Object System.Windows.Forms.Label
$lblBLKey.Text       = 'BitLocker C:'
$lblBLKey.AutoSize   = $true
$lblBLKey.TextAlign  = 'MiddleLeft'
$lblBLKey.Height     = 22
$lblBLKey.Margin     = New-Object System.Windows.Forms.Padding(0, 2, 6, 0)
$lblBLKey.Font       = New-Object System.Drawing.Font('Segoe UI', 9)

$blBar               = New-Object System.Windows.Forms.ProgressBar
$blBar.Width         = 130
$blBar.Height        = 16
$blBar.Minimum       = 0
$blBar.Maximum       = 100
$blBar.Value         = 0
$blBar.Margin        = New-Object System.Windows.Forms.Padding(0, 4, 6, 0)
$blBar.Style         = 'Continuous'

$lblBLStatus         = New-Object System.Windows.Forms.Label
$lblBLStatus.Text    = '-'
$lblBLStatus.AutoSize = $true
$lblBLStatus.TextAlign = 'MiddleLeft'
$lblBLStatus.Height  = 22
$lblBLStatus.Margin  = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$lblBLStatus.Font    = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

$infoRow2.Controls.AddRange(@($lblBLKey, $blBar, $lblBLStatus))

$infoPanel.Controls.Add($infoRow2)
$infoPanel.Controls.Add($infoRow1)   # Top-docked row1 goes above fill row2

# ── Status strip ──────────────────────────────────────────────────────────────
$statusStrip    = New-Object System.Windows.Forms.StatusStrip
$statusLabel    = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = 'Ready'
$statusStrip.Items.Add($statusLabel) | Out-Null

# ── DataGridView ──────────────────────────────────────────────────────────────
$grid                              = New-Object System.Windows.Forms.DataGridView
$grid.Dock                         = 'Fill'
$grid.AllowUserToAddRows           = $false
$grid.AllowUserToDeleteRows        = $false
$grid.ReadOnly                     = $true
$grid.SelectionMode                = 'FullRowSelect'
$grid.MultiSelect                  = $false
$grid.AutoSizeColumnsMode          = 'Fill'
$grid.RowHeadersVisible            = $false
$grid.BackgroundColor              = [System.Drawing.Color]::White
$grid.GridColor                    = [System.Drawing.Color]::FromArgb(220, 220, 220)
$grid.BorderStyle                  = 'None'
$grid.CellBorderStyle              = 'SingleHorizontal'
$grid.EnableHeadersVisualStyles    = $false
$grid.ColumnHeadersHeight          = 30
$grid.ColumnHeadersHeightSizeMode  = 'DisableResizing'
$grid.RowTemplate.Height           = 24

$grid.ColumnHeadersDefaultCellStyle.BackColor  = [System.Drawing.Color]::FromArgb(0, 120, 212)
$grid.ColumnHeadersDefaultCellStyle.ForeColor  = [System.Drawing.Color]::White
$grid.ColumnHeadersDefaultCellStyle.Font       = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersDefaultCellStyle.Padding    = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)

$grid.DefaultCellStyle.Padding                 = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 248)

$colName   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.HeaderText  = 'Name'
$colName.FillWeight  = 35

$colGUID   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colGUID.HeaderText  = 'GUID'
$colGUID.FillWeight  = 40

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.HeaderText = 'Status'
$colStatus.FillWeight  = 25

$grid.Columns.AddRange($colName, $colGUID, $colStatus)

# Row and status column formatting
$statusColIndex = $colStatus.Index
$grid.Add_CellFormatting({
    $rowStatus = $grid.Rows[$_.RowIndex].Cells[$statusColIndex].Value

    if ($rowStatus -eq 'Installed') {
        $_.CellStyle.BackColor          = [System.Drawing.Color]::FromArgb(39, 174, 96)
        $_.CellStyle.ForeColor          = [System.Drawing.Color]::White
        $_.CellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(27, 142, 74)
        $_.CellStyle.SelectionForeColor = [System.Drawing.Color]::White
        $_.FormattingApplied = $true
        return
    }

    if ($_.ColumnIndex -ne $statusColIndex) { return }
    $color = Get-StatusColor $_.Value
    $_.CellStyle.ForeColor = $color
    $_.CellStyle.Font = if ($color.R -lt 50 -and $color.G -lt 50 -and $color.B -lt 50) {
        $grid.Font
    } else {
        New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Bold)
    }
    $_.FormattingApplied = $true
})

# ── Layout assembly ───────────────────────────────────────────────────────────
# Docking order matters: higher Controls index docks first (outermost).
# toolbar (Top) must dock before infoPanel (Top) so toolbar sits above it.
$form.Controls.Add($grid)        # Fill  — gets remaining space
$form.Controls.Add($infoPanel)   # Top   — docks above grid, below toolbar
$form.Controls.Add($statusStrip) # Bottom
$form.Controls.Add($toolbar)     # Top   — outermost, docks first

# ── Info panel update ─────────────────────────────────────────────────────────
function Update-InfoPanel {
    $profile = Get-AutopilotProfile
    $user    = Get-AssignedUser
    $bl      = Get-BitLockerStatus

    $lblDevice.Text = "Device: $($profile.DeviceName)"
    $lblTenant.Text = "Tenant: $($profile.Tenant)"
    $lblMode.Text   = "Mode: $($profile.Mode)"
    $lblUser.Text   = "User: $user"

    $blBar.Value       = [Math]::Max(0, [Math]::Min($bl.Percentage, 100))
    $blBar.Visible     = ($bl.State -in @(1, 2, 3, 4, 5))
    $lblBLStatus.Text  = $bl.Text

    $lblBLStatus.ForeColor = switch ($bl.State) {
        1                  { [System.Drawing.Color]::FromArgb(0, 150, 0) }          # Encrypted
        { $_ -in @(2, 4) } { [System.Drawing.Color]::FromArgb(0, 100, 200) }       # Encrypting
        { $_ -in @(3, 5) } { [System.Drawing.Color]::FromArgb(180, 100, 0) }       # Decrypting
        0                  { if ($bl.Protected) {
                                 [System.Drawing.Color]::FromArgb(180, 100, 0)      # Pending - orange
                             } else {
                                 [System.Drawing.Color]::FromArgb(120, 120, 120)    # Not Encrypted - grey
                             }}
        default            { [System.Drawing.Color]::FromArgb(120, 120, 120) }
    }
}

# ── Refresh logic ─────────────────────────────────────────────────────────────
function Invoke-Refresh {
    if ($Script:RefreshBusy) { return }
    $Script:RefreshBusy = $true

    # Remember selected row's GUID so we can restore it after repopulating
    $selectedGuid = $null
    if ($grid.SelectedRows.Count -gt 0) {
        $selectedGuid = $grid.SelectedRows[0].Cells[$colGUID.Index].Value
    }

    $statusLabel.Text = 'Loading...'
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        Update-InfoPanel

        $items = Get-TrackedItems

        $grid.SuspendLayout()
        $grid.Rows.Clear()

        foreach ($item in $items) {
            $grid.Rows.Add($item.Name, $item.GUID, $item.Status) | Out-Null
        }

        # Restore selection if the row still exists
        if ($selectedGuid) {
            foreach ($row in $grid.Rows) {
                if ($row.Cells[$colGUID.Index].Value -eq $selectedGuid) {
                    $row.Selected = $true
                    $grid.FirstDisplayedScrollingRowIndex = $row.Index
                    break
                }
            }
        }

        $total   = $items.Count
        $ok      = ($items | Where-Object Status -eq 'Installed').Count
        $failed  = ($items | Where-Object { $_.Status -in @('Failed', 'Download Failed') }).Count
        $na      = ($items | Where-Object Status -eq 'Not Applicable').Count
        $named   = ($items | Where-Object { $_.Name -ne $_.GUID }).Count
        $naTag   = if ($na -gt 0) { ", $na n/a" } else { '' }
        $autoTag = if ($chkAuto.Checked) { ' (auto)' } else { '' }
        $statusLabel.Text = "$total items | $ok installed, $failed failed$naTag | $named/$total named | $(Get-Date -Format 'HH:mm:ss')$autoTag"
    }
    finally {
        $grid.ResumeLayout()
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $Script:RefreshBusy = $false
    }
}

# ── Script dropdown population ────────────────────────────────────────────────
function Initialize-ScriptDropdown {
    $scriptDropdown.Items.Clear()
    $map = Get-AvailableScripts
    $form.Tag = $map   # stash map for click handler

    foreach ($label in $map.Keys) {
        $scriptDropdown.Items.Add($label) | Out-Null
    }
    if ($scriptDropdown.Items.Count -gt 0) { $scriptDropdown.SelectedIndex = 0 }

    $btnRun.Enabled = ($scriptDropdown.Items.Count -gt 0)
}

# ── Tool button events ────────────────────────────────────────────────────────
$btnEventViewer.Add_Click({ Start-Process 'eventvwr.exe' })

$btnRegedit.Add_Click({ Start-Process 'regedit.exe' })

$btnExplorer.Add_Click({ Start-Process 'explorer.exe' })

$btnSetupAct.Add_Click({
    if (Test-Path $Script:CMTrace) {
        Start-Process $Script:CMTrace -ArgumentList $Script:SetupActLog
    } else {
        Start-Process 'notepad.exe' -ArgumentList $Script:SetupActLog
    }
})

$btnIMELog.Add_Click({
    if (Test-Path $Script:CMTrace) {
        Start-Process $Script:CMTrace -ArgumentList $Script:IMELog
    } else {
        Start-Process 'notepad.exe' -ArgumentList $Script:IMELog
    }
})

$btnRun.Add_Click({
    $map      = $form.Tag
    $selected = $scriptDropdown.SelectedItem
    if (-not $selected -or -not $map -or -not $map.Contains($selected)) { return }

    $target = $map[$selected]
    $args   = if ($target -like '*.ps1') {
        "-NoLogo -NoProfile -NoExit -ExecutionPolicy Bypass -File `"$target`""
    } else {
        "-NoLogo -NoProfile -NoExit -ExecutionPolicy Bypass -Command `"& $target`""
    }
    Start-Process powershell.exe -ArgumentList $args
})

$btnRefresh.Add_Click({ Invoke-Refresh })

# ── Auto-refresh timer ────────────────────────────────────────────────────────
$timer          = New-Object System.Windows.Forms.Timer
$timer.Interval = $Script:AutoRefreshIntervalMs

$timer.Add_Tick({ Invoke-Refresh })

$chkAuto.Add_CheckedChanged({
    if ($chkAuto.Checked) { $timer.Start() } else { $timer.Stop() }
})

$form.Add_FormClosed({ $timer.Stop(); $timer.Dispose() })

# ── Init ──────────────────────────────────────────────────────────────────────
Initialize-ScriptDropdown
Invoke-Refresh
$timer.Start()   # matches $chkAuto.Checked = $true default

[void]$form.ShowDialog()
