# ==============================================================================
# Script:           ExchangeAdminTool.ps1
# Description:      A fully asynchronous, multi-tenant capable GUI for managing
#                   Exchange Online. Prevents UI freezing using Runspaces.
# Version:          2.1
# ==============================================================================

#Requires -Version 5.1
Set-StrictMode -Version Latest

# Ensure required assemblies are loaded
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Enable DPI awareness before any form is created (Win32 API — compatible with PowerShell 5.1)
try {
    Add-Type -TypeDefinition @"
using System.Runtime.InteropServices;
public class DpiHelper {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
}
"@
    [DpiHelper]::SetProcessDPIAware() | Out-Null
} catch { <# Non-fatal — continue without DPI awareness #> }

# ==============================================================================
# 1. Global Variables & Theme Setup
# ==============================================================================
$script:Runspace        = $null
$script:PS              = $null
$script:AsyncResult     = $null
$script:IsConnected     = $false
$script:CurrentAction   = ''
$script:ConnectedDomain = ''   # Populated on successful connect; used by Resolve-Identity

# Robust script directory detection — StrictMode-safe
$script:ScriptDir = $null
if ($PSScriptRoot -and $PSScriptRoot -ne '') {
    $script:ScriptDir = $PSScriptRoot
}
if (-not $script:ScriptDir) {
    try {
        $cmdPath = $MyInvocation.MyCommand | Select-Object -ExpandProperty Path -ErrorAction SilentlyContinue
        if ($cmdPath) { $script:ScriptDir = Split-Path -Parent $cmdPath }
    } catch { <# Property doesn't exist in all hosts #> }
}
if (-not $script:ScriptDir) {
    $script:ScriptDir = $PWD.Path
}
# Final fallback — guarantee non-null
if (-not $script:ScriptDir -or $script:ScriptDir -eq '') { $script:ScriptDir = $env:TEMP }

$script:LogFilePath = Join-Path $script:ScriptDir "ExchangeAdminTool_$(Get-Date -Format 'yyyyMMdd').log"

$Theme = @{
    Background  = [System.Drawing.ColorTranslator]::FromHtml("#1E1E1E")
    Panel       = [System.Drawing.ColorTranslator]::FromHtml("#252526")
    Text        = [System.Drawing.ColorTranslator]::FromHtml("#FFFFFF")
    TextDark    = [System.Drawing.ColorTranslator]::FromHtml("#CCCCCC")
    Accent      = [System.Drawing.ColorTranslator]::FromHtml("#0078D7")
    AccentHover = [System.Drawing.ColorTranslator]::FromHtml("#005A9E")
    InputBG     = [System.Drawing.ColorTranslator]::FromHtml("#333333")
    InputText   = [System.Drawing.ColorTranslator]::FromHtml("#FFFFFF")
    Success     = [System.Drawing.ColorTranslator]::FromHtml("#28A745")
    Warning     = [System.Drawing.ColorTranslator]::FromHtml("#FFC107")
    Error       = [System.Drawing.ColorTranslator]::FromHtml("#DC3545")
    LogBG       = [System.Drawing.ColorTranslator]::FromHtml("#1A1A1A")
}

# Find script directory for logo
$logoPath = $null
try { $logoPath = Join-Path $script:ScriptDir "StackPoint IT Logo (no background).png" } catch {}

# ==============================================================================
# 2. UI Helper Functions
# ==============================================================================
function Write-Log {
    param(
        [string]$Message,
        [System.Drawing.Color]$Color = $script:Theme_Text,
        [switch]$NoFile
    )
    # Use a module-level reference to Theme since this may be called from timer context
    if (-not $Color) { $Color = [System.Drawing.ColorTranslator]::FromHtml("#FFFFFF") }

    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $fullLine  = "[$timestamp] $Message"

    # Write to file log (audit trail) unless suppressed
    if (-not $NoFile) {
        try { Add-Content -Path $script:LogFilePath -Value $fullLine -ErrorAction SilentlyContinue } catch {}
    }

    $Form.Invoke([action]{
        $LogBox.SelectionStart  = $LogBox.TextLength
        $LogBox.SelectionLength = 0
        $LogBox.SelectionColor  = [System.Drawing.ColorTranslator]::FromHtml("#CCCCCC")
        $LogBox.AppendText("[$timestamp] ")

        $LogBox.SelectionStart  = $LogBox.TextLength
        $LogBox.SelectionLength = 0
        $LogBox.SelectionColor  = $Color
        $LogBox.AppendText("$Message`r`n")
        $LogBox.ScrollToCaret()
    })
}

function New-Label   { param($Text,$X,$Y,$W,$H)
    $c = New-Object System.Windows.Forms.Label
    $c.Text = $Text; $c.Location = [System.Drawing.Point]::new($X,$Y)
    $c.Size = [System.Drawing.Size]::new($W,$H); $c.ForeColor = $Theme.Text
    $c.BackColor = [System.Drawing.Color]::Transparent; return $c }

function New-TextBox { param($X,$Y,$W,$H,[string]$Placeholder='')
    $c = New-Object System.Windows.Forms.TextBox
    $c.Location = [System.Drawing.Point]::new($X,$Y)
    $c.Size = [System.Drawing.Size]::new($W,$H)
    $c.BackColor = $Theme.InputBG; $c.ForeColor = $Theme.InputText
    $c.BorderStyle = 'FixedSingle'; $c.Font = [System.Drawing.Font]::new("Segoe UI", 9)
    if ($Placeholder) {
        $c.Tag = $Placeholder; $c.Text = $Placeholder; $c.ForeColor = [System.Drawing.Color]::Gray
        $c.Add_Enter({ if ($this.Text -eq $this.Tag) { $this.Text = ''; $this.ForeColor = $Theme.InputText } })
        $c.Add_Leave({ if ($this.Text -eq '') { $this.Text = $this.Tag; $this.ForeColor = [System.Drawing.Color]::Gray } })
    }
    return $c }

function Get-TextValue { param($TextBox)
    $v = $TextBox.Text.Trim()
    if ($v -eq $TextBox.Tag) { return '' }
    return $v }

function New-Button  { param($Text,$X,$Y,$W,$H,$BgColor)
    $c = New-Object System.Windows.Forms.Button
    $c.Text = $Text; $c.Location = [System.Drawing.Point]::new($X,$Y)
    $c.Size = [System.Drawing.Size]::new($W,$H)
    $c.BackColor = $BgColor; $c.ForeColor = $Theme.Text
    $c.FlatStyle = 'Flat'; $c.FlatAppearance.BorderSize = 0
    $c.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $c.Cursor = [System.Windows.Forms.Cursors]::Hand
    $c.Tag = $BgColor
    $c.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb([Math]::Min(255,$this.BackColor.R+25),[Math]::Min(255,$this.BackColor.G+25),[Math]::Min(255,$this.BackColor.B+25)) })
    $c.Add_MouseLeave({ $this.BackColor = $this.Tag })
    return $c }

function New-ComboBox { param($X,$Y,$W,$H,[string[]]$Items)
    $c = New-Object System.Windows.Forms.ComboBox
    $c.Location = [System.Drawing.Point]::new($X,$Y); $c.Size = [System.Drawing.Size]::new($W,$H)
    $c.BackColor = $Theme.InputBG; $c.ForeColor = $Theme.InputText
    $c.DropDownStyle = 'DropDownList'; $c.FlatStyle = 'Flat'
    $c.Font = [System.Drawing.Font]::new("Segoe UI", 9)
    foreach ($i in $Items) { $c.Items.Add($i) | Out-Null }
    $c.SelectedIndex = 0; return $c }

function New-Separator { param($X,$Y,$W)
    $c = New-Object System.Windows.Forms.Panel
    $c.Location = [System.Drawing.Point]::new($X,$Y); $c.Size = [System.Drawing.Size]::new($W,1)
    $c.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#444444"); return $c }

# ==============================================================================
# 2b. Identity Resolution Helper
# ==============================================================================
# Appends the connected tenant domain when the user types just a username (no @).
# Passes through any value that already contains '@' unchanged.
function Resolve-Identity {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($Value.Contains('@')) { return $Value }
    if ($script:ConnectedDomain) {
        return "$Value@$($script:ConnectedDomain)"
    }
    # No domain available yet — return as-is and let Exchange resolve by alias
    return $Value
}

# ==============================================================================
# 3. Pre-flight Module Check
# ==============================================================================
function Test-ExchangeModule {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "MISSING MODULE: 'ExchangeOnlineManagement' is not installed." ([System.Drawing.ColorTranslator]::FromHtml("#DC3545"))
        Write-Log "Fix: Open PowerShell as Administrator and run:" ([System.Drawing.ColorTranslator]::FromHtml("#FFC107"))
        Write-Log "     Install-Module ExchangeOnlineManagement -Force" ([System.Drawing.ColorTranslator]::FromHtml("#FFC107"))
        return $false
    }
    return $true
}

# ==============================================================================
# 6. Main Form Construction
# ==============================================================================
$Form = New-Object System.Windows.Forms.Form
$Form.Text            = "StackPoint IT - Exchange Online Admin Tool v2.1"
$Form.Size            = [System.Drawing.Size]::new(820, 900)
$Form.MinimumSize     = [System.Drawing.Size]::new(820, 900)
$Form.BackColor       = $Theme.Background
$Form.ForeColor       = $Theme.Text
$Form.StartPosition   = "CenterScreen"
$Form.FormBorderStyle = "Sizable"
$Form.AutoScaleMode   = [System.Windows.Forms.AutoScaleMode]::Dpi
$Form.Font            = [System.Drawing.Font]::new("Segoe UI", 9)

# -- Header Panel --
$HeaderPanel           = New-Object System.Windows.Forms.Panel
$HeaderPanel.Location  = [System.Drawing.Point]::new(0, 0)
$HeaderPanel.Size      = [System.Drawing.Size]::new(820, 65)
$HeaderPanel.BackColor = $Theme.Panel
$HeaderPanel.Anchor    = "Top, Left, Right"

if ($logoPath -and (Test-Path $logoPath)) {
    $LogoBox          = New-Object System.Windows.Forms.PictureBox
    $LogoBox.Image    = [System.Drawing.Image]::FromFile($logoPath)
    $LogoBox.SizeMode = "Zoom"
    $LogoBox.Size     = [System.Drawing.Size]::new(50, 50)
    $LogoBox.Location = [System.Drawing.Point]::new(10, 7)
    $HeaderPanel.Controls.Add($LogoBox)
}

$TitleLabel           = New-Label "Exchange Online Administrator" 70 18 400 28
$TitleLabel.Font      = [System.Drawing.Font]::new("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$HeaderPanel.Controls.Add($TitleLabel)

$StatusLabel           = New-Label "Status: Disconnected" 580 22 220 22
$StatusLabel.Font      = [System.Drawing.Font]::new("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$StatusLabel.ForeColor = $Theme.Error
$StatusLabel.TextAlign = "MiddleRight"
$HeaderPanel.Controls.Add($StatusLabel)

$Form.Controls.Add($HeaderPanel)

# -- Tab Control --
$TabControl           = New-Object System.Windows.Forms.TabControl
$TabControl.Location  = [System.Drawing.Point]::new(10, 75)
$TabControl.Size      = [System.Drawing.Size]::new(780, 400)
$TabControl.Font      = [System.Drawing.Font]::new("Segoe UI", 9)
$TabControl.Anchor    = "Top, Left, Right, Bottom"
$TabControl.Appearance = "Normal"

function New-Tab {
    param([string]$Title)
    $t = New-Object System.Windows.Forms.TabPage
    $t.Text      = $Title
    $t.BackColor = $Theme.Background
    $t.ForeColor = $Theme.Text
    $t.Padding   = [System.Windows.Forms.Padding]::new(10)
    $TabControl.Controls.Add($t)
    return $t
}

# ==============================================================================
# 4. Runspace Teardown Helper
# ==============================================================================
function Remove-Runspace {
    if ($script:PS) {
        try { $script:PS.Stop() } catch {}
        try { $script:PS.Dispose() } catch {}
        $script:PS = $null
    }
    if ($script:Runspace) {
        try { $script:Runspace.Close() } catch {}
        try { $script:Runspace.Dispose() } catch {}
        $script:Runspace = $null
    }
    $script:AsyncResult = $null
}

# ==============================================================================
# 5. Asynchronous Runspace Framework
# ==============================================================================
$script:JobTimer          = New-Object System.Windows.Forms.Timer
$script:JobTimer.Interval = 250

$script:JobTimer.add_Tick({
    if ($null -eq $script:AsyncResult -or -not $script:AsyncResult.IsCompleted) { return }
    $script:JobTimer.Stop()

    $errorCount = 0
    try {
        $results = $script:PS.EndInvoke($script:AsyncResult)
        $errorCount = $script:PS.Streams.Error.Count

        foreach ($err in $script:PS.Streams.Error) {
            Write-Log "ERROR: $($err.Exception.Message)" ([System.Drawing.ColorTranslator]::FromHtml("#DC3545"))
        }
        $script:PS.Streams.Error.Clear()

        if ($results) {
            $sentinel    = $results | Where-Object { $_ -eq '__CONNECTED__' }
            $domainLines = $results | Where-Object { $_ -like '__DOMAIN__*' }
            $rowLines    = $results | Where-Object { $_ -like '__ROW__*' }
            $clearGrid   = $results | Where-Object { $_ -eq '__CLEARGRID__' }
            $outputLines = $results | Where-Object { $_ -ne '__CONNECTED__' -and $_ -notlike '__ROW__*' -and $_ -ne '__CLEARGRID__' -and $_ -notlike '__DOMAIN__*' }

            # Capture connected domain from sentinel
            foreach ($dl in $domainLines) {
                $script:ConnectedDomain = ($dl -replace '^__DOMAIN__\|','').Trim()
            }

            # Populate results grid if we got structured rows
            if ($clearGrid -or $rowLines) {
                $Form.Invoke([action]{
                    $ResultsGrid.Items.Clear()
                    $ResultsPanel.Visible = $true
                })
            }
            foreach ($rowLine in $rowLines) {
                $parts = ($rowLine -replace '^__ROW__\|','') -split '\|'
                $li = New-Object System.Windows.Forms.ListViewItem($parts[0])
                $li.SubItems.Add($(if ($parts.Count -gt 1) { $parts[1] } else { '' })) | Out-Null
                $li.SubItems.Add($(if ($parts.Count -gt 2) { $parts[2] } else { '' })) | Out-Null
                $li.SubItems.Add($(if ($parts.Count -gt 3) { $parts[3] } else { '' })) | Out-Null
                $liRef = $li
                $Form.Invoke([action]{ $ResultsGrid.Items.Add($liRef) | Out-Null })
            }
            if ($rowLines) {
                $count = @($rowLines).Count
                $Form.Invoke([action]{ $ResultsHeaderLabel.Text = "Current Delegates  ($count entries - select rows to remove)" })
            }

            $output = $outputLines | Out-String
            if (![string]::IsNullOrWhiteSpace($output)) {
                Write-Log $output.Trim() ([System.Drawing.ColorTranslator]::FromHtml("#FFFFFF"))
            }
            if ($sentinel) { $errorCount = -1 }
        }

        if ($errorCount -eq 0) {
            Write-Log "Operation completed successfully." ([System.Drawing.ColorTranslator]::FromHtml("#28A745"))
        } elseif ($errorCount -ne -1) {
            Write-Log "Operation completed with $errorCount error(s)." ([System.Drawing.ColorTranslator]::FromHtml("#FFC107"))
        }
    } catch {
        Write-Log "CRITICAL ERROR: $_" ([System.Drawing.ColorTranslator]::FromHtml("#DC3545"))
        $errorCount = 1
    } finally {
        $script:PS.Commands.Clear()
        $Form.Invoke([action]{
            $Form.Enabled = $true
            $Form.Cursor  = [System.Windows.Forms.Cursors]::Default
        })

        if ($script:CurrentAction -eq 'Connect') {
            if ($errorCount -eq -1) {
                $script:IsConnected = $true
                $Form.Invoke([action]{
                    $StatusLabel.Text      = "Status: Connected ($($script:ConnectedDomain))"
                    $StatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#28A745")
                })
                Write-Log "Successfully connected to Exchange Online ($($script:ConnectedDomain))." ([System.Drawing.ColorTranslator]::FromHtml("#28A745"))
            } else {
                $script:IsConnected = $false
                $script:ConnectedDomain = ''
                Write-Log "Connection failed. Check credentials and module." ([System.Drawing.ColorTranslator]::FromHtml("#DC3545"))
            }
        } elseif ($script:CurrentAction -eq 'Disconnect') {
            $script:IsConnected = $false
            $script:ConnectedDomain = ''
            Remove-Runspace
            $Form.Invoke([action]{
                $StatusLabel.Text      = "Status: Disconnected"
                $StatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#DC3545")
                $ResultsGrid.Items.Clear()
                $ResultsPanel.Visible = $false
            })
            Write-Log "Disconnected from Exchange Online." ([System.Drawing.ColorTranslator]::FromHtml("#FFC107"))
        }
    }
})

function Invoke-ExchangeCommand {
    param(
        [string]$ActionName,
        [scriptblock]$ScriptBlock,
        [hashtable]$Variables = @{}
    )

    if (-not $script:IsConnected -and $ActionName -ne 'Connect') {
        Write-Log "Not connected. Please connect to Exchange Online first." ([System.Drawing.ColorTranslator]::FromHtml("#FFC107"))
        return
    }
    if (-not (Test-ExchangeModule)) { return }

    $script:CurrentAction = $ActionName
    Write-Log "Starting: $ActionName ..." ([System.Drawing.ColorTranslator]::FromHtml("#FFC107"))

    $Form.Enabled = $false
    $Form.Cursor  = [System.Windows.Forms.Cursors]::WaitCursor

    # Build fresh runspace if needed (always fresh after disconnect)
    if ($null -eq $script:Runspace -or $script:Runspace.RunspaceStateInfo.State -ne 'Opened') {
        $script:Runspace = [runspacefactory]::CreateRunspace()
        $script:Runspace.ThreadOptions = "ReuseThread"
        $script:Runspace.ApartmentState = "STA"
        $script:Runspace.Open()

        $script:PS = [powershell]::Create()
        $script:PS.Runspace = $script:Runspace

        # Silence progress bars
        $script:Runspace.SessionStateProxy.SetVariable('ProgressPreference', 'SilentlyContinue')
        $script:Runspace.SessionStateProxy.SetVariable('ErrorActionPreference', 'Stop')
    }

    # Safely inject caller variables into runspace scope
    foreach ($key in $Variables.Keys) {
        $script:Runspace.SessionStateProxy.SetVariable($key, $Variables[$key])
    }

    $script:PS.Commands.Clear()
    $script:PS.AddScript($ScriptBlock) | Out-Null
    $script:AsyncResult = $script:PS.BeginInvoke()
    $script:JobTimer.Start()
}

# ==============================================================================
# 7. TAB 1 — Connection
# ==============================================================================
$TabConnect = New-Tab "Connection"

$lblUPN       = New-Label "Admin UPN (e.g. admin@contoso.com):" 15 20 310 20
$txtUPN       = New-TextBox 15 42 310 26 "admin@domain.com"
$txtUPN.Font  = [System.Drawing.Font]::new("Segoe UI", 9)

$btnConnect    = New-Button "Connect" 15 88 140 36 $Theme.Accent
$btnDisconnect = New-Button "Disconnect" 170 88 140 36 $Theme.Panel

$btnConnect.Add_Click({
    $upn = Get-TextValue $txtUPN
    if (-not $upn) { Write-Log "Admin UPN is required." $Theme.Warning; return }
    if (-not $upn.Contains('@')) { Write-Log "Enter a full UPN with domain (e.g. admin@contoso.com)." $Theme.Warning; return }
    if (-not (Test-ExchangeModule)) { return }

    # Extract the domain portion from the UPN for org auto-detection
    $domain = $upn.Split('@')[-1]

    Invoke-ExchangeCommand -ActionName 'Connect' -Variables @{ upn = $upn; domain = $domain } -ScriptBlock {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        $params = @{
            UserPrincipalName = $upn
            ShowProgress      = $false
            ShowBanner        = $false
        }
        Connect-ExchangeOnline @params
        # Sentinel — verifies cmdlets are available
        Get-Mailbox -ResultSize 1 -ErrorAction Stop | Out-Null
        Write-Output "__DOMAIN__|$domain"
        Write-Output '__CONNECTED__'
    }
})

$btnDisconnect.Add_Click({
    if (-not $script:IsConnected) { Write-Log "Not currently connected." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName 'Disconnect' -ScriptBlock {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
})

$sep1 = New-Separator 15 140 730
$lblConnNote = New-Label "The tenant organization is auto-detected from your sign-in domain. Just enter your admin UPN and connect." 15 150 700 36
$lblConnNote.ForeColor = $Theme.TextDark
$lblConnNote.Font      = [System.Drawing.Font]::new("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)

$lblInputNote = New-Label "All input fields across the tool accept short usernames (e.g. `"jsmith`") in addition to full email addresses. The connected domain is appended automatically." 15 185 700 36
$lblInputNote.ForeColor = $Theme.TextDark
$lblInputNote.Font      = [System.Drawing.Font]::new("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)

$TabConnect.Controls.AddRange(@($lblUPN,$txtUPN,$btnConnect,$btnDisconnect,$sep1,$lblConnNote,$lblInputNote))

# ==============================================================================
# 8. TAB 2 — Mailbox Permissions
# ==============================================================================
$TabPerms = New-Tab "Mailbox Permissions"

$lblTargetMbx  = New-Label "Target Mailbox (UPN or Username):" 15 15 280 20
$txtTargetMbx  = New-TextBox 15 36 230 26 "username or user@domain.com"

$lblUserGroup  = New-Label "User / Group to Grant or Remove:" 265 15 280 20
$txtUserGroup  = New-TextBox 265 36 230 26 "username or user@domain.com"

$btnValidateMbx = New-Button "Validate" 510 34 100 28 $Theme.Panel
$btnValidateMbx.Font = [System.Drawing.Font]::new("Segoe UI", 8)
$btnValidateMbx.Add_Click({
    $mbx = Resolve-Identity (Get-TextValue $txtTargetMbx)
    if (-not $mbx) { Write-Log "Enter a mailbox to validate." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName 'Validate Mailbox' -Variables @{ mbx = $mbx } -ScriptBlock {
        $result = Get-Mailbox -Identity $mbx -ErrorAction Stop | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails
        Write-Output "Found: $($result.DisplayName) | $($result.PrimarySmtpAddress) | $($result.RecipientTypeDetails)"
    }
})

$lblPermType  = New-Label "Permission Type:" 15 78 150 20
$cmbPermType  = New-ComboBox 15 98 200 26 @("FullAccess","SendAs","SendOnBehalf","Calendar - Reviewer","Calendar - Editor","Calendar - Owner")

$btnGetPerms  = New-Button "Get All Permissions" 15 145 175 34 $Theme.Panel
$btnAddPerm   = New-Button "Add Permission"       205 145 150 34 $Theme.Accent
$btnRemPerm   = New-Button "Remove Permission"    370 145 150 34 $Theme.Error

$btnGetPerms.Add_Click({
    $mbx = Resolve-Identity (Get-TextValue $txtTargetMbx)
    if (-not $mbx) { Write-Log "Target Mailbox required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName 'Get Permissions' -Variables @{ mbx = $mbx } -ScriptBlock {
        Write-Output '__CLEARGRID__'

        # FullAccess
        $fa = Get-MailboxPermission -Identity $mbx |
            Where-Object { -not $_.IsInherited -and $_.User -notlike "NT AUTHORITY*" }
        foreach ($r in $fa) { Write-Output "__ROW__|$mbx|FullAccess|$($r.User)|" }

        # SendAs
        $sa = Get-RecipientPermission -Identity $mbx |
            Where-Object { -not $_.IsInherited -and $_.Trustee -notlike "NT AUTHORITY*" }
        foreach ($r in $sa) { Write-Output "__ROW__|$mbx|SendAs|$($r.Trustee)|" }

        # SendOnBehalf
        $sob = (Get-Mailbox -Identity $mbx).GrantSendOnBehalfTo
        foreach ($r in $sob) { Write-Output "__ROW__|$mbx|SendOnBehalf|$r|" }

        # Calendar
        $calFolder = Get-MailboxFolderStatistics -Identity $mbx |
            Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
        if ($calFolder) {
            $calPath = "$($mbx):\$($calFolder.FolderPath.TrimStart('/').Replace('/','\'))"
            $calPerms = Get-MailboxFolderPermission -Identity $calPath |
                Where-Object { $_.User.DisplayName -notmatch "Default|Anonymous" }
            foreach ($r in $calPerms) { Write-Output "__ROW__|$mbx|Calendar|$($r.User.DisplayName)|$($r.AccessRights -join ',')" }
        }

        # Human-readable summary to log
        $faCount  = @($fa).Count
        $saCount  = @($sa).Count
        $sobCount = if ($sob) { @($sob).Count } else { 0 }
        Write-Output "Permissions loaded: $faCount FullAccess, $saCount SendAs, $sobCount SendOnBehalf + Calendar entries."
    }
})

$btnAddPerm.Add_Click({
    $mbx  = Resolve-Identity (Get-TextValue $txtTargetMbx)
    $user = Resolve-Identity (Get-TextValue $txtUserGroup)
    $type = $cmbPermType.Text
    if (-not $mbx -or -not $user) { Write-Log "Mailbox and User are required." $Theme.Warning; return }

    Invoke-ExchangeCommand -ActionName "Add $type" -Variables @{ mbx=$mbx; user=$user; type=$type } -ScriptBlock {
        switch -Wildcard ($type) {
            "FullAccess"   { Add-MailboxPermission    -Identity $mbx -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $false -Confirm:$false }
            "SendAs"       { Add-RecipientPermission  -Identity $mbx -Trustee $user -AccessRights SendAs -Confirm:$false }
            "SendOnBehalf" { Set-Mailbox -Identity $mbx -GrantSendOnBehalfTo @{Add=$user} }
            "Calendar*"    {
                $level = $type.Split(' - ')[-1]
                $calFolder = Get-MailboxFolderStatistics -Identity $mbx |
                    Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
                if (-not $calFolder) { throw "Could not locate Calendar folder for $mbx" }
                $calPath = "$($mbx):\$($calFolder.FolderPath.TrimStart('/').Replace('/','\'))"
                $existing = Get-MailboxFolderPermission -Identity $calPath -User $user -ErrorAction SilentlyContinue
                if ($existing) {
                    Set-MailboxFolderPermission -Identity $calPath -User $user -AccessRights $level -Confirm:$false
                    Write-Output "Updated existing calendar permission to $level."
                } else {
                    Add-MailboxFolderPermission -Identity $calPath -User $user -AccessRights $level -Confirm:$false
                }
            }
        }
    }
})

$btnRemPerm.Add_Click({
    $mbx  = Resolve-Identity (Get-TextValue $txtTargetMbx)
    $user = Resolve-Identity (Get-TextValue $txtUserGroup)
    $type = $cmbPermType.Text
    if (-not $mbx -or -not $user) { Write-Log "Mailbox and User are required." $Theme.Warning; return }

    Invoke-ExchangeCommand -ActionName "Remove $type" -Variables @{ mbx=$mbx; user=$user; type=$type } -ScriptBlock {
        switch -Wildcard ($type) {
            "FullAccess"   { Remove-MailboxPermission   -Identity $mbx -User $user -AccessRights FullAccess -Confirm:$false }
            "SendAs"       { Remove-RecipientPermission -Identity $mbx -Trustee $user -AccessRights SendAs -Confirm:$false }
            "SendOnBehalf" { Set-Mailbox -Identity $mbx -GrantSendOnBehalfTo @{Remove=$user} }
            "Calendar*"    {
                $calFolder = Get-MailboxFolderStatistics -Identity $mbx |
                    Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
                if (-not $calFolder) { throw "Could not locate Calendar folder for $mbx" }
                $calPath = "$($mbx):\$($calFolder.FolderPath.TrimStart('/').Replace('/','\'))"
                $existing = Get-MailboxFolderPermission -Identity $calPath -User $user -ErrorAction SilentlyContinue
                if ($existing) {
                    Remove-MailboxFolderPermission -Identity $calPath -User $user -Confirm:$false
                } else {
                    Write-Output "No calendar permission found for $user on $mbx - nothing to remove."
                }
            }
        }
    }
})

$TabPerms.Controls.AddRange(@($lblTargetMbx,$txtTargetMbx,$lblUserGroup,$txtUserGroup,$btnValidateMbx,
    $lblPermType,$cmbPermType,$btnGetPerms,$btnAddPerm,$btnRemPerm))


# ==============================================================================
# 9. TAB 3 — Mail Forwarding
# ==============================================================================
$TabFwd = New-Tab "Mail Forwarding"

$lblFwdMbx  = New-Label "Target Mailbox (UPN or Username):" 15 15 280 20
$txtFwdMbx  = New-TextBox 15 36 230 26 "username or user@domain.com"

$lblFwdTo   = New-Label "Forward To (UPN, Username, or External Email):" 265 15 320 20
$txtFwdTo   = New-TextBox 265 36 280 26 "username or dest@domain.com"

$chkKeepCopy = New-Object System.Windows.Forms.CheckBox
$chkKeepCopy.Text      = "Keep a copy in the original mailbox (Deliver & Forward)"
$chkKeepCopy.Location  = [System.Drawing.Point]::new(15, 80)
$chkKeepCopy.Size      = [System.Drawing.Size]::new(450, 22)
$chkKeepCopy.ForeColor = $Theme.Text
$chkKeepCopy.Checked   = $true

$btnGetFwd = New-Button "Get Forwarding"    15 118 155 34 $Theme.Panel
$btnSetFwd = New-Button "Set Forwarding"    185 118 155 34 $Theme.Accent
$btnRemFwd = New-Button "Remove Forwarding" 355 118 155 34 $Theme.Error

$btnGetFwd.Add_Click({
    $mbx = Resolve-Identity (Get-TextValue $txtFwdMbx)
    if (-not $mbx) { Write-Log "Mailbox required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName 'Get Forwarding' -Variables @{ mbx=$mbx } -ScriptBlock {
        Write-Output '__CLEARGRID__'
        $result = Get-Mailbox -Identity $mbx | Select-Object DisplayName, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward
        if (-not $result.ForwardingAddress -and -not $result.ForwardingSmtpAddress) {
            Write-Output "No forwarding configured on $mbx."
        } else {
            $fwdTarget = if ($result.ForwardingSmtpAddress) { $result.ForwardingSmtpAddress } else { $result.ForwardingAddress }
            $keepCopy  = $result.DeliverToMailboxAndForward
            Write-Output "__ROW__|$mbx|Forwarding|$fwdTarget|KeepCopy:$keepCopy"
            Write-Output "Forwarding active: $mbx -> $fwdTarget (KeepCopy: $keepCopy)"
        }
    }
})

$btnSetFwd.Add_Click({
    $mbx  = Resolve-Identity (Get-TextValue $txtFwdMbx)
    $fwd  = Resolve-Identity (Get-TextValue $txtFwdTo)
    $keep = $chkKeepCopy.Checked
    if (-not $mbx -or -not $fwd) { Write-Log "Mailbox and Forward Address required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName 'Set Forwarding' -Variables @{ mbx=$mbx; fwd=$fwd; keep=$keep } -ScriptBlock {
        Set-Mailbox -Identity $mbx -ForwardingSmtpAddress $fwd -DeliverToMailboxAndForward $keep -ErrorAction Stop
        Write-Output "Forwarding set: $mbx -> $fwd | Keep copy: $keep"
    }
})

$btnRemFwd.Add_Click({
    $mbx = Resolve-Identity (Get-TextValue $txtFwdMbx)
    if (-not $mbx) { Write-Log "Mailbox required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName 'Remove Forwarding' -Variables @{ mbx=$mbx } -ScriptBlock {
        Set-Mailbox -Identity $mbx -ForwardingSmtpAddress $null -ForwardingAddress $null -DeliverToMailboxAndForward $false -ErrorAction Stop
        Write-Output "All forwarding removed from $mbx."
    }
})

$sep2 = New-Separator 15 168 730
$lblFwdNote = New-Label "Warning: Setting external SMTP forwarding may be blocked by your tenant's outbound spam policy. Use internal forwarding addresses where possible." 15 176 700 36
$lblFwdNote.ForeColor = $Theme.Warning
$lblFwdNote.Font      = [System.Drawing.Font]::new("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)

$TabFwd.Controls.AddRange(@($lblFwdMbx,$txtFwdMbx,$lblFwdTo,$txtFwdTo,$chkKeepCopy,$btnGetFwd,$btnSetFwd,$btnRemFwd,$sep2,$lblFwdNote))

# ==============================================================================
# 10. TAB 4 — Distribution Groups
# ==============================================================================
$TabGroups = New-Tab "Distribution Groups"

$lblGrpName  = New-Label "Group Email or Alias:" 15 15 220 20
$txtGrpName  = New-TextBox 15 36 250 26 "groupname or group@domain.com"

$lblGrpUser  = New-Label "User to Add / Remove:" 285 15 220 20
$txtGrpUser  = New-TextBox 285 36 250 26 "username or user@domain.com"

$btnGetMembers  = New-Button "List Members"   15 90 148 34 $Theme.Panel
$btnAddMember   = New-Button "Add Member"      178 90 148 34 $Theme.Accent
$btnRemMember   = New-Button "Remove Member"   340 90 148 34 $Theme.Error

$btnGetMembers.Add_Click({
    $grp = Resolve-Identity (Get-TextValue $txtGrpName)
    if (-not $grp) { Write-Log "Group name required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName 'List Members' -Variables @{ grp=$grp } -ScriptBlock {
        Write-Output '__CLEARGRID__'
        $members = Get-DistributionGroupMember -Identity $grp -ResultSize Unlimited |
            Select-Object DisplayName, PrimarySmtpAddress, RecipientType
        foreach ($m in $members) {
            Write-Output "__ROW__|$grp|GroupMember|$($m.PrimarySmtpAddress)|$($m.DisplayName)"
        }
        Write-Output "Group '$grp' - $(@($members).Count) member(s) loaded."
    }
})

$btnAddMember.Add_Click({
    $grp  = Resolve-Identity (Get-TextValue $txtGrpName)
    $user = Resolve-Identity (Get-TextValue $txtGrpUser)
    if (-not $grp -or -not $user) { Write-Log "Group and User required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName 'Add Member' -Variables @{ grp=$grp; user=$user } -ScriptBlock {
        Add-DistributionGroupMember -Identity $grp -Member $user -Confirm:$false -ErrorAction Stop
        Write-Output "Added $user to $grp."
    }
})

$btnRemMember.Add_Click({
    $grp  = Resolve-Identity (Get-TextValue $txtGrpName)
    $user = Resolve-Identity (Get-TextValue $txtGrpUser)
    if (-not $grp -or -not $user) { Write-Log "Group and User required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName 'Remove Member' -Variables @{ grp=$grp; user=$user } -ScriptBlock {
        Remove-DistributionGroupMember -Identity $grp -Member $user -Confirm:$false -ErrorAction Stop
        Write-Output "Removed $user from $grp."
    }
})

$TabGroups.Controls.AddRange(@($lblGrpName,$txtGrpName,$lblGrpUser,$txtGrpUser,$btnGetMembers,$btnAddMember,$btnRemMember))

# ==============================================================================
# 11. TAB 5 — Mailbox Conversion
# ==============================================================================
$TabShared = New-Tab "Mailbox Conversion"

$lblConvMbx  = New-Label "Target Mailbox (UPN or Username):" 15 15 280 20
$txtConvMbx  = New-TextBox 15 36 280 26 "username or user@domain.com"

$btnGetType   = New-Button "Check Type"          15 90 148 34 $Theme.Panel
$btnToShared  = New-Button "Convert to Shared"   178 90 170 34 $Theme.Accent
$btnToReg     = New-Button "Convert to Regular"  362 90 170 34 $Theme.Panel

$btnGetType.Add_Click({
    $mbx = Resolve-Identity (Get-TextValue $txtConvMbx)
    if (-not $mbx) { Write-Log "Mailbox required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName 'Get Mailbox Type' -Variables @{ mbx=$mbx } -ScriptBlock {
        Get-Mailbox -Identity $mbx | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, IsShared | Format-List
    }
})

$btnToShared.Add_Click({
    $mbx = Resolve-Identity (Get-TextValue $txtConvMbx)
    if (-not $mbx) { Write-Log "Mailbox required." $Theme.Warning; return }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Convert '$mbx' to a Shared Mailbox?`n`nNote: The user's license can be removed after conversion.",
        "Confirm Conversion", "YesNo", "Warning")
    if ($confirm -ne "Yes") { Write-Log "Conversion cancelled." $Theme.TextDark; return }
    Invoke-ExchangeCommand -ActionName 'Convert to Shared' -Variables @{ mbx=$mbx } -ScriptBlock {
        Set-Mailbox -Identity $mbx -Type Shared -ErrorAction Stop
        Write-Output "Converted $mbx to Shared Mailbox."
    }
})

$btnToReg.Add_Click({
    $mbx = Resolve-Identity (Get-TextValue $txtConvMbx)
    if (-not $mbx) { Write-Log "Mailbox required." $Theme.Warning; return }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Convert '$mbx' to a Regular (User) Mailbox?`n`nNote: A valid license must be assigned after conversion.",
        "Confirm Conversion", "YesNo", "Warning")
    if ($confirm -ne "Yes") { Write-Log "Conversion cancelled." $Theme.TextDark; return }
    Invoke-ExchangeCommand -ActionName 'Convert to Regular' -Variables @{ mbx=$mbx } -ScriptBlock {
        Set-Mailbox -Identity $mbx -Type Regular -ErrorAction Stop
        Write-Output "Converted $mbx to Regular User Mailbox."
    }
})

$sep3 = New-Separator 15 140 730
$lblConvNote = New-Label "After converting to Shared: sign into M365 Admin Center to remove the user license (saves cost). The mailbox remains accessible." 15 148 700 36
$lblConvNote.ForeColor = $Theme.TextDark
$lblConvNote.Font      = [System.Drawing.Font]::new("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)

$TabShared.Controls.AddRange(@($lblConvMbx,$txtConvMbx,$btnGetType,$btnToShared,$btnToReg,$sep3,$lblConvNote))

# ==============================================================================
# 12. TAB 6 — Bulk Operations
# ==============================================================================
$TabBulk = New-Tab "Bulk Operations"

$lblBulkInfo = New-Label "CSV Format  -  Required columns:  Mailbox, Target, Permission  (one row per operation)" 15 15 700 20
$lblBulkInfo.ForeColor = $Theme.TextDark

$lblBulkFile = New-Label "CSV File Path:" 15 50 150 20
$txtBulkFile = New-TextBox 15 70 480 26 "C:\path\to\bulk_operations.csv"

$btnBrowseCSV = New-Button "Browse" 510 68 100 28 $Theme.Panel
$btnBrowseCSV.Font = [System.Drawing.Font]::new("Segoe UI", 8)
$btnBrowseCSV.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.Title  = "Select Bulk Operations CSV"
    if ($dlg.ShowDialog() -eq "OK") {
        $txtBulkFile.Text     = $dlg.FileName
        $txtBulkFile.ForeColor = $Theme.InputText
    }
})

$lblBulkAction = New-Label "Bulk Action:" 15 112 120 20
$cmbBulkAction = New-ComboBox 15 132 200 26 @("Add Permission","Remove Permission","Set Forwarding","Add Group Member","Remove Group Member")

$btnPreviewCSV = New-Button "Preview CSV"    15 178 140 34 $Theme.Panel
$btnRunBulk    = New-Button "Run Bulk Job"   170 178 140 34 $Theme.Accent
$lblBulkStatus = New-Label "" 330 185 380 22
$lblBulkStatus.ForeColor = $Theme.TextDark

$btnPreviewCSV.Add_Click({
    $path = Get-TextValue $txtBulkFile
    if (-not $path -or -not (Test-Path $path)) { Write-Log "Valid CSV path required." $Theme.Warning; return }
    try {
        $rows = Import-Csv -Path $path | Select-Object -First 5
        Write-Log "=== CSV Preview (first 5 rows) ===" $Theme.TextDark
        Write-Log ($rows | Format-Table -AutoSize | Out-String) $Theme.Text
        $lblBulkStatus.Text = "Loaded: $((Import-Csv $path).Count) rows"
    } catch {
        Write-Log "Could not read CSV: $_" $Theme.Error
    }
})

$btnRunBulk.Add_Click({
    $path   = Get-TextValue $txtBulkFile
    $action = $cmbBulkAction.Text
    if (-not $path -or -not (Test-Path $path)) { Write-Log "Valid CSV path required." $Theme.Warning; return }

    $rows = Import-Csv -Path $path
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Run '$action' for $($rows.Count) rows from CSV?`nThis cannot be undone.",
        "Confirm Bulk Operation", "YesNo", "Warning")
    if ($confirm -ne "Yes") { Write-Log "Bulk operation cancelled." $Theme.TextDark; return }

    Invoke-ExchangeCommand -ActionName "Bulk: $action" -Variables @{ rows=$rows; action=$action } -ScriptBlock {
        $success = 0; $fail = 0
        foreach ($row in $rows) {
            try {
                switch ($action) {
                    "Add Permission"       { Add-MailboxPermission   -Identity $row.Mailbox -User $row.Target -AccessRights $row.Permission -Confirm:$false }
                    "Remove Permission"    { Remove-MailboxPermission -Identity $row.Mailbox -User $row.Target -AccessRights $row.Permission -Confirm:$false }
                    "Set Forwarding"       { Set-Mailbox -Identity $row.Mailbox -ForwardingSmtpAddress $row.Target -DeliverToMailboxAndForward $true }
                    "Add Group Member"     { Add-DistributionGroupMember    -Identity $row.Mailbox -Member $row.Target -Confirm:$false }
                    "Remove Group Member"  { Remove-DistributionGroupMember -Identity $row.Mailbox -Member $row.Target -Confirm:$false }
                }
                Write-Output "  OK: $($row.Mailbox)"
                $success++
            } catch {
                Write-Output "  FAIL: $($row.Mailbox) - $($_.Exception.Message)"
                $fail++
            }
        }
        Write-Output "--- Bulk complete: $success succeeded, $fail failed ---"
    }
})

$sep4 = New-Separator 15 222 730
$lblCSVTemplate = New-Label "Example CSV:   Mailbox,Target,Permission" 15 230 500 20
$lblCSVTemplate.ForeColor = $Theme.TextDark
$lblCSVEx       = New-Label "               john@contoso.com,jane@contoso.com,FullAccess" 15 248 500 20
$lblCSVEx.ForeColor = $Theme.TextDark
$lblCSVEx.Font  = [System.Drawing.Font]::new("Consolas", 8)

$TabBulk.Controls.AddRange(@($lblBulkInfo,$lblBulkFile,$txtBulkFile,$btnBrowseCSV,$lblBulkAction,$cmbBulkAction,$btnPreviewCSV,$btnRunBulk,$lblBulkStatus,$sep4,$lblCSVTemplate,$lblCSVEx))


# ==============================================================================
# 13. Assemble Tab Control, Results Panel & Log Panel
# ==============================================================================
$Form.Controls.Add($TabControl)

# -- Results Panel (sits between tabs and log, populated by search operations) --
$ResultsPanel          = New-Object System.Windows.Forms.Panel
$ResultsPanel.Location = [System.Drawing.Point]::new(10, 485)
$ResultsPanel.Size     = [System.Drawing.Size]::new(780, 170)
$ResultsPanel.BackColor = $Theme.Panel
$ResultsPanel.Anchor   = "Top, Left, Right"
$ResultsPanel.Visible  = $false

$ResultsHeaderPanel          = New-Object System.Windows.Forms.Panel
$ResultsHeaderPanel.Location = [System.Drawing.Point]::new(0, 0)
$ResultsHeaderPanel.Size     = [System.Drawing.Size]::new(780, 28)
$ResultsHeaderPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1A1A1A")

$ResultsHeaderLabel      = New-Label "Current Delegates" 8 5 500 20
$ResultsHeaderLabel.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$ResultsHeaderPanel.Controls.Add($ResultsHeaderLabel)

$btnRemoveSelected      = New-Button "Remove Selected" 580 2 170 24 $Theme.Error
$btnRemoveSelected.Font = [System.Drawing.Font]::new("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnRemoveSelected.Add_Click({
    $selected = $ResultsGrid.SelectedItems
    if ($selected.Count -eq 0) {
        Write-Log "No rows selected. Click a row in the results panel to select it." $Theme.Warning
        return
    }

    $confirmMsg = "Remove $($selected.Count) selected permission(s)?`n"
    foreach ($row in $selected) { $confirmMsg += "`n  $($row.SubItems[1].Text): $($row.SubItems[2].Text) on $($row.SubItems[0].Text)" }
    $confirm = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "Confirm Removal", "YesNo", "Warning")
    if ($confirm -ne "Yes") { return }

    $removeList = @()
    foreach ($row in $selected) {
        $removeList += [PSCustomObject]@{
            Mailbox  = $row.SubItems[0].Text
            PermType = $row.SubItems[1].Text
            Target   = $row.SubItems[2].Text
            Extra    = $row.SubItems[3].Text
        }
    }

    Invoke-ExchangeCommand -ActionName 'Remove Selected' -Variables @{ removeList=$removeList } -ScriptBlock {
        foreach ($item in $removeList) {
            try {
                switch ($item.PermType) {
                    'FullAccess'   { Remove-MailboxPermission   -Identity $item.Mailbox -User $item.Target -AccessRights FullAccess -Confirm:$false }
                    'SendAs'       { Remove-RecipientPermission -Identity $item.Mailbox -Trustee $item.Target -AccessRights SendAs -Confirm:$false }
                    'SendOnBehalf' { Set-Mailbox -Identity $item.Mailbox -GrantSendOnBehalfTo @{Remove=$item.Target} }
                    'Forwarding'   { Set-Mailbox -Identity $item.Mailbox -ForwardingSmtpAddress $null -ForwardingAddress $null -DeliverToMailboxAndForward $false }
                    'GroupMember'  { Remove-DistributionGroupMember -Identity $item.Mailbox -Member $item.Target -Confirm:$false }
                    'Calendar'     {
                        $calFolder = Get-MailboxFolderStatistics -Identity $item.Mailbox |
                            Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
                        if ($calFolder) {
                            $calPath = "$($item.Mailbox):\$($calFolder.FolderPath.TrimStart('/').Replace('/','\'))"
                            $existing = Get-MailboxFolderPermission -Identity $calPath -User $item.Target -ErrorAction SilentlyContinue
                            if ($existing) { Remove-MailboxFolderPermission -Identity $calPath -User $item.Target -Confirm:$false }
                        }
                    }
                }
                Write-Output "  OK: Removed $($item.PermType) for $($item.Target) on $($item.Mailbox)"
            } catch {
                Write-Output "  FAIL: $($item.PermType) for $($item.Target): $($_.Exception.Message)"
            }
        }
    }

    $Form.Invoke([action]{
        foreach ($row in @($selected)) { $ResultsGrid.Items.Remove($row) }
        $remaining = $ResultsGrid.Items.Count
        $ResultsHeaderLabel.Text = "Current Delegates  ($remaining entries - select rows to remove)"
        if ($remaining -eq 0) { $ResultsPanel.Visible = $false }
    })
})
$ResultsHeaderPanel.Controls.Add($btnRemoveSelected)

$btnClearResults      = New-Button "Clear" 510 2 64 24 $Theme.Panel
$btnClearResults.Font = [System.Drawing.Font]::new("Segoe UI", 8)
$btnClearResults.Add_Click({
    $ResultsGrid.Items.Clear()
    $ResultsPanel.Visible = $false
})
$ResultsHeaderPanel.Controls.Add($btnClearResults)

# ListView — full-width grid of delegates
$ResultsGrid                    = New-Object System.Windows.Forms.ListView
$ResultsGrid.Location           = [System.Drawing.Point]::new(0, 28)
$ResultsGrid.Size               = [System.Drawing.Size]::new(780, 142)
$ResultsGrid.View               = 'Details'
$ResultsGrid.FullRowSelect      = $true
$ResultsGrid.MultiSelect        = $true
$ResultsGrid.GridLines          = $true
$ResultsGrid.BackColor          = $Theme.InputBG
$ResultsGrid.ForeColor          = $Theme.Text
$ResultsGrid.Font               = [System.Drawing.Font]::new("Consolas", 8.5)
$ResultsGrid.BorderStyle        = 'None'
$ResultsGrid.HeaderStyle        = 'Nonclickable'
$ResultsGrid.Anchor             = "Top, Bottom, Left, Right"

$col1 = New-Object System.Windows.Forms.ColumnHeader; $col1.Text = "Mailbox / Group";   $col1.Width = 220
$col2 = New-Object System.Windows.Forms.ColumnHeader; $col2.Text = "Permission Type";   $col2.Width = 130
$col3 = New-Object System.Windows.Forms.ColumnHeader; $col3.Text = "User / Trustee";    $col3.Width = 220
$col4 = New-Object System.Windows.Forms.ColumnHeader; $col4.Text = "Details";           $col4.Width = 190
$ResultsGrid.Columns.AddRange(@($col1,$col2,$col3,$col4))

$ResultsGrid.Add_SelectedIndexChanged({
    $btnRemoveSelected.BackColor = if ($ResultsGrid.SelectedItems.Count -gt 0) {
        [System.Drawing.ColorTranslator]::FromHtml("#DC3545")
    } else {
        $Theme.Panel
    }
})

$ResultsPanel.Controls.AddRange(@($ResultsHeaderPanel, $ResultsGrid))
$Form.Controls.Add($ResultsPanel)

$ResultsPanel.Add_VisibleChanged({
    if ($ResultsPanel.Visible) {
        $LogPanel.Location = [System.Drawing.Point]::new(10, 665)
    } else {
        $LogPanel.Location = [System.Drawing.Point]::new(10, 490)
    }
})

# -- Log Panel --
$LogPanel          = New-Object System.Windows.Forms.Panel
$LogPanel.Location = [System.Drawing.Point]::new(10, 665)
$LogPanel.Size     = [System.Drawing.Size]::new(780, 195)
$LogPanel.BackColor = $Theme.Panel
$LogPanel.Anchor   = "Bottom, Left, Right"

$LogHeaderPanel          = New-Object System.Windows.Forms.Panel
$LogHeaderPanel.Location = [System.Drawing.Point]::new(0, 0)
$LogHeaderPanel.Size     = [System.Drawing.Size]::new(780, 28)
$LogHeaderPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1A1A1A")

$LogLabel          = New-Label "Operation Log" 8 5 160 20
$LogLabel.Font     = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$LogHeaderPanel.Controls.Add($LogLabel)

$btnSaveLog          = New-Button "Save Log" 640 2 110 24 $Theme.Panel
$btnSaveLog.Font     = [System.Drawing.Font]::new("Segoe UI", 8)
$btnSaveLog.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = "Text Files (*.txt)|*.txt|Log Files (*.log)|*.log"
    $dlg.FileName = "ExchangeLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    if ($dlg.ShowDialog() -eq "OK") {
        $LogBox.Text | Out-File -FilePath $dlg.FileName -Encoding UTF8
        Write-Log "Log saved to: $($dlg.FileName)" $Theme.Success
    }
})
$LogHeaderPanel.Controls.Add($btnSaveLog)

$btnClearLog      = New-Button "Clear" 560 2 75 24 $Theme.Panel
$btnClearLog.Font = [System.Drawing.Font]::new("Segoe UI", 8)
$btnClearLog.Add_Click({ $LogBox.Clear() })
$LogHeaderPanel.Controls.Add($btnClearLog)

$LogBox          = New-Object System.Windows.Forms.RichTextBox
$LogBox.Location = [System.Drawing.Point]::new(0, 28)
$LogBox.Size     = [System.Drawing.Size]::new(780, 167)
$LogBox.BackColor = $Theme.LogBG
$LogBox.ForeColor = $Theme.Text
$LogBox.ReadOnly  = $true
$LogBox.Font      = [System.Drawing.Font]::new("Consolas", 8.5)
$LogBox.ScrollBars = "Vertical"
$LogBox.BorderStyle = "None"
$LogBox.Anchor    = "Top, Bottom, Left, Right"
$LogBox.WordWrap  = $false

$LogPanel.Controls.AddRange(@($LogHeaderPanel, $LogBox))
$Form.Controls.Add($LogPanel)

# -- Status Bar --
$StatusBar          = New-Object System.Windows.Forms.StatusStrip
$StatusBar.BackColor = $Theme.Panel
$StatusBar.ForeColor = $Theme.TextDark

$StatusBarLabel           = New-Object System.Windows.Forms.ToolStripStatusLabel
$logPathDisplay = if ($script:LogFilePath) { $script:LogFilePath } else { 'N/A' }
$StatusBarLabel.Text      = "Ready  |  Log file: $logPathDisplay"
$StatusBarLabel.ForeColor = $Theme.TextDark
$StatusBar.Items.Add($StatusBarLabel) | Out-Null
$Form.Controls.Add($StatusBar)

# -- Session timeout watchdog (warn after 50 mins) --
$script:SessionTimer          = New-Object System.Windows.Forms.Timer
$script:SessionTimer.Interval = 3000000  # 50 minutes
$script:SessionTimer.add_Tick({
    if ($script:IsConnected) {
        Write-Log "WARNING: Exchange Online session may be near expiry (50 min). Consider reconnecting." $Theme.Warning
    }
})

# ==============================================================================
# 14. Cleanup & Execution
# ==============================================================================
$Form.Add_FormClosing({
    $script:JobTimer.Stop()
    $script:SessionTimer.Stop()
    Remove-Runspace
})

$Form.Add_Shown({
    $script:SessionTimer.Start()
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "ExchangeOnlineManagement module not found." ([System.Drawing.ColorTranslator]::FromHtml("#DC3545")) -NoFile
        Write-Log "  Run in an elevated PowerShell window:" ([System.Drawing.ColorTranslator]::FromHtml("#FFC107")) -NoFile
        Write-Log "  Install-Module ExchangeOnlineManagement -Force" ([System.Drawing.ColorTranslator]::FromHtml("#FFC107")) -NoFile
    } else {
        $modVer = (Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1).Version
        Write-Log "ExchangeOnlineManagement v$modVer found. Ready to connect." ([System.Drawing.ColorTranslator]::FromHtml("#28A745")) -NoFile
    }
    $logPath = $null
    try { $logPath = Get-Variable -Name 'LogFilePath' -Scope Script -ValueOnly -ErrorAction SilentlyContinue } catch {}
    if ($logPath) {
        Write-Log "Session log: $logPath" ([System.Drawing.ColorTranslator]::FromHtml("#CCCCCC")) -NoFile
    }
})

[System.Windows.Forms.Application]::Run($Form)
