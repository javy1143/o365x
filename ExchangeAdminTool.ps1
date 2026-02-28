# ==============================================================================
# Script:           ExchangeAdminTool.ps1
# Description:      A fully asynchronous, multi-tenant capable GUI for managing 
#                   Exchange Online. Prevents UI freezing using Runspaces.
# ==============================================================================

# Ensure required assemblies are loaded
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==============================================================================
# 1. Global Variables & Theme Setup
# ==============================================================================
$script:Runspace = $null
$script:PS = $null
$script:AsyncResult = $null
$script:IsConnected = $false

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
}

# Find the script directory to locate the logo
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if (-not $scriptDir) { $scriptDir = $PWD.Path }
$logoPath = Join-Path $scriptDir "StackPoint IT Logo (no background).png"

# ==============================================================================
# 2. UI Helper Functions
# ==============================================================================
function Write-Log {
    param(
        [string]$Message,
        [System.Drawing.Color]$Color = $Theme.Text
    )
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $Form.Invoke([action]{
        $LogBox.SelectionStart = $LogBox.TextLength
        $LogBox.SelectionLength = 0
        $LogBox.SelectionColor = $Theme.TextDark
        $LogBox.AppendText("[$timestamp] ")
        
        $LogBox.SelectionStart = $LogBox.TextLength
        $LogBox.SelectionLength = 0
        $LogBox.SelectionColor = $Color
        $LogBox.AppendText("$Message`r`n")
        
        $LogBox.ScrollToCaret()
    })
}

function New-UIElement {
    param($Type, $Text, $X, $Y, $W, $H)
    $ctrl = New-Object "System.Windows.Forms.$Type"
    if ($Text) { $ctrl.Text = $Text }
    $ctrl.Location = New-Object System.Drawing.Point($X, $Y)
    if ($W -and $H) { $ctrl.Size = New-Object System.Drawing.Size($W, $H) }
    return $ctrl
}

# ==============================================================================
# 3. Asynchronous Runspace Framework
# ==============================================================================
$script:JobTimer = New-Object System.Windows.Forms.Timer
$script:JobTimer.Interval = 250
$script:JobTimer.add_Tick({
    if ($script:AsyncResult -and $script:AsyncResult.IsCompleted) {
        $script:JobTimer.Stop()
        
        try {
            $results = $script:PS.EndInvoke($script:AsyncResult)
            $errors = $script:PS.Streams.Error
            
            if ($errors.Count -gt 0) {
                foreach ($err in $errors) {
                    Write-Log "ERROR: $($err.Exception.Message)" $Theme.Error
                }
                $script:PS.Streams.Error.Clear()
            }
            
            if ($results) {
                $outString = ($results | Out-String).Trim()
                if (![string]::IsNullOrWhiteSpace($outString)) {
                    Write-Log $outString $Theme.Text
                }
            }
            Write-Log "Operation finished." $Theme.Success
        } catch {
            Write-Log "CRITICAL ERROR: $_" $Theme.Error
        } finally {
            $script:PS.Commands.Clear()
            $Form.Enabled = $true
            $Form.Cursor = [System.Windows.Forms.Cursors]::Default
            
            # If this was a connection attempt, update UI
            if ($script:CurrentAction -eq 'Connect' -and ($errors.Count -eq 0)) {
                $script:IsConnected = $true
                $StatusLabel.Text = "Status: Connected"
                $StatusLabel.ForeColor = $Theme.Success
            } elseif ($script:CurrentAction -eq 'Disconnect') {
                $script:IsConnected = $false
                $StatusLabel.Text = "Status: Disconnected"
                $StatusLabel.ForeColor = $Theme.Error
            }
        }
    }
})

function Invoke-ExchangeCommand {
    param(
        [string]$ActionName,
        [scriptblock]$ScriptBlock
    )
    
    if (-not $script:IsConnected -and $ActionName -ne 'Connect') {
        Write-Log "Please connect to Exchange Online first." $Theme.Warning
        return
    }

    $script:CurrentAction = $ActionName
    Write-Log "Executing: $ActionName..." $Theme.Warning
    
    $Form.Enabled = $false
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    # Initialize Runspace if it doesn't exist
    if ($null -eq $script:Runspace -or $script:Runspace.RunspaceStateInfo.State -ne 'Opened') {
        $script:Runspace = [runspacefactory]::CreateRunspace()
        $script:Runspace.ThreadOptions = "ReuseThread"
        $script:Runspace.Open()
        
        $script:PS = [powershell]::Create()
        $script:PS.Runspace = $script:Runspace
        
        # Suppress progress bars to prevent runspace UI bleeding
        $script:PS.AddScript({ $ProgressPreference = 'SilentlyContinue' }).Invoke()
        $script:PS.Commands.Clear()
    }

    $script:PS.AddScript($ScriptBlock)
    $script:AsyncResult = $script:PS.BeginInvoke()
    $script:JobTimer.Start()
}

# ==============================================================================
# 4. Main Form & UI Construction
# ==============================================================================
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "StackPoint IT - Exchange Online Admin Tool"
$Form.Size = New-Object System.Drawing.Size(800, 700)
$Form.BackColor = $Theme.Background
$Form.ForeColor = $Theme.Text
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "FixedDialog"
$Form.MaximizeBox = $false

# -- Header Panel --
$HeaderPanel = New-UIElement "Panel" "" 0 0 800 60
$HeaderPanel.BackColor = $Theme.Panel

if (Test-Path $logoPath) {
    $LogoBox = New-Object System.Windows.Forms.PictureBox
    $LogoBox.Image = [System.Drawing.Image]::FromFile($logoPath)
    $LogoBox.SizeMode = "Zoom"
    $LogoBox.Size = New-Object System.Drawing.Size(50, 50)
    $LogoBox.Location = New-Object System.Drawing.Point(10, 5)
    $HeaderPanel.Controls.Add($LogoBox)
}

$TitleLabel = New-UIElement "Label" "Exchange Online Administrator" 70 15 400 30
$TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$HeaderPanel.Controls.Add($TitleLabel)

$StatusLabel = New-UIElement "Label" "Status: Disconnected" 600 20 180 20
$StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$StatusLabel.ForeColor = $Theme.Error
$HeaderPanel.Controls.Add($StatusLabel)

$Form.Controls.Add($HeaderPanel)

# -- Tab Control --
$TabControl = New-UIElement "TabControl" "" 10 70 760 380
$TabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9)

function Format-Tab {
    param($Tab)
    $Tab.BackColor = $Theme.Background
    $TabControl.Controls.Add($Tab)
}

# --- Tab 1: Connection ---
$TabConnect = New-UIElement "TabPage" "Connection" 0 0 0 0
Format-Tab $TabConnect

$lblUPN = New-UIElement "Label" "Admin UPN:" 20 30 150 20
$txtUPN = New-UIElement "TextBox" "" 20 50 300 25
$txtUPN.BackColor = $Theme.InputBG; $txtUPN.ForeColor = $Theme.InputText

$btnConnect = New-UIElement "Button" "Connect" 20 90 140 35
$btnConnect.BackColor = $Theme.Accent; $btnConnect.FlatStyle = "Flat"
$btnConnect.Add_Click({
    $upn = $txtUPN.Text
    if (-not $upn) { Write-Log "UPN required!" $Theme.Warning; return }
    
    Invoke-ExchangeCommand -ActionName "Connect" -ScriptBlock {
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -UserPrincipalName "$using:upn" -ShowProgress $false -ShowBanner $false
        Get-Mailbox -ResultSize 1 | Out-Null # Verify connection
    }
})

$btnDisconnect = New-UIElement "Button" "Disconnect" 180 90 140 35
$btnDisconnect.BackColor = $Theme.Panel; $btnDisconnect.FlatStyle = "Flat"
$btnDisconnect.Add_Click({
    Invoke-ExchangeCommand -ActionName "Disconnect" -ScriptBlock {
        Disconnect-ExchangeOnline -Confirm:$false
    }
})

$TabConnect.Controls.AddRange(@($lblUPN, $txtUPN, $btnConnect, $btnDisconnect))

# --- Tab 2: Mailbox Permissions ---
$TabPerms = New-UIElement "TabPage" "Mailbox Permissions" 0 0 0 0
Format-Tab $TabPerms

$lblTargetMbx = New-UIElement "Label" "Target Mailbox (UPN/Alias):" 20 20 200 20
$txtTargetMbx = New-UIElement "TextBox" "" 20 40 250 25
$txtTargetMbx.BackColor = $Theme.InputBG; $txtTargetMbx.ForeColor = $Theme.InputText

$lblUserGroup = New-UIElement "Label" "User/Group to Grant/Remove:" 300 20 200 20
$txtUserGroup = New-UIElement "TextBox" "" 300 40 250 25
$txtUserGroup.BackColor = $Theme.InputBG; $txtUserGroup.ForeColor = $Theme.InputText

$lblPermType = New-UIElement "Label" "Permission Level:" 20 80 150 20
$cmbPermType = New-UIElement "ComboBox" "" 20 100 250 25
$cmbPermType.BackColor = $Theme.InputBG; $cmbPermType.ForeColor = $Theme.InputText
$cmbPermType.DropDownStyle = "DropDownList"
$cmbPermType.Items.AddRange(@("FullAccess", "SendAs", "SendOnBehalf", "Calendar"))
$cmbPermType.SelectedIndex = 0

$btnGetPerms = New-UIElement "Button" "Get Current Permissions" 20 150 160 35
$btnGetPerms.BackColor = $Theme.Panel; $btnGetPerms.FlatStyle = "Flat"
$btnGetPerms.Add_Click({
    $mbx = $txtTargetMbx.Text
    if (-not $mbx) { Write-Log "Target Mailbox required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName "Get Permissions" -ScriptBlock {
        $mbx = "$using:mbx"
        Write-Output "--- Full Access / Send On Behalf ---"
        Get-MailboxPermission -Identity $mbx | Where-Object {(-not $_.IsInherited) -and ($_.User -notlike "NT AUTHORITY*")} | Select-Object User, AccessRights | Format-Table -AutoSize
        Write-Output "--- Send As ---"
        Get-RecipientPermission -Identity $mbx | Where-Object {(-not $_.IsInherited) -and ($_.Trustee -notlike "NT AUTHORITY*")} | Select-Object Trustee, AccessRights | Format-Table -AutoSize
    }
})

$btnAddPerm = New-UIElement "Button" "Add Permission" 200 150 160 35
$btnAddPerm.BackColor = $Theme.Accent; $btnAddPerm.FlatStyle = "Flat"
$btnAddPerm.Add_Click({
    $mbx = $txtTargetMbx.Text; $user = $txtUserGroup.Text; $type = $cmbPermType.Text
    if (-not $mbx -or -not $user) { Write-Log "Mailbox and User required." $Theme.Warning; return }
    
    Invoke-ExchangeCommand -ActionName "Add $type" -ScriptBlock {
        $mbx = "$using:mbx"; $user = "$using:user"; $type = "$using:type"
        switch ($type) {
            "FullAccess" { Add-MailboxPermission -Identity $mbx -User $user -AccessRights FullAccess -InheritanceType All -Confirm:$false }
            "SendAs" { Add-RecipientPermission -Identity $mbx -Trustee $user -AccessRights SendAs -Confirm:$false }
            "SendOnBehalf" { Set-Mailbox -Identity $mbx -GrantSendOnBehalfTo @{Add=$user} }
            "Calendar" {
                # Localized dynamic calendar discovery
                $calFolder = Get-MailboxFolderStatistics -Identity $mbx | Where-Object {$_.FolderType -eq 'Calendar'} | Select-Object -First 1
                if ($calFolder) {
                    $calPath = "$mbx`:$($calFolder.FolderPath.Replace('/','\'))"
                    Add-MailboxFolderPermission -Identity $calPath -User $user -AccessRights Reviewer -Confirm:$false
                } else { throw "Could not find Calendar folder." }
            }
        }
    }
})

$btnRemPerm = New-UIElement "Button" "Remove Permission" 380 150 160 35
$btnRemPerm.BackColor = $Theme.Error; $btnRemPerm.FlatStyle = "Flat"
$btnRemPerm.Add_Click({
    $mbx = $txtTargetMbx.Text; $user = $txtUserGroup.Text; $type = $cmbPermType.Text
    if (-not $mbx -or -not $user) { Write-Log "Mailbox and User required." $Theme.Warning; return }
    
    Invoke-ExchangeCommand -ActionName "Remove $type" -ScriptBlock {
        $mbx = "$using:mbx"; $user = "$using:user"; $type = "$using:type"
        switch ($type) {
            "FullAccess" { Remove-MailboxPermission -Identity $mbx -User $user -AccessRights FullAccess -Confirm:$false }
            "SendAs" { Remove-RecipientPermission -Identity $mbx -Trustee $user -AccessRights SendAs -Confirm:$false }
            "SendOnBehalf" { Set-Mailbox -Identity $mbx -GrantSendOnBehalfTo @{Remove=$user} }
            "Calendar" {
                $calFolder = Get-MailboxFolderStatistics -Identity $mbx | Where-Object {$_.FolderType -eq 'Calendar'} | Select-Object -First 1
                if ($calFolder) {
                    $calPath = "$mbx`:$($calFolder.FolderPath.Replace('/','\'))"
                    Remove-MailboxFolderPermission -Identity $calPath -User $user -Confirm:$false
                }
            }
        }
    }
})

$TabPerms.Controls.AddRange(@($lblTargetMbx, $txtTargetMbx, $lblUserGroup, $txtUserGroup, $lblPermType, $cmbPermType, $btnGetPerms, $btnAddPerm, $btnRemPerm))

# --- Tab 3: Mail Forwarding ---
$TabFwd = New-UIElement "TabPage" "Mail Forwarding" 0 0 0 0
Format-Tab $TabFwd

$lblFwdMbx = New-UIElement "Label" "Target Mailbox (UPN/Alias):" 20 20 200 20
$txtFwdMbx = New-UIElement "TextBox" "" 20 40 250 25
$txtFwdMbx.BackColor = $Theme.InputBG; $txtFwdMbx.ForeColor = $Theme.InputText

$lblFwdTo = New-UIElement "Label" "Forward To (UPN/Email):" 300 20 200 20
$txtFwdTo = New-UIElement "TextBox" "" 300 40 250 25
$txtFwdTo.BackColor = $Theme.InputBG; $txtFwdTo.ForeColor = $Theme.InputText

$chkKeepCopy = New-UIElement "CheckBox" "Deliver to Mailbox and Forward" 20 80 250 20

$btnGetFwd = New-UIElement "Button" "Get Forwarding" 20 120 140 35
$btnGetFwd.BackColor = $Theme.Panel; $btnGetFwd.FlatStyle = "Flat"
$btnGetFwd.Add_Click({
    $mbx = $txtFwdMbx.Text
    if (-not $mbx) { Write-Log "Mailbox required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName "Get Forwarding" -ScriptBlock {
        Get-Mailbox -Identity "$using:mbx" | Select-Object DisplayName, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward | Format-List
    }
})

$btnSetFwd = New-UIElement "Button" "Set Forwarding" 180 120 140 35
$btnSetFwd.BackColor = $Theme.Accent; $btnSetFwd.FlatStyle = "Flat"
$btnSetFwd.Add_Click({
    $mbx = $txtFwdMbx.Text; $fwd = $txtFwdTo.Text; $keep = $chkKeepCopy.Checked
    if (-not $mbx -or -not $fwd) { Write-Log "Mailbox and Forward Address required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName "Set Forwarding" -ScriptBlock {
        Set-Mailbox -Identity "$using:mbx" -ForwardingSmtpAddress "$using:fwd" -DeliverToMailboxAndForward $using:keep
    }
})

$btnRemFwd = New-UIElement "Button" "Remove Forwarding" 340 120 140 35
$btnRemFwd.BackColor = $Theme.Error; $btnRemFwd.FlatStyle = "Flat"
$btnRemFwd.Add_Click({
    $mbx = $txtFwdMbx.Text
    if (-not $mbx) { Write-Log "Mailbox required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName "Remove Forwarding" -ScriptBlock {
        Set-Mailbox -Identity "$using:mbx" -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false
    }
})

$TabFwd.Controls.AddRange(@($lblFwdMbx, $txtFwdMbx, $lblFwdTo, $txtFwdTo, $chkKeepCopy, $btnGetFwd, $btnSetFwd, $btnRemFwd))

# --- Tab 4: Distribution Groups ---
$TabGroups = New-UIElement "TabPage" "Distribution Groups" 0 0 0 0
Format-Tab $TabGroups

$lblGrpName = New-UIElement "Label" "Group Email/Alias:" 20 20 200 20
$txtGrpName = New-UIElement "TextBox" "" 20 40 250 25
$txtGrpName.BackColor = $Theme.InputBG; $txtGrpName.ForeColor = $Theme.InputText

$lblGrpUser = New-UIElement "Label" "User to Add/Remove:" 300 20 200 20
$txtGrpUser = New-UIElement "TextBox" "" 300 40 250 25
$txtGrpUser.BackColor = $Theme.InputBG; $txtGrpUser.ForeColor = $Theme.InputText

$btnGetMembers = New-UIElement "Button" "List Members" 20 90 140 35
$btnGetMembers.BackColor = $Theme.Panel; $btnGetMembers.FlatStyle = "Flat"
$btnGetMembers.Add_Click({
    $grp = $txtGrpName.Text
    if (-not $grp) { Write-Log "Group name required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName "Get Members" -ScriptBlock {
        Get-DistributionGroupMember -Identity "$using:grp" | Select-Object DisplayName, PrimarySmtpAddress | Format-Table -AutoSize
    }
})

$btnAddMember = New-UIElement "Button" "Add Member" 180 90 140 35
$btnAddMember.BackColor = $Theme.Accent; $btnAddMember.FlatStyle = "Flat"
$btnAddMember.Add_Click({
    $grp = $txtGrpName.Text; $user = $txtGrpUser.Text
    if (-not $grp -or -not $user) { Write-Log "Group and User required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName "Add Member" -ScriptBlock {
        Add-DistributionGroupMember -Identity "$using:grp" -Member "$using:user" -Confirm:$false
    }
})

$btnRemMember = New-UIElement "Button" "Remove Member" 340 90 140 35
$btnRemMember.BackColor = $Theme.Error; $btnRemMember.FlatStyle = "Flat"
$btnRemMember.Add_Click({
    $grp = $txtGrpName.Text; $user = $txtGrpUser.Text
    if (-not $grp -or -not $user) { Write-Log "Group and User required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName "Remove Member" -ScriptBlock {
        Remove-DistributionGroupMember -Identity "$using:grp" -Member "$using:user" -Confirm:$false
    }
})

$TabGroups.Controls.AddRange(@($lblGrpName, $txtGrpName, $lblGrpUser, $txtGrpUser, $btnGetMembers, $btnAddMember, $btnRemMember))

# --- Tab 5: Shared Mailboxes ---
$TabShared = New-UIElement "TabPage" "Mailbox Conversion" 0 0 0 0
Format-Tab $TabShared

$lblConvMbx = New-UIElement "Label" "Target Mailbox:" 20 20 200 20
$txtConvMbx = New-UIElement "TextBox" "" 20 40 250 25
$txtConvMbx.BackColor = $Theme.InputBG; $txtConvMbx.ForeColor = $Theme.InputText

$btnGetType = New-UIElement "Button" "Check Current Type" 20 90 150 35
$btnGetType.BackColor = $Theme.Panel; $btnGetType.FlatStyle = "Flat"
$btnGetType.Add_Click({
    $mbx = $txtConvMbx.Text
    if (-not $mbx) { Write-Log "Mailbox required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName "Get Mailbox Type" -ScriptBlock {
        Get-Mailbox -Identity "$using:mbx" | Select-Object DisplayName, RecipientTypeDetails | Format-List
    }
})

$btnToShared = New-UIElement "Button" "Convert to Shared" 190 90 150 35
$btnToShared.BackColor = $Theme.Accent; $btnToShared.FlatStyle = "Flat"
$btnToShared.Add_Click({
    $mbx = $txtConvMbx.Text
    if (-not $mbx) { Write-Log "Mailbox required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName "Convert to Shared" -ScriptBlock {
        Set-Mailbox -Identity "$using:mbx" -Type Shared
    }
})

$btnToReg = New-UIElement "Button" "Convert to Regular" 360 90 150 35
$btnToReg.BackColor = $Theme.Panel; $btnToReg.FlatStyle = "Flat"
$btnToReg.Add_Click({
    $mbx = $txtConvMbx.Text
    if (-not $mbx) { Write-Log "Mailbox required." $Theme.Warning; return }
    Invoke-ExchangeCommand -ActionName "Convert to Regular" -ScriptBlock {
        Set-Mailbox -Identity "$using:mbx" -Type Regular
    }
})

$TabShared.Controls.AddRange(@($lblConvMbx, $txtConvMbx, $btnGetType, $btnToShared, $btnToReg))

$Form.Controls.Add($TabControl)

# -- Log Box Panel --
$LogPanel = New-UIElement "Panel" "" 10 460 760 190
$LogPanel.BackColor = $Theme.Panel

$LogLabel = New-UIElement "Label" "Operation Log:" 5 5 200 20
$LogLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$LogBox = New-Object System.Windows.Forms.RichTextBox
$LogBox.Location = New-Object System.Drawing.Point(5, 25)
$LogBox.Size = New-Object System.Drawing.Size(750, 160)
$LogBox.BackColor = $Theme.InputBG
$LogBox.ForeColor = $Theme.Text
$LogBox.ReadOnly = $true
$LogBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$LogBox.ScrollBars = "Vertical"

$LogPanel.Controls.AddRange(@($LogLabel, $LogBox))
$Form.Controls.Add($LogPanel)

# ==============================================================================
# 5. Cleanup & Execution
# ==============================================================================
$Form.Add_FormClosing({
    if ($script:Runspace) {
        $script:Runspace.Close()
        $script:Runspace.Dispose()
    }
    if ($script:PS) {
        $script:PS.Dispose()
    }
})

Write-Log "Application Started. Awaiting Exchange Online connection..." $Theme.TextDark
$Form.ShowDialog() | Out-Null
