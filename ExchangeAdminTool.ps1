#requires -Version 5.1
<#
.SYNOPSIS
    O365 Exchange Online WinForms administration tool.

.DESCRIPTION
    GUI helper for common mailbox administration tasks:
    - Connect/disconnect to Exchange Online
    - Mailbox permissions
    - Calendar permissions
    - Distribution group membership
    - Shared mailbox access
    - Automatic replies (Out of Office)

.NOTES
    Requires ExchangeOnlineManagement module.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Theme colors sampled from provided logo palette ---
$Theme = @{
    Primary       = [System.Drawing.Color]::FromArgb(0, 137, 214)  # Bright blue
    PrimaryDark   = [System.Drawing.Color]::FromArgb(0, 96, 158)   # Deep blue
    PrimaryLight  = [System.Drawing.Color]::FromArgb(51, 166, 233) # Light blue
    AccentDark    = [System.Drawing.Color]::FromArgb(53, 58, 66)    # Charcoal
    AccentMid     = [System.Drawing.Color]::FromArgb(76, 83, 92)    # Slate gray
    Surface       = [System.Drawing.Color]::FromArgb(241, 245, 249) # Light surface
    Text          = [System.Drawing.Color]::FromArgb(31, 41, 55)    # Dark text
    Success       = [System.Drawing.Color]::FromArgb(16, 185, 129)
    Warning       = [System.Drawing.Color]::FromArgb(245, 158, 11)
    Danger        = [System.Drawing.Color]::FromArgb(220, 38, 38)
}

$script:IsConnected = $false

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error')]
        [string]$Level = 'Info'
    )

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $entry = "[$timestamp] [$Level] $Message"
    $txtLog.AppendText($entry + [Environment]::NewLine)

    switch ($Level) {
        'Success' { $statusLabel.ForeColor = $Theme.Success }
        'Warning' { $statusLabel.ForeColor = $Theme.Warning }
        'Error'   { $statusLabel.ForeColor = $Theme.Danger }
        default   { $statusLabel.ForeColor = $Theme.Text }
    }

    $statusLabel.Text = $Message
}

function global:InvokeSafely {
    param(
        [ScriptBlock]$Script,
        [string]$SuccessMessage,
        [string]$ErrorPrefix = 'Operation failed'
    )

    try {
        & $Script
        if ($SuccessMessage) {
            Write-Log -Message $SuccessMessage -Level Success
        }
    }
    catch {
        Write-Log -Message "${ErrorPrefix}: $($_.Exception.Message)" -Level Error
    }
}

function Ensure-ExchangeModule {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        [System.Windows.Forms.MessageBox]::Show(
            "ExchangeOnlineManagement module is not installed.`n`nInstall with:`nInstall-Module ExchangeOnlineManagement -Scope CurrentUser",
            'Missing Module',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $false
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    return $true
}

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = 'O365 Exchange Admin Console'
$form.Size = New-Object System.Drawing.Size(1080, 760)
$form.StartPosition = 'CenterScreen'
$form.BackColor = $Theme.Surface
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Top'
$headerPanel.Height = 58
$headerPanel.BackColor = $Theme.PrimaryDark
$form.Controls.Add($headerPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'Office 365 Exchange Administration Tool'
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(16, 14)
$headerPanel.Controls.Add($titleLabel)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = [System.Drawing.Color]::White
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = 'Ready. Connect to Exchange Online to begin.'
$statusLabel.ForeColor = $Theme.Text
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(10, 68)
$tabs.Size = New-Object System.Drawing.Size(1044, 600)
$tabs.Appearance = 'Normal'
$form.Controls.Add($tabs)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.Location = New-Object System.Drawing.Point(10, 670)
$txtLog.Size = New-Object System.Drawing.Size(1044, 50)
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.ForeColor = $Theme.Text
$form.Controls.Add($txtLog)

function New-LabeledTextBox {
    param(
        [System.Windows.Forms.Control]$Parent,
        [string]$Label,
        [int]$X,
        [int]$Y,
        [int]$Width = 280
    )

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label
    $lbl.Location = New-Object System.Drawing.Point($X, $Y)
    $lbl.AutoSize = $true
    $lbl.ForeColor = $Theme.Text
    $Parent.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point($X, ($Y + 20))
    $txt.Size = New-Object System.Drawing.Size($Width, 24)
    $Parent.Controls.Add($txt)

    return $txt
}

function New-ThemeButton {
    param(
        [System.Windows.Forms.Control]$Parent,
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 180,
        [System.Drawing.Color]$BackColor = $Theme.Primary
    )

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, $Y)
    $btn.Size = New-Object System.Drawing.Size($Width, 32)
    $btn.BackColor = $BackColor
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatStyle = 'Flat'
    $btn.FlatAppearance.BorderSize = 0
    $btn.Cursor = 'Hand'
    $Parent.Controls.Add($btn)

    return $btn
}

# --- Connection tab ---
$tabConnection = New-Object System.Windows.Forms.TabPage
$tabConnection.Text = 'Connection'
$tabConnection.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabConnection)

$txtAdminUPN = New-LabeledTextBox -Parent $tabConnection -Label 'Admin UPN' -X 20 -Y 20 -Width 320
$btnConnect = New-ThemeButton -Parent $tabConnection -Text 'Connect Exchange Online' -X 20 -Y 80 -Width 220 -BackColor $Theme.Primary
$btnDisconnect = New-ThemeButton -Parent $tabConnection -Text 'Disconnect' -X 250 -Y 80 -Width 120 -BackColor $Theme.AccentDark

$lblConnectionInfo = New-Object System.Windows.Forms.Label
$lblConnectionInfo.Text = 'Tip: Use a role account with Exchange Administrator permissions.'
$lblConnectionInfo.Location = New-Object System.Drawing.Point(20, 130)
$lblConnectionInfo.AutoSize = $true
$lblConnectionInfo.ForeColor = $Theme.AccentMid
$tabConnection.Controls.Add($lblConnectionInfo)

# --- Permissions tab ---
$tabPermissions = New-Object System.Windows.Forms.TabPage
$tabPermissions.Text = 'Mailbox Permissions'
$tabPermissions.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabPermissions)

$txtPermMailbox = New-LabeledTextBox -Parent $tabPermissions -Label 'Target Mailbox (UPN or alias)' -X 20 -Y 20
$txtPermUser = New-LabeledTextBox -Parent $tabPermissions -Label 'Delegate User (UPN)' -X 330 -Y 20

$cmbPermType = New-Object System.Windows.Forms.ComboBox
$cmbPermType.Location = New-Object System.Drawing.Point(640, 40)
$cmbPermType.Size = New-Object System.Drawing.Size(180, 24)
$cmbPermType.DropDownStyle = 'DropDownList'
$cmbPermType.Items.AddRange(@('FullAccess','SendAs','SendOnBehalf'))
$cmbPermType.SelectedIndex = 0
$tabPermissions.Controls.Add($cmbPermType)

$lblPermType = New-Object System.Windows.Forms.Label
$lblPermType.Text = 'Permission Type'
$lblPermType.Location = New-Object System.Drawing.Point(640, 20)
$lblPermType.AutoSize = $true
$tabPermissions.Controls.Add($lblPermType)

$btnGrantPerm = New-ThemeButton -Parent $tabPermissions -Text 'Grant Permission' -X 20 -Y 90
$btnRemovePerm = New-ThemeButton -Parent $tabPermissions -Text 'Remove Permission' -X 210 -Y 90 -BackColor $Theme.AccentDark

# --- Calendar tab ---
$tabCalendar = New-Object System.Windows.Forms.TabPage
$tabCalendar.Text = 'Calendars'
$tabCalendar.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabCalendar)

$txtCalMailbox = New-LabeledTextBox -Parent $tabCalendar -Label 'Mailbox (owner)' -X 20 -Y 20
$txtCalUser = New-LabeledTextBox -Parent $tabCalendar -Label 'User to grant/revoke' -X 330 -Y 20

$cmbCalAccess = New-Object System.Windows.Forms.ComboBox
$cmbCalAccess.Location = New-Object System.Drawing.Point(640, 40)
$cmbCalAccess.Size = New-Object System.Drawing.Size(180, 24)
$cmbCalAccess.DropDownStyle = 'DropDownList'
$cmbCalAccess.Items.AddRange(@('Reviewer','Editor','PublishingEditor','AvailabilityOnly','LimitedDetails'))
$cmbCalAccess.SelectedIndex = 0
$tabCalendar.Controls.Add($cmbCalAccess)

$lblCalAccess = New-Object System.Windows.Forms.Label
$lblCalAccess.Text = 'Access Right'
$lblCalAccess.Location = New-Object System.Drawing.Point(640, 20)
$lblCalAccess.AutoSize = $true
$tabCalendar.Controls.Add($lblCalAccess)

$btnAddCal = New-ThemeButton -Parent $tabCalendar -Text 'Set Calendar Permission' -X 20 -Y 90
$btnRemoveCal = New-ThemeButton -Parent $tabCalendar -Text 'Remove Calendar Permission' -X 210 -Y 90 -BackColor $Theme.AccentDark -Width 210

# --- Groups tab ---
$tabGroups = New-Object System.Windows.Forms.TabPage
$tabGroups.Text = 'Email Groups'
$tabGroups.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabGroups)

$txtGroupName = New-LabeledTextBox -Parent $tabGroups -Label 'Distribution Group (name or alias)' -X 20 -Y 20
$txtGroupMember = New-LabeledTextBox -Parent $tabGroups -Label 'User to add/remove (UPN)' -X 330 -Y 20

$btnGroupAdd = New-ThemeButton -Parent $tabGroups -Text 'Add Member' -X 20 -Y 90
$btnGroupRemove = New-ThemeButton -Parent $tabGroups -Text 'Remove Member' -X 210 -Y 90 -BackColor $Theme.AccentDark

# --- Shared mailbox tab ---
$tabShared = New-Object System.Windows.Forms.TabPage
$tabShared.Text = 'Shared Mailboxes'
$tabShared.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabShared)

$txtSharedMailbox = New-LabeledTextBox -Parent $tabShared -Label 'Shared Mailbox (UPN or alias)' -X 20 -Y 20
$txtSharedUser = New-LabeledTextBox -Parent $tabShared -Label 'User (UPN)' -X 330 -Y 20

$btnSharedGrant = New-ThemeButton -Parent $tabShared -Text 'Grant Full + SendAs' -X 20 -Y 90
$btnSharedRevoke = New-ThemeButton -Parent $tabShared -Text 'Revoke Full + SendAs' -X 210 -Y 90 -BackColor $Theme.AccentDark

# --- Out of Office tab ---
$tabOOF = New-Object System.Windows.Forms.TabPage
$tabOOF.Text = 'Out of Office'
$tabOOF.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabOOF)

$txtOOFMailbox = New-LabeledTextBox -Parent $tabOOF -Label 'Mailbox (UPN)' -X 20 -Y 20

$lblInternal = New-Object System.Windows.Forms.Label
$lblInternal.Text = 'Internal Message'
$lblInternal.Location = New-Object System.Drawing.Point(20, 70)
$lblInternal.AutoSize = $true
$tabOOF.Controls.Add($lblInternal)

$txtInternal = New-Object System.Windows.Forms.TextBox
$txtInternal.Multiline = $true
$txtInternal.Location = New-Object System.Drawing.Point(20, 90)
$txtInternal.Size = New-Object System.Drawing.Size(460, 120)
$tabOOF.Controls.Add($txtInternal)

$lblExternal = New-Object System.Windows.Forms.Label
$lblExternal.Text = 'External Message'
$lblExternal.Location = New-Object System.Drawing.Point(500, 70)
$lblExternal.AutoSize = $true
$tabOOF.Controls.Add($lblExternal)

$txtExternal = New-Object System.Windows.Forms.TextBox
$txtExternal.Multiline = $true
$txtExternal.Location = New-Object System.Drawing.Point(500, 90)
$txtExternal.Size = New-Object System.Drawing.Size(460, 120)
$tabOOF.Controls.Add($txtExternal)

$btnEnableOOF = New-ThemeButton -Parent $tabOOF -Text 'Enable OOF' -X 20 -Y 230
$btnDisableOOF = New-ThemeButton -Parent $tabOOF -Text 'Disable OOF' -X 210 -Y 230 -BackColor $Theme.AccentDark

function Assert-Connected {
    if (-not $script:IsConnected) {
        throw 'Not connected to Exchange Online. Use the Connection tab first.'
    }
}

$btnConnect.Add_Click({
    if (-not (Ensure-ExchangeModule)) { return }

    $adminUpn = $txtAdminUPN.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($adminUpn)) {
        Write-Log -Message 'Enter an admin UPN before connecting.' -Level Warning
        return
    }

    global:InvokeSafely -Script {
        Connect-ExchangeOnline -UserPrincipalName $adminUpn -ShowBanner:$false
        $script:IsConnected = $true
    } -SuccessMessage "Connected to Exchange Online as $adminUpn" -ErrorPrefix 'Connection failed'
})

$btnDisconnect.Add_Click({
    global:InvokeSafely -Script {
        Disconnect-ExchangeOnline -Confirm:$false
        $script:IsConnected = $false
    } -SuccessMessage 'Disconnected from Exchange Online.' -ErrorPrefix 'Disconnect failed'
})

$btnGrantPerm.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtPermMailbox.Text.Trim()
        $user = $txtPermUser.Text.Trim()
        $perm = $cmbPermType.SelectedItem

        switch ($perm) {
            'FullAccess'   { Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -AutoMapping:$true }
            'SendAs'       { Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false }
            'SendOnBehalf' { Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo @{Add=$user} }
        }
    } -SuccessMessage 'Permission granted successfully.'
})

$btnRemovePerm.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtPermMailbox.Text.Trim()
        $user = $txtPermUser.Text.Trim()
        $perm = $cmbPermType.SelectedItem

        switch ($perm) {
            'FullAccess'   { Remove-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -Confirm:$false }
            'SendAs'       { Remove-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false }
            'SendOnBehalf' { Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo @{Remove=$user} }
        }
    } -SuccessMessage 'Permission removed successfully.'
})

$btnAddCal.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtCalMailbox.Text.Trim()
        $user = $txtCalUser.Text.Trim()
        $access = $cmbCalAccess.SelectedItem
        Add-MailboxFolderPermission -Identity "$mailbox`:\Calendar" -User $user -AccessRights $access
    } -SuccessMessage 'Calendar permission updated.'
})

$btnRemoveCal.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtCalMailbox.Text.Trim()
        $user = $txtCalUser.Text.Trim()
        Remove-MailboxFolderPermission -Identity "$mailbox`:\Calendar" -User $user -Confirm:$false
    } -SuccessMessage 'Calendar permission removed.'
})

$btnGroupAdd.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        Add-DistributionGroupMember -Identity $txtGroupName.Text.Trim() -Member $txtGroupMember.Text.Trim()
    } -SuccessMessage 'Group member added.'
})

$btnGroupRemove.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        Remove-DistributionGroupMember -Identity $txtGroupName.Text.Trim() -Member $txtGroupMember.Text.Trim() -Confirm:$false
    } -SuccessMessage 'Group member removed.'
})

$btnSharedGrant.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtSharedMailbox.Text.Trim()
        $user = $txtSharedUser.Text.Trim()
        Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -AutoMapping:$true
        Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false
    } -SuccessMessage 'Shared mailbox access granted.'
})

$btnSharedRevoke.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtSharedMailbox.Text.Trim()
        $user = $txtSharedUser.Text.Trim()
        Remove-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -Confirm:$false
        Remove-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false
    } -SuccessMessage 'Shared mailbox access revoked.'
})

$btnEnableOOF.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        Set-MailboxAutoReplyConfiguration -Identity $txtOOFMailbox.Text.Trim() -AutoReplyState Enabled -InternalMessage $txtInternal.Text -ExternalMessage $txtExternal.Text -ExternalAudience All
    } -SuccessMessage 'Out of Office enabled.'
})

$btnDisableOOF.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        Set-MailboxAutoReplyConfiguration -Identity $txtOOFMailbox.Text.Trim() -AutoReplyState Disabled
    } -SuccessMessage 'Out of Office disabled.'
})

$form.Add_FormClosing({
    if ($script:IsConnected) {
        try {
            Disconnect-ExchangeOnline -Confirm:$false
        }
        catch {
            # ignore disconnect errors during shutdown
        }
    }
})

Write-Log -Message 'Application loaded. Ready for connection.' -Level Info
[void]$form.ShowDialog()
