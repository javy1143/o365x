#requires -Version 5.1
<#!
O365 Exchange Online WinForms admin console
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

$Theme = @{
    Primary      = [System.Drawing.Color]::FromArgb(0,137,214)
    PrimaryDark  = [System.Drawing.Color]::FromArgb(0,96,158)
    AccentDark   = [System.Drawing.Color]::FromArgb(53,58,66)
    Surface      = [System.Drawing.Color]::FromArgb(241,245,249)
    Text         = [System.Drawing.Color]::FromArgb(31,41,55)
    Success      = [System.Drawing.Color]::FromArgb(16,185,129)
    Warning      = [System.Drawing.Color]::FromArgb(245,158,11)
    Danger       = [System.Drawing.Color]::FromArgb(220,38,38)
}

$script:IsConnected = $false

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error')]
        [string]$Level = 'Info'
    )

    $entry = "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [$Level] $Message"
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

    $previousErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Stop'
        & $Script
        if ($SuccessMessage) { Write-Log -Message $SuccessMessage -Level Success }
    }
    catch {
        Write-Log -Message "${ErrorPrefix}: $($_.Exception.Message)" -Level Error
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
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

function Assert-Connected {
    if (-not $script:IsConnected) {
        throw 'Not connected to Exchange Online. Use the Connection tab first.'
    }
}

function New-LabeledTextBox {
    param([System.Windows.Forms.Control]$Parent,[string]$Label,[int]$X,[int]$Y,[int]$Width = 280)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label
    $lbl.Location = [System.Drawing.Point]::new($X,$Y)
    $lbl.AutoSize = $true
    $lbl.ForeColor = $Theme.Text
    $Parent.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = [System.Drawing.Point]::new($X,($Y + 20))
    $txt.Size = [System.Drawing.Size]::new($Width,24)
    $Parent.Controls.Add($txt)
    return $txt
}

function New-LabeledComboBox {
    param([System.Windows.Forms.Control]$Parent,[string]$Label,[int]$X,[int]$Y,[int]$Width = 280)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label
    $lbl.Location = [System.Drawing.Point]::new($X,$Y)
    $lbl.AutoSize = $true
    $lbl.ForeColor = $Theme.Text
    $Parent.Controls.Add($lbl)

    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.Location = [System.Drawing.Point]::new($X,($Y + 20))
    $cmb.Size = [System.Drawing.Size]::new($Width,24)
    $cmb.DropDownStyle = 'DropDownList'
    $Parent.Controls.Add($cmb)
    return $cmb
}

function New-ThemeButton {
    param([System.Windows.Forms.Control]$Parent,[string]$Text,[int]$X,[int]$Y,[int]$Width = 180,[System.Drawing.Color]$BackColor = $Theme.Primary)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = [System.Drawing.Point]::new($X,$Y)
    $btn.Size = [System.Drawing.Size]::new($Width,32)
    $btn.BackColor = $BackColor
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatStyle = 'Flat'
    $btn.FlatAppearance.BorderSize = 0
    $btn.Cursor = 'Hand'
    $Parent.Controls.Add($btn)
    return $btn
}

function New-ResultGrid {
    param([System.Windows.Forms.Control]$Parent,[int]$X,[int]$Y,[int]$Width = 980,[int]$Height = 380)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = [System.Drawing.Point]::new($X,$Y)
    $grid.Size = [System.Drawing.Size]::new($Width,$Height)
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $Parent.Controls.Add($grid)
    return $grid
}

function ConvertTo-DataTable {
    param([object[]]$InputObject)
    $table = New-Object System.Data.DataTable
    if (-not $InputObject -or $InputObject.Count -eq 0) { return $table }

    $properties = $InputObject[0].PSObject.Properties.Name
    foreach ($name in $properties) { [void]$table.Columns.Add($name) }

    foreach ($item in $InputObject) {
        $row = $table.NewRow()
        foreach ($name in $properties) {
            $v = $item.$name
            $row[$name] = if ($null -eq $v) { '' } else { [string]$v }
        }
        [void]$table.Rows.Add($row)
    }

    return $table
}

function Set-ResultGrid {
    param([System.Windows.Forms.DataGridView]$Grid,[object[]]$Data,[string]$NoDataMessage = 'No data returned.')

    $Grid.DataSource = $null
    $Grid.Columns.Clear()
    $Grid.AutoGenerateColumns = $true

    if (-not $Data -or $Data.Count -eq 0) {
        $Grid.DataSource = ConvertTo-DataTable -InputObject @([pscustomobject]@{ Result = $NoDataMessage })
        return
    }

    $Grid.DataSource = ConvertTo-DataTable -InputObject $Data
    $Grid.Refresh()
}

function Set-ComboValues {
    param([System.Windows.Forms.ComboBox]$Combo,[string[]]$Values)
    $Combo.Items.Clear()
    if ($Values) {
        [void]$Combo.Items.AddRange($Values)
        if ($Combo.Items.Count -gt 0) { $Combo.SelectedIndex = 0 }
    }
}

function Refresh-DistributionGroups {
    Assert-Connected
    $groups = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop | Sort-Object DisplayName
    Set-ComboValues -Combo $cmbGroupName -Values @($groups | ForEach-Object { $_.PrimarySmtpAddress.ToString() })
}

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = 'O365 Exchange Admin Console'
$form.Size = [System.Drawing.Size]::new(1080,760)
$form.StartPosition = 'CenterScreen'
$form.BackColor = $Theme.Surface
$form.Font = New-Object System.Drawing.Font('Segoe UI',9)

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Top'
$headerPanel.Height = 58
$headerPanel.BackColor = $Theme.PrimaryDark
$form.Controls.Add($headerPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'Office 365 Exchange Administration Tool'
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold',14)
$titleLabel.AutoSize = $true
$titleLabel.Location = [System.Drawing.Point]::new(16,14)
$headerPanel.Controls.Add($titleLabel)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = [System.Drawing.Color]::White
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = 'Ready. Connect to Exchange Online to begin.'
$statusLabel.ForeColor = $Theme.Text
[void]$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = [System.Drawing.Point]::new(10,68)
$tabs.Size = [System.Drawing.Size]::new(1044,600)
$form.Controls.Add($tabs)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.Location = [System.Drawing.Point]::new(10,670)
$txtLog.Size = [System.Drawing.Size]::new(1044,50)
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.ForeColor = $Theme.Text
$form.Controls.Add($txtLog)

# --- Connection tab ---
$tabConnection = New-Object System.Windows.Forms.TabPage
$tabConnection.Text = 'Connection'
$tabConnection.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabConnection)

$txtAdminUPN = New-LabeledTextBox -Parent $tabConnection -Label 'Admin UPN' -X 20 -Y 20 -Width 320
$btnConnect = New-ThemeButton -Parent $tabConnection -Text 'Connect Exchange Online' -X 20 -Y 80 -Width 220
$btnDisconnect = New-ThemeButton -Parent $tabConnection -Text 'Disconnect' -X 250 -Y 80 -Width 120 -BackColor $Theme.AccentDark

# --- Mailbox Permissions tab ---
$tabPermissions = New-Object System.Windows.Forms.TabPage
$tabPermissions.Text = 'Mailbox Permissions'
$tabPermissions.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabPermissions)

$txtPermMailbox = New-LabeledTextBox -Parent $tabPermissions -Label 'Target Mailbox (UPN or alias)' -X 20 -Y 20
$txtPermUser = New-LabeledTextBox -Parent $tabPermissions -Label 'Delegate User (UPN)' -X 330 -Y 20

$cmbPermType = New-Object System.Windows.Forms.ComboBox
$cmbPermType.Location = [System.Drawing.Point]::new(640,40)
$cmbPermType.Size = [System.Drawing.Size]::new(180,24)
$cmbPermType.DropDownStyle = 'DropDownList'
$cmbPermType.Items.AddRange(@('FullAccess','SendAs','SendOnBehalf'))
$cmbPermType.SelectedIndex = 0
$tabPermissions.Controls.Add($cmbPermType)

$lblPermType = New-Object System.Windows.Forms.Label
$lblPermType.Text = 'Permission Type'
$lblPermType.Location = [System.Drawing.Point]::new(640,20)
$lblPermType.AutoSize = $true
$tabPermissions.Controls.Add($lblPermType)

$btnGrantPerm = New-ThemeButton -Parent $tabPermissions -Text 'Grant Permission' -X 20 -Y 90
$btnRemovePerm = New-ThemeButton -Parent $tabPermissions -Text 'Remove Permission' -X 210 -Y 90 -BackColor $Theme.AccentDark
$btnViewPerms = New-ThemeButton -Parent $tabPermissions -Text 'View Current Permissions' -X 400 -Y 90 -Width 210 -BackColor $Theme.PrimaryDark
$gridPerms = New-ResultGrid -Parent $tabPermissions -X 20 -Y 140

# --- Calendar tab ---
$tabCalendar = New-Object System.Windows.Forms.TabPage
$tabCalendar.Text = 'Calendars'
$tabCalendar.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabCalendar)

$txtCalMailbox = New-LabeledTextBox -Parent $tabCalendar -Label 'Mailbox (owner)' -X 20 -Y 20
$txtCalUser = New-LabeledTextBox -Parent $tabCalendar -Label 'User to grant/revoke' -X 330 -Y 20

$cmbCalAccess = New-Object System.Windows.Forms.ComboBox
$cmbCalAccess.Location = [System.Drawing.Point]::new(640,40)
$cmbCalAccess.Size = [System.Drawing.Size]::new(180,24)
$cmbCalAccess.DropDownStyle = 'DropDownList'
$cmbCalAccess.Items.AddRange(@('Reviewer','Editor','PublishingEditor','AvailabilityOnly','LimitedDetails'))
$cmbCalAccess.SelectedIndex = 0
$tabCalendar.Controls.Add($cmbCalAccess)

$lblCalAccess = New-Object System.Windows.Forms.Label
$lblCalAccess.Text = 'Access Right'
$lblCalAccess.Location = [System.Drawing.Point]::new(640,20)
$lblCalAccess.AutoSize = $true
$tabCalendar.Controls.Add($lblCalAccess)

$btnAddCal = New-ThemeButton -Parent $tabCalendar -Text 'Set Calendar Permission' -X 20 -Y 90
$btnRemoveCal = New-ThemeButton -Parent $tabCalendar -Text 'Remove Calendar Permission' -X 210 -Y 90 -Width 210 -BackColor $Theme.AccentDark
$btnViewCalPerms = New-ThemeButton -Parent $tabCalendar -Text 'View Current Calendar Access' -X 430 -Y 90 -Width 230 -BackColor $Theme.PrimaryDark
$gridCal = New-ResultGrid -Parent $tabCalendar -X 20 -Y 140

# --- Groups tab ---
$tabGroups = New-Object System.Windows.Forms.TabPage
$tabGroups.Text = 'Email Groups'
$tabGroups.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabGroups)

$cmbGroupName = New-LabeledComboBox -Parent $tabGroups -Label 'Distribution Group' -X 20 -Y 20
$txtGroupMember = New-LabeledTextBox -Parent $tabGroups -Label 'User to add/remove (UPN)' -X 330 -Y 20
$btnRefreshGroups = New-ThemeButton -Parent $tabGroups -Text 'Load Current Groups' -X 20 -Y 90 -Width 180
$btnGroupAdd = New-ThemeButton -Parent $tabGroups -Text 'Add Member' -X 210 -Y 90
$btnGroupRemove = New-ThemeButton -Parent $tabGroups -Text 'Remove Member' -X 400 -Y 90 -BackColor $Theme.AccentDark
$btnViewGroupMembers = New-ThemeButton -Parent $tabGroups -Text 'View Current Members' -X 590 -Y 90 -Width 190 -BackColor $Theme.PrimaryDark
$gridGroups = New-ResultGrid -Parent $tabGroups -X 20 -Y 140

# --- Shared tab ---
$tabShared = New-Object System.Windows.Forms.TabPage
$tabShared.Text = 'Shared Mailboxes'
$tabShared.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabShared)

$txtSharedMailbox = New-LabeledTextBox -Parent $tabShared -Label 'Shared Mailbox (UPN or alias)' -X 20 -Y 20
$txtSharedUser = New-LabeledTextBox -Parent $tabShared -Label 'User (UPN)' -X 330 -Y 20
$btnSharedGrant = New-ThemeButton -Parent $tabShared -Text 'Grant Full + SendAs' -X 20 -Y 90
$btnSharedRevoke = New-ThemeButton -Parent $tabShared -Text 'Revoke Full + SendAs' -X 210 -Y 90 -BackColor $Theme.AccentDark
$btnViewSharedPerms = New-ThemeButton -Parent $tabShared -Text 'View Current Access' -X 400 -Y 90 -Width 180 -BackColor $Theme.PrimaryDark
$gridShared = New-ResultGrid -Parent $tabShared -X 20 -Y 140

# --- OOF tab ---
$tabOOF = New-Object System.Windows.Forms.TabPage
$tabOOF.Text = 'Out of Office'
$tabOOF.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabOOF)

$txtOOFMailbox = New-LabeledTextBox -Parent $tabOOF -Label 'Mailbox (UPN)' -X 20 -Y 20
$lblInternal = New-Object System.Windows.Forms.Label
$lblInternal.Text = 'Internal Message'
$lblInternal.Location = [System.Drawing.Point]::new(20,70)
$lblInternal.AutoSize = $true
$tabOOF.Controls.Add($lblInternal)

$txtInternal = New-Object System.Windows.Forms.TextBox
$txtInternal.Multiline = $true
$txtInternal.Location = [System.Drawing.Point]::new(20,90)
$txtInternal.Size = [System.Drawing.Size]::new(460,90)
$tabOOF.Controls.Add($txtInternal)

$lblExternal = New-Object System.Windows.Forms.Label
$lblExternal.Text = 'External Message'
$lblExternal.Location = [System.Drawing.Point]::new(500,70)
$lblExternal.AutoSize = $true
$tabOOF.Controls.Add($lblExternal)

$txtExternal = New-Object System.Windows.Forms.TextBox
$txtExternal.Multiline = $true
$txtExternal.Location = [System.Drawing.Point]::new(500,90)
$txtExternal.Size = [System.Drawing.Size]::new(460,90)
$tabOOF.Controls.Add($txtExternal)

$btnEnableOOF = New-ThemeButton -Parent $tabOOF -Text 'Enable OOF' -X 20 -Y 200
$btnDisableOOF = New-ThemeButton -Parent $tabOOF -Text 'Disable OOF' -X 210 -Y 200 -BackColor $Theme.AccentDark
$btnViewOOF = New-ThemeButton -Parent $tabOOF -Text 'View Current OOF Status' -X 390 -Y 200 -Width 190 -BackColor $Theme.PrimaryDark
$gridOOF = New-ResultGrid -Parent $tabOOF -X 20 -Y 250 -Height 270

# --- Events ---
$btnConnect.Add_Click({
    if (-not (Ensure-ExchangeModule)) { return }
    $adminUpn = $txtAdminUPN.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($adminUpn)) { Write-Log -Message 'Enter an admin UPN before connecting.' -Level Warning; return }

    global:InvokeSafely -Script {
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            Write-Log -Message 'Connecting in Windows PowerShell mode (UseWebLogin)...' -Level Info
            Connect-ExchangeOnline -UserPrincipalName $adminUpn -UseWebLogin -ShowBanner:$false
        }
        else {
            Connect-ExchangeOnline -UserPrincipalName $adminUpn -ShowBanner:$false
        }
        $script:IsConnected = $true
        Refresh-DistributionGroups
    } -SuccessMessage "Connected to Exchange Online as $adminUpn" -ErrorPrefix 'Connection failed'
})

$btnDisconnect.Add_Click({
    global:InvokeSafely -Script {
        Disconnect-ExchangeOnline -Confirm:$false
        $script:IsConnected = $false
        $cmbGroupName.Items.Clear()
    } -SuccessMessage 'Disconnected from Exchange Online.' -ErrorPrefix 'Disconnect failed'
})

$btnGrantPerm.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtPermMailbox.Text.Trim(); $user = $txtPermUser.Text.Trim(); $perm = $cmbPermType.SelectedItem
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
        $mailbox = $txtPermMailbox.Text.Trim(); $user = $txtPermUser.Text.Trim(); $perm = $cmbPermType.SelectedItem
        switch ($perm) {
            'FullAccess'   { Remove-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -Confirm:$false }
            'SendAs'       { Remove-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false }
            'SendOnBehalf' { Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo @{Remove=$user} }
        }
    } -SuccessMessage 'Permission removed successfully.'
})

$btnViewPerms.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtPermMailbox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { throw 'Enter a target mailbox first.' }

        $rows = @()
        $rows += Get-MailboxPermission -Identity $mailbox -ErrorAction Stop |
            Where-Object { -not $_.IsInherited -and $_.User -ne 'NT AUTHORITY\SELF' } |
            ForEach-Object { [pscustomobject]@{ Type='Mailbox'; User=$_.User; Rights=($_.AccessRights -join ','); Inherited=$_.IsInherited } }

        $rows += Get-RecipientPermission -Identity $mailbox -ErrorAction Stop |
            ForEach-Object { [pscustomobject]@{ Type='Recipient'; User=$_.Trustee; Rights=($_.AccessRights -join ','); Inherited='' } }

        $sendOnBehalf = (Get-Mailbox -Identity $mailbox -ErrorAction Stop).GrantSendOnBehalfTo
        foreach ($entry in $sendOnBehalf) {
            $rows += [pscustomobject]@{ Type='SendOnBehalf'; User=$entry; Rights='SendOnBehalf'; Inherited='' }
        }

        Set-ResultGrid -Grid $gridPerms -Data $rows
    } -SuccessMessage 'Current mailbox permissions loaded.'
})

$btnAddCal.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        Add-MailboxFolderPermission -Identity "$($txtCalMailbox.Text.Trim())`:\Calendar" -User $txtCalUser.Text.Trim() -AccessRights $cmbCalAccess.SelectedItem
    } -SuccessMessage 'Calendar permission updated.'
})

$btnRemoveCal.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        Remove-MailboxFolderPermission -Identity "$($txtCalMailbox.Text.Trim())`:\Calendar" -User $txtCalUser.Text.Trim() -Confirm:$false
    } -SuccessMessage 'Calendar permission removed.'
})

$btnViewCalPerms.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtCalMailbox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { throw 'Enter mailbox owner first.' }
        $data = Get-MailboxFolderPermission -Identity "$mailbox`:\Calendar" -ErrorAction Stop | Select-Object User,AccessRights,SharingPermissionFlags
        Set-ResultGrid -Grid $gridCal -Data $data
    } -SuccessMessage 'Current calendar permissions loaded.'
})

$btnRefreshGroups.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        Refresh-DistributionGroups
    } -SuccessMessage 'Current distribution groups loaded.'
})

$btnGroupAdd.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $groupIdentity = $cmbGroupName.SelectedItem
        if (-not $groupIdentity) { throw 'Select a distribution group first.' }
        Add-DistributionGroupMember -Identity $groupIdentity -Member $txtGroupMember.Text.Trim()
    } -SuccessMessage 'Group member added.'
})

$btnGroupRemove.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $groupIdentity = $cmbGroupName.SelectedItem
        if (-not $groupIdentity) { throw 'Select a distribution group first.' }
        Remove-DistributionGroupMember -Identity $groupIdentity -Member $txtGroupMember.Text.Trim() -Confirm:$false
    } -SuccessMessage 'Group member removed.'
})

$btnViewGroupMembers.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $groupIdentity = $cmbGroupName.SelectedItem
        if (-not $groupIdentity) { throw 'Select a distribution group first.' }
        $data = Get-DistributionGroupMember -Identity $groupIdentity -ResultSize Unlimited -ErrorAction Stop | Select-Object Name,PrimarySmtpAddress,RecipientType
        Set-ResultGrid -Grid $gridGroups -Data $data
    } -SuccessMessage 'Current group membership loaded.'
})

$btnSharedGrant.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtSharedMailbox.Text.Trim(); $user = $txtSharedUser.Text.Trim()
        Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -AutoMapping:$true
        Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false
    } -SuccessMessage 'Shared mailbox access granted.'
})

$btnSharedRevoke.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtSharedMailbox.Text.Trim(); $user = $txtSharedUser.Text.Trim()
        Remove-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -Confirm:$false
        Remove-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false
    } -SuccessMessage 'Shared mailbox access revoked.'
})

$btnViewSharedPerms.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtSharedMailbox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { throw 'Enter a shared mailbox first.' }
        $rows = @()
        $rows += Get-MailboxPermission -Identity $mailbox -ErrorAction Stop |
            Where-Object { -not $_.IsInherited -and $_.User -ne 'NT AUTHORITY\SELF' } |
            Select-Object @{n='Type';e={'Mailbox'}},User,@{n='Rights';e={$_.AccessRights -join ','}}
        $rows += Get-RecipientPermission -Identity $mailbox -ErrorAction Stop |
            Select-Object @{n='Type';e={'Recipient'}},@{n='User';e={$_.Trustee}},@{n='Rights';e={$_.AccessRights -join ','}}
        Set-ResultGrid -Grid $gridShared -Data $rows
    } -SuccessMessage 'Current shared mailbox access loaded.'
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

$btnViewOOF.Add_Click({
    global:InvokeSafely -Script {
        Assert-Connected
        $mailbox = $txtOOFMailbox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { throw 'Enter a mailbox first.' }
        $cfg = Get-MailboxAutoReplyConfiguration -Identity $mailbox -ErrorAction Stop
        $data = @([pscustomobject]@{
            Identity         = $cfg.Identity
            AutoReplyState   = $cfg.AutoReplyState
            ExternalAudience = $cfg.ExternalAudience
            StartTime        = $cfg.StartTime
            EndTime          = $cfg.EndTime
            InternalMessage  = $cfg.InternalMessage
            ExternalMessage  = $cfg.ExternalMessage
        })
        Set-ResultGrid -Grid $gridOOF -Data $data
    } -SuccessMessage 'Current OOF configuration loaded.'
})

$form.Add_FormClosing({
    if ($script:IsConnected) {
        try { Disconnect-ExchangeOnline -Confirm:$false } catch { }
    }
})

Write-Log -Message 'Application loaded. Ready for connection.' -Level Info
[void]$form.ShowDialog()
