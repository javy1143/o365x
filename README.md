# O365 Exchange Admin Console (PowerShell WinForms)

This repository contains `ExchangeAdminTool.ps1`, a Windows Forms GUI for Exchange Online administration.

## Features
- Connect/disconnect from Exchange Online
- Mailbox permission management (Full Access, Send As, Send On Behalf)
- Calendar permission management
- Distribution group member management with dropdown of current groups
- Shared mailbox access assignment/removal
- Out-of-office (automatic replies) enable/disable
- Current permissions viewer tab for mailbox/calendar/groups/shared mailbox/OOF

## Color palette
The UI uses a custom palette based on your logo tones:
- Bright Blue: `#0089D6`
- Deep Blue: `#00609E`
- Light Blue: `#33A6E9`
- Charcoal: `#353A42`
- Slate: `#4C535C`

## Requirements
- Windows PowerShell 5.1+ (or PowerShell 7 with WinForms support on Windows)
- Exchange Online PowerShell module:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

## Run
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\ExchangeAdminTool.ps1
```

## Notes
- Use an admin account with Exchange Administrator rights.
- This tool executes Exchange cmdlets directly.
- Consider adding RBAC restrictions and audit logging for production use.
