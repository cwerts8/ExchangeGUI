[exchange_admin_tool_read_me.md](https://github.com/user-attachments/files/22799985/exchange_admin_tool_read_me.md)
# Exchange Online Management Tool (WPF PowerShell GUI)

A full-featured PowerShell GUI for managing **Exchange Online** mailbox and calendar permissions, as well as group memberships. Built with WPF for a modern, interactive interface.

---

## Features

### Mailbox Permissions
- View, add, or remove **Full Access** and **Send As** permissions.
- Automatically filters inherited or system entries for clarity.
- Integrates with Exchange Online cmdlets:
  - `Get-MailboxPermission`
  - `Add-MailboxPermission`
  - `Remove-MailboxPermission`
  - `Get-RecipientPermission`
  - `Add-RecipientPermission`
  - `Remove-RecipientPermission`

### Calendar Permissions
- Manage permissions on users' **\Calendar** folders.
- Supports roles like **Owner**, **Editor**, **Reviewer**, **Author**, **LimitedDetails**, and **AvailabilityOnly**.
- Easily add, edit, or remove sharing permissions.

### Group Memberships
- Load and view members of any Exchange group.
- Double-click a user or group to view detailed properties.

### AD-Style Properties Viewer
- Comprehensive view of selected object attributes:
  - **General**: Display Name, Primary SMTP, UPN, GUID, Recipient Type.
  - **Contact**: Phone, Company, Office, Address.
  - **Organization**: Department, Title, Manager.
  - **Group Info**: Members, Owners, ManagedBy.
- Automatically resolves GUIDs to friendly names when possible.

### Excel Export
- Exports current data views (e.g., mailbox permissions) to timestamped Excel files.
- Uses the `ImportExcel` module for fast, dependency-free export.
- Prompts for save location and can automatically open the file post-export.

### Connection Handling
- Uses `Connect-ExchangeOnline` and `Disconnect-ExchangeOnline` for session management.
- Detects connection state and displays current user in the header.
- Configures TLS 1.2 and honors corporate proxy settings.

---

## Requirements
- **Windows PowerShell 5.1** (WPF not supported in PowerShell 7+)
- Modules:
  - `ExchangeOnlineManagement`
  - `ImportExcel`

Install prerequisites:
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module ImportExcel -Scope CurrentUser -Force
```

---

## Usage
Run the tool in STA mode (required for WPF):
```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\Exchange-AdminTool.ps1
```

1. Click **Connect** to authenticate with Exchange Online.
2. Select a tab: **Mailbox**, **Calendar**, or **Groups**.
3. Enter a target mailbox or group name.
4. View, add, or remove permissions as needed.
5. Export results to Excel for documentation.
6. Click **Disconnect** when finished.

---

## Troubleshooting
- **STA Error**: Always launch PowerShell with `-STA`. The script will attempt to relaunch automatically if needed.
- **Missing Modules**: Run the installation commands listed above.
- **Insufficient Permissions**: Ensure your account has the necessary Exchange Online admin roles.
- **Proxy Issues**: The script preconfigures default web proxy credentials.

---

## Security
- Executes actions only under your signed-in account context.
- Does not store credentials or log sensitive data.
- All Exchange modifications follow standard RBAC rules.

---

## Roadmap
- Bulk import operations from CSV
- Send-on-behalf support
- Custom folder path selection beyond `\Calendar`
- Enhanced logging and change tracking

---

## Author
**Craig Werts**  
Senior Desktop Engineer – Geller & Company  

PowerShell | WPF | Exchange Online | Automation

---

## License
MIT License

