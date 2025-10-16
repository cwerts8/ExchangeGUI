# Exchange Online Management Tool (WPF PowerShell GUI)

A full-featured PowerShell GUI for managing **Exchange Online** mailbox and calendar permissions, **Active Directory** group memberships, and **Out of Office** settings. Built with WPF for a modern, interactive interface.

**Current Version:** 2.9.0

---

## Features

### Mailbox Permissions
- View, add, or remove **Full Access** and **Send As** permissions.
- Automatically filters inherited or system entries for clarity.
- Edit existing permissions without removing and re-adding.
- AutoMapping disabled by default for Full Access grants.
- GUID resolution for AD groups in permission lists.
- Integrates with Exchange Online cmdlets:
  - `Get-MailboxPermission`
  - `Add-MailboxPermission`
  - `Remove-MailboxPermission`
  - `Get-RecipientPermission`
  - `Add-RecipientPermission`
  - `Remove-RecipientPermission`

### Calendar Permissions
- Manage permissions on users' **\Calendar** folders.
- Supports 7 permission levels: **Owner**, **Editor**, **Author**, **Contributor**, **Reviewer**, **LimitedDetails**, and **AvailabilityOnly**.
- Add, edit, or remove calendar sharing permissions.
- Export calendar permissions to Excel for reporting.

### Automatic Replies (Out of Office)
- **Rich text editor** with formatting toolbar (Bold, Italic, Underline).
- Create formatted messages without HTML knowledge - automatic HTML conversion.
- Support for **internal** and **external** automatic reply messages.
- **Three reply states**: Disabled, Enabled (always on), Scheduled (date/time range).
- Date/time picker for scheduled automatic replies.
- External audience controls (All senders or Contacts only).
- Visual status indicators with color coding (Gray=Disabled, Green=Enabled, Orange=Scheduled).
- Clear formatting button to remove all text styling.

### AD Group Members (v2.9.0 - New!)
- **Now uses Active Directory** instead of Exchange Online.
- View members of **any AD group type**:
  - Security Groups
  - Distribution Groups
  - Mail-enabled Security Groups
  - Universal, Global, and DomainLocal groups
- Works with groups that aren't mail-enabled.
- Displays comprehensive member information:
  - Display Name, Email, Object Type
  - Title, Department (for user objects)
  - Nested groups, computers, and contacts
- Search by group name, email, or SAM account name.
- Copy all member email addresses to clipboard (Outlook format).
- Export to Excel with full details.
- **No Exchange Online connection required** for this feature.

### AD-Style Properties Viewer
- Comprehensive view of selected object attributes across 4 tabs:
  - **General**: Display Name, Email, Title, Department, Office, Company, Recipient Type.
  - **Contact Information**: Phone, Mobile, Fax, Street Address, City, State, Postal Code, Country.
  - **Organization**: Manager, Direct Reports, Group Memberships (for users), Group Members (for groups).
  - **Account**: UPN, SAM Account, Distinguished Name, GUID, Creation/Modification dates.
- Automatically resolves GUIDs to friendly names when possible.
- **Double-click any user or group** in any grid to view their properties.
- Copy email address to clipboard with one click.

### Excel Export
- Export permissions and group members to professionally formatted Excel files.
- Uses the `ImportExcel` module for fast, dependency-free export.
- Timestamped filenames for easy organization.
- Auto-sized columns and filtered tables.
- Prompts for save location and can automatically open the file post-export.
- Available for:
  - Mailbox permissions
  - Calendar permissions
  - AD group members

### Connection Management
- **Optional Exchange Online connection** - GUI launches immediately.
- Connect/Disconnect buttons in the interface.
- Visual connection status indicator (Red=Disconnected, Green=Connected).
- Connection validation before accessing EXO-dependent modules.
- Console-based authentication with automatic GUI restore.
- Reconnect capability if session expires.
- Configures TLS 1.2 and honors corporate proxy settings.

---

## Requirements

### PowerShell Version
- **Windows PowerShell 5.1** (WPF not supported in PowerShell 7+)

### Required Modules
- **ExchangeOnlineManagement** (for mailbox/calendar/OOO features)
- **ActiveDirectory** (for AD Group Members feature)
  - Install via RSAT: Settings > Apps > Optional Features > RSAT: Active Directory Domain Services
- **ImportExcel** (automatically installed if missing)

### Permissions
- Exchange Online administrator role (for mailbox/calendar/OOO management)
- Active Directory read permissions (for group member viewing)

Install Exchange Online Management module:
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

Install ImportExcel module (optional - script will prompt):
```powershell
Install-Module ImportExcel -Scope CurrentUser -Force
```

---

## Usage

### Launch the Tool
Run in STA mode (required for WPF):
```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\Exchange-AdminTool.ps1
```

The script will automatically relaunch in STA mode if needed.

### Workflow

1. **GUI launches immediately** - no connection required initially.

2. **For Exchange Online features** (Mailbox/Calendar/OOO):
   - Click **Connect to Exchange Online** button.
   - Authenticate via browser (Modern Auth).
   - Green indicator shows connected status.

3. **For AD Group Members**:
   - No connection needed - works independently.
   - Enter group name, email, or SAM account.
   - View all members regardless of group type.

4. **Manage Permissions**:
   - Select a module (Mailbox, Calendar, OOO, or Groups).
   - Enter target mailbox or group identifier.
   - Add, edit, or remove permissions.
   - Double-click any user/group to view AD properties.

5. **Export Data**:
   - Click "Export to Excel" button.
   - Choose save location.
   - Optionally open the file immediately.

6. **Disconnect** when finished (EXO sessions only).

---

## Key Features by Version

### v2.9.0 (Current)
- AD Group Members now uses Active Directory module
- Supports all AD group types (Security, Distribution, etc.)
- Group members feature works without Exchange Online connection
- Added Group Scope display
- Enhanced member type support (Users, Groups, Computers, Contacts)

### v2.8.1
- Fixed permission loading bugs for single delegates
- Improved array handling for DataGrid binding

### v2.8.0
- Optional Exchange Online connection
- Visual connection status indicator
- Connect/Disconnect controls in GUI

### v2.7.0
- Automatic Replies (OOO) management module
- Rich text editor with formatting toolbar
- HTML-enabled messages without coding

### v2.6.1
- Company logo and dynamic version display

### v2.6.0
- Excel export for all modules
- GUID-to-name resolution for AD groups
- Double-click to view AD properties

---

## Troubleshooting

### STA Mode Error
Always launch PowerShell with `-STA`. The script will attempt to relaunch automatically if needed.

### Missing Active Directory Module
For AD Group Members functionality:
1. Open **Settings** > **Apps** > **Optional Features**
2. Click **Add a feature**
3. Search for and install: **RSAT: Active Directory Domain Services and Lightweight Directory Services Tools**
4. Restart PowerShell

### Missing Exchange Online Module
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

### Connection Issues
- Ensure your account has necessary Exchange Online admin roles.
- Check corporate proxy settings (configured in script lines 21-23).
- Verify Modern Authentication is enabled for your tenant.

### Insufficient Permissions
- **Exchange features**: Requires Exchange Administrator, Global Admin, or custom role with appropriate permissions.
- **AD Group Members**: Requires Active Directory read permissions (standard domain user typically sufficient).

### Performance Issues
- Large groups (1000+ members) may take time to load all details.
- Consider exporting to Excel for offline analysis of very large groups.

---

## Security

- Executes actions only under your signed-in account context.
- Does not store credentials or log sensitive data.
- All Exchange modifications follow standard RBAC rules.
- AD queries are read-only - no modification capability.
- Activity log shows all operations performed.

---

## Corporate Environment Notes

### Proxy Configuration
Update proxy settings in the script (lines 21-23) if different from default:
```powershell
[system.net.webrequest]::defaultwebproxy = new-object system.net.webproxy('http://proxy.example.com:8080')
```

### Logo Customization
Place your company logo file at:
```
$PSScriptRoot\FullColorLogo.png
```
Recommended size: 250x60 pixels (PNG format)

---

## Roadmap

- [ ] Bulk import operations from CSV
- [ ] Send-on-behalf permission support
- [ ] Custom folder path selection beyond `\Calendar`
- [ ] Enhanced logging with export capability
- [ ] Shared mailbox management features
- [ ] Distribution list management
- [ ] Mailbox size and quota reporting
- [ ] Message trace integration

---

## Author

**Craig Werts**  
Senior Desktop Engineer

PowerShell | WPF | Exchange Online | Active Directory | Automation

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 2.9.0 | 10-16-25 | AD Group Members now uses AD module; supports all group types |
| 2.8.1 | 10-15-25 | Fixed permission loading bugs for single delegates |
| 2.8.0 | 10-15-25 | Optional EXO connection; visual status indicators |
| 2.7.0 | 10-14-25 | Automatic Replies (OOO) module with rich text editor |
| 2.6.1 | 10-09-25 | Company logo and version display |
| 2.6.0 | 10-09-25 | Excel export; GUID resolution; AD properties viewer |
| 2.0.0 | 10-08-25 | Major feature additions |
| 1.0.0 | 10-06-25 | Initial release |

---

## License

MIT License

Copyright (c) 2025 Craig Werts

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

---

## Screenshots

*(Consider adding screenshots of the main interface, each module, and the AD Properties viewer)*

---

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

---

## Support

For issues or questions:
- Open an issue in the GitHub repository

---

**Note:** This tool is designed for IT administrators managing Microsoft 365 environments. Always test in a non-production environment before deploying to production systems.
