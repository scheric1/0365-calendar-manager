# O365 Calendar Manager

PowerShell tool for managing Exchange Online calendar permissions — view, modify, and remove calendar sharing permissions from the command line.

**Version 2.0.0**

## Quick Start

```powershell
# Install prerequisites
./setup/Install-Prerequisites.ps1

# Verify installation
./setup/Test-Installation.ps1

# Run interactively
./Invoke-CalendarManager.ps1

# Or with parameters (skips prompts for provided values)
./Invoke-CalendarManager.ps1 -CalendarOwner john@contoso.com -Action Default -PermissionLevel Reviewer
```

## Features

- View current calendar permissions for any mailbox
- Modify default calendar permissions (what everyone in the org sees)
- Add or update permissions for specific users
- Remove custom permissions
- Confirmation prompts before every change
- Colour-coded output for clear feedback
- CLI parameters to skip interactive prompts
- Auto-update check against the latest GitHub version
- Cross-platform (Windows, macOS, Linux with PowerShell 7+)

## Prerequisites

- **PowerShell** 5.1+ (Windows) or 7+ (cross-platform)
- **ExchangeOnlineManagement** module
- Exchange Online administrator or delegate permissions

### Installing Prerequisites

Run the included installer:

```powershell
# apt (Ubuntu/Debian — installs PowerShell)
sudo apt install powershell

# Then install the Exchange module
./setup/Install-Prerequisites.ps1

# Or install manually
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

## Usage

### Interactive Mode

Run with no parameters for the full interactive experience:

```powershell
./Invoke-CalendarManager.ps1
```

You will be prompted for:
1. Your email address (to authenticate)
2. The calendar owner's email
3. The action to perform
4. The permission level (for add/modify actions)
5. Confirmation before applying

### CLI Mode

Provide parameters to skip the corresponding prompts. Authentication is always interactive.

```powershell
# View permissions only
./Invoke-CalendarManager.ps1 -CalendarOwner john@contoso.com

# Set default calendar permission
./Invoke-CalendarManager.ps1 -CalendarOwner john@contoso.com -Action Default -PermissionLevel Reviewer

# Grant a specific user access
./Invoke-CalendarManager.ps1 -CalendarOwner john@contoso.com -Action User -User jane@contoso.com -PermissionLevel Editor

# Remove a user's custom permission
./Invoke-CalendarManager.ps1 -CalendarOwner john@contoso.com -Action Remove -User jane@contoso.com
```

**Note:** Even in CLI mode, you will be asked to confirm the change before it is applied.

## Permission Levels

| Level | Description |
|-------|-------------|
| Owner | Full control — read, create, modify, delete all items and subfolders |
| PublishingEditor | Create, read, modify, delete all items; create subfolders |
| Editor | Create, read, modify, delete all items |
| PublishingAuthor | Create and read all items; modify and delete own items; create subfolders |
| Author | Create and read all items; modify and delete own items |
| NonEditingAuthor | Create and read all items; delete own items |
| Reviewer | Read all items (view only) |
| Contributor | Create items only (cannot read) |
| AvailabilityOnly | See free/busy status only |
| LimitedDetails | See free/busy status with subject and location |

## Project Structure

```
0365-calendar-manager/
├── Invoke-CalendarManager.ps1      # Main script
├── _Shared.ps1                     # Shared helpers (auth, output, version check)
├── README.md
├── CHANGELOG.md
└── setup/
    ├── Install-Prerequisites.ps1   # Prerequisite installer
    └── Test-Installation.ps1       # Diagnostic test suite
```

## Security

- **No credentials stored** — uses interactive OAuth 2.0 authentication only
- **Read-only by default** — viewing permissions requires no confirmation
- **Confirmation required** — every permission change shows a summary and requires explicit `y` to proceed
- **No data exfiltration** — version check is the only outbound call (to GitHub raw content)

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `Connect-ExchangeOnline` not recognised | Run `./setup/Install-Prerequisites.ps1` or `Install-Module ExchangeOnlineManagement -Scope CurrentUser` |
| Untrusted repository error | `Set-PSRepository -Name PSGallery -InstallationPolicy Trusted` |
| Script blocked by execution policy | `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` |
| Authentication fails | Ensure your account has Exchange administrator or delegate permissions |
| `\_Shared.ps1` not found | Run the script from the project root directory |

## Version Check

On startup, the script compares its version against the latest copy on GitHub. If an update is available, you will see a yellow notification with a download link. This check is non-blocking — if GitHub is unreachable, the script continues normally.

## Resources

- [Exchange Online PowerShell](https://learn.microsoft.com/powershell/exchange/exchange-online-powershell)
- [ExchangeOnlineManagement module](https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
- [Calendar folder permissions](https://learn.microsoft.com/powershell/module/exchange/set-mailboxfolderpermission)

---

Version 2.0.0 | Tested on PowerShell 7.5 (Ubuntu, Windows, macOS) and Windows PowerShell 5.1
