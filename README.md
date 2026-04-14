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
- Auto-install missing dependencies on first run (with prompt, or `-AutoInstall` to skip)
- Cross-platform (Windows, macOS, Linux with PowerShell 7+)

## Prerequisites

- **PowerShell** 5.1+ (Windows) or 7+ (cross-platform)
- **ExchangeOnlineManagement** module
- Exchange Online administrator or delegate permissions
- **Recipient Management** role group membership (required to manage other users' calendars — see [Troubleshooting](#mailbox-identity-error) if you get "mailbox doesn't exist" errors)

### Installing Prerequisites

Run the included installer:

```powershell
# Run the installer (from PowerShell)
./setup/Install-Prerequisites.ps1

# Or install the module manually
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

> **Installing PowerShell itself:** See [Microsoft's installation guide](https://learn.microsoft.com/powershell/scripting/install/installing-powershell) for your platform. On Ubuntu/Debian, PowerShell is available via Microsoft's apt repository.

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

# Auto-install the ExchangeOnlineManagement module if it's missing (skips the prompt)
./Invoke-CalendarManager.ps1 -AutoInstall
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
o365-calendar-manager/
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
| `_Shared.ps1` not found | Run the script from the project root directory |
| "The specified mailbox Identity doesn't exist" | See **Mailbox Identity Error** section below |
| "The user is either not a valid SMTP address, or there is no matching information" | See **Duplicate Recipient Error** section below |

### Mailbox Identity Error

If the script works for your own calendar but fails for other users' calendars with:

```
The specified mailbox Identity:"user@domain.com" doesn't exist.
```

This is **not** a missing mailbox — it's a misleading error caused by insufficient Exchange RBAC roles. The `Get-MailboxFolderPermission` cmdlet requires membership in the **Recipient Management** role group to query other users' calendars. Global Admin alone is not sufficient.

**Diagnosis:**

```powershell
# Check if you're in Recipient Management
Get-RoleGroupMember "Recipient Management"
```

If your account is not listed, that's the problem.

**Fix:**

```powershell
# Add yourself (requires Exchange admin privileges)
Add-RoleGroupMember "Recipient Management" -Member you@yourdomain.com

# Verify
Get-RoleGroupMember "Recipient Management"

# Disconnect and reconnect (the token must refresh to pick up the new role)
Disconnect-ExchangeOnline -Confirm:$false
Connect-ExchangeOnline -UserPrincipalName you@yourdomain.com
```

> **Note:** Role changes can take 15–30 minutes to propagate. If it still fails after reconnecting, wait and try again.

### Duplicate Recipient Error

If adding a new user fails with:

```
The user "user@domain.com" is either not valid SMTP address, or there is no matching information.
```

…but you have confirmed the user exists and the email is correct, there is likely a **duplicate recipient** in Exchange. This commonly happens when a user was offboarded (their mailbox converted to a shared mailbox) and later rehired (a new user mailbox created with the same or overlapping SMTP address).

**Diagnosis:**

```powershell
# Search for all recipients matching the user
Get-Recipient -ResultSize Unlimited | Where-Object { $_.DisplayName -like "*FirstName*" } | Select-Object DisplayName, PrimarySmtpAddress, Alias, RecipientTypeDetails, ExchangeObjectId
```

If you see two recipients (typically one `SharedMailbox` and one `UserMailbox`) with the same or conflicting email addresses, that's the problem.

**Fix:** Clean up the duplicate. The cleanest approach depends on which mailbox contains the user's current data:

- If the user has been working out of the **shared mailbox** post-rehire (common when offboarding was never fully reversed): convert it back to a user mailbox and remove the empty new one
  ```powershell
  # Remove the empty duplicate first
  Remove-Mailbox -Identity <empty-mailbox-guid> -Confirm:$false

  # Convert the shared mailbox back to a regular user mailbox
  Set-Mailbox -Identity <shared-mailbox-guid> -Type Regular

  # Then assign a license in the M365 Admin Centre
  ```
- If the **new user mailbox** is already being used and the shared mailbox is stale data: export/archive the shared content, then `Remove-Mailbox` it

Consult with whoever handled the rehire to confirm which mailbox has the current data before making changes.

## Version Check

On startup, the script compares its version against the latest copy on GitHub. If an update is available, you will see a yellow notification with a download link. This check is non-blocking — if GitHub is unreachable, the script continues normally.

## Resources

- [Exchange Online PowerShell](https://learn.microsoft.com/powershell/exchange/exchange-online-powershell)
- [ExchangeOnlineManagement module](https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
- [Calendar folder permissions](https://learn.microsoft.com/powershell/module/exchange/set-mailboxfolderpermission)

---

Version 2.0.0 | Tested on PowerShell 7.5 (Ubuntu, Windows, macOS) and Windows PowerShell 5.1
