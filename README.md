# 0365-calendar-manager

A PowerShell script for viewing and managing Exchange Online calendar permissions.

## Features

- Connects to Exchange Online using the supplied user principal name.
- Displays existing permissions for a mailbox calendar.
- Supports:
  - Modifying default calendar permissions.
  - Adding or updating permissions for a specific user.
  - Removing custom permissions for a specific user.
- Presents a menu of common permission levels (Owner, Editor, Reviewer, etc.).
- Shows updated permissions after each change and allows multiple edits per session.
- Disconnects from Exchange Online when finished.

## Prerequisites

- PowerShell 7+ or Windows PowerShell.
- [Exchange Online Management](https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell) module.
- Permission to modify the target mailbox's calendar in Exchange Online.

### Installing prerequisites on Windows

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
# If scripts are blocked:
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
```

### Installing prerequisites on macOS or Linux

1. [Install PowerShell](https://learn.microsoft.com/powershell/scripting/install/installing-powershell).
2. Install the Exchange Online module:

```powershell
pwsh -Command "Install-Module ExchangeOnlineManagement -Scope CurrentUser"
```

## Usage

Run the script from PowerShell:

```powershell
./calendarManager.ps1
# or
pwsh ./calendarManager.ps1
```

When prompted:

1. Enter your email address to authenticate.
2. Provide the calendar owner's email address to view or modify.
3. Choose an action:
   - Modify default permissions
   - Update permissions for a specific user
   - Remove a user's custom permission
4. If modifying or adding a permission, select the desired access level from the list.
5. After each change, decide whether to make more changes; type `N` to exit.

The script will display permissions before and after each change and disconnect from Exchange Online upon completion.

## Troubleshooting

- **`Connect-ExchangeOnline` not recognized**: Ensure the Exchange Online Management module is installed and imported.

  ```powershell
  Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
  Import-Module ExchangeOnlineManagement
  ```

- **Untrusted repository errors**: Allow unsigned scripts for your user

  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
  ```

- **Script blocked by execution policy**: Bulk Fix your Scripts folder

  ```powershell
  Get-ChildItem "C:\Users\username\Scripts" -Filter *.ps1 -Recurse | Unblock-File
  ```
## Version check

When the script starts it compares its version with the latest copy in the GitHub repository. If an update is available, it will notify you and provide a link to download the most recent version.

## Notes

The script requires network connectivity to Microsoft's cloud services. If you have not installed the Exchange Online module yet, run the installation commands above before executing the script.
