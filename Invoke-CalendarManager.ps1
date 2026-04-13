[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Email address of the calendar owner")]
    [string]$CalendarOwner,

    [Parameter(HelpMessage = "Action to perform: Default, User, or Remove")]
    [ValidateSet("Default", "User", "Remove")]
    [string]$Action,

    [Parameter(HelpMessage = "Target user email (required for User/Remove actions)")]
    [string]$User,

    [Parameter(HelpMessage = "Permission level to set")]
    [ValidateSet("Owner", "PublishingEditor", "Editor", "PublishingAuthor",
                 "Author", "NonEditingAuthor", "Reviewer", "Contributor",
                 "AvailabilityOnly", "LimitedDetails")]
    [string]$PermissionLevel
)

# Dot-source shared helpers
. "$PSScriptRoot\_Shared.ps1"

function Read-ActionChoice {
    do {
        Write-Host ""
        Write-Host "Choose an action:" -ForegroundColor Cyan
        Write-Host "  1. Modify default permissions"
        Write-Host "  2. Update permissions for a specific user"
        Write-Host "  3. Remove a custom permission for a specific user"

        $PermissionActionChoice = Read-Host -Prompt "Enter the number corresponding to the action you want to perform"

        $PermissionAction = switch ($PermissionActionChoice) {
            "1" { "default" }
            "2" { "specific" }
            "3" { "remove" }
            default {
                Write-Status -Level Error -Message "Invalid selection. Please enter 1, 2, or 3."
                $null
            }
        }
    } while (-not $PermissionAction)

    return $PermissionAction
}

function Read-PermissionLevel {
    $permissionLevels = [ordered]@{
        "1"  = "Owner"
        "2"  = "PublishingEditor"
        "3"  = "Editor"
        "4"  = "PublishingAuthor"
        "5"  = "Author"
        "6"  = "NonEditingAuthor"
        "7"  = "Reviewer"
        "8"  = "Contributor"
        "9"  = "AvailabilityOnly"
        "10" = "LimitedDetails"
    }

    do {
        Write-Host ""
        Write-Host "Please choose the permission level you'd like to set:" -ForegroundColor Cyan
        foreach ($k in $permissionLevels.Keys) {
            Write-Host "  $($k.PadLeft(2)). $($permissionLevels[$k])"
        }
        $selection = Read-Host -Prompt "Enter the number corresponding to the permission level"
    } until ($permissionLevels.Contains($selection))

    return $permissionLevels[$selection]
}

function Get-CalendarPermission {
    param(
        [string]$CalendarOwner
    )
    Write-Host ""
    try {
        Get-MailboxFolderPermission -Identity "${CalendarOwner}:\Calendar" -ErrorAction Stop
    }
    catch {
        Write-Status -Level Error -Message "Failed to retrieve permissions for ${CalendarOwner}: $_"
        return
    }
    Write-Host ""
}

function Set-CalendarPermission {
    param(
        [string]$Email,
        [string]$User,
        [string]$Access
    )

    # Check if the permission already exists
    try {
        $existingPermission = Get-MailboxFolderPermission -Identity "${Email}:\Calendar" -User $User -ErrorAction Stop
    }
    catch {
        Write-Status -Level Error -Message "Failed to retrieve existing permission for ${User}: $_"
        return
    }

    if ($null -eq $existingPermission) {
        try {
            Add-MailboxFolderPermission -Identity "${Email}:\Calendar" -User $User -AccessRights $Access -ErrorAction Stop
            Write-Status -Level OK -Message "Added '$Access' permission for $User on ${Email}'s calendar."
        }
        catch {
            Write-Status -Level Error -Message "Failed to add permission for ${User}: $_"
        }
    } else {
        try {
            Set-MailboxFolderPermission -Identity "${Email}:\Calendar" -User $User -AccessRights $Access -ErrorAction Stop
            Write-Status -Level OK -Message "Updated permission for $User to '$Access' on ${Email}'s calendar."
        }
        catch {
            Write-Status -Level Error -Message "Failed to update permission for ${User}: $_"
        }
    }
}

# --- Main ---

Write-Header "O365 Calendar Manager"
Test-ScriptVersion
Connect-ExchangeSession

# Map CLI -Action values to internal action codes
$actionMap = @{ "Default" = "default"; "User" = "specific"; "Remove" = "remove" }

# Determine if running in single-shot mode (all required params provided via CLI)
$singleShot = $false
if ($CalendarOwner -and $Action) {
    $PermissionAction = $actionMap[$Action]

    if ($Action -eq "User" -and -not $User) {
        Write-Status -Level Error -Message "-User parameter is required when -Action is 'User'."
        Invoke-SafeDisconnect
        exit 1
    }
    if ($Action -eq "Remove" -and -not $User) {
        Write-Status -Level Error -Message "-User parameter is required when -Action is 'Remove'."
        Invoke-SafeDisconnect
        exit 1
    }
    if ($Action -in @("Default", "User") -and -not $PermissionLevel) {
        Write-Status -Level Error -Message "-PermissionLevel parameter is required when -Action is '$Action'."
        Invoke-SafeDisconnect
        exit 1
    }

    $UserToModifyPermission = $User
    $PermissionLevelText = $PermissionLevel
    $singleShot = $true
}

do {
    # Prompt for values not provided via CLI
    if (-not $CalendarOwner) {
        Write-Host ""
        $CalendarOwner = Read-Host -Prompt "Please enter the email address of the calendar owner you want to view or modify"
    }

    Get-CalendarPermission $CalendarOwner

    if (-not $singleShot) {
        $PermissionAction = Read-ActionChoice

        if ($PermissionAction -eq 'specific') {
            $UserToModifyPermission = Read-Host -Prompt "Please enter the email address of the user permission you want to add or modify"
        }

        if ($PermissionAction -eq 'specific' -or $PermissionAction -eq 'default') {
            $PermissionLevelText = Read-PermissionLevel
        }

        if ($PermissionAction -eq 'remove') {
            $UserToModifyPermission = Read-Host -Prompt "Please enter the email address of the user permission you want to remove"
        }
    }

    # Show confirmation summary
    Write-Host ""
    Write-Host "=== Pending Change ===" -ForegroundColor Cyan
    Write-Host "  Calendar:    $CalendarOwner" -ForegroundColor White

    switch ($PermissionAction) {
        'default' {
            Write-Host "  Action:      Set default permission" -ForegroundColor White
            Write-Host "  Permission:  $PermissionLevelText" -ForegroundColor White
        }
        'specific' {
            Write-Host "  Action:      Set user permission" -ForegroundColor White
            Write-Host "  User:        $UserToModifyPermission" -ForegroundColor White
            Write-Host "  Permission:  $PermissionLevelText" -ForegroundColor White
        }
        'remove' {
            Write-Host "  Action:      Remove custom permission" -ForegroundColor White
            Write-Host "  User:        $UserToModifyPermission" -ForegroundColor Yellow
            Write-Status -Level Warning -Message "This will remove all custom permissions for this user."
        }
    }

    Write-Host ""
    $confirm = Read-Host -Prompt "Apply this change? (y/N)"

    if ($confirm -notin @('y', 'yes', 'Y')) {
        Write-Status -Level Info -Message "No changes made."
    } else {
        if ($PermissionAction -eq 'default') {
            try {
                Set-MailboxFolderPermission -Identity "${CalendarOwner}:\Calendar" -User Default -AccessRights $PermissionLevelText -ErrorAction Stop
                Write-Status -Level OK -Message "Set default permission to '$PermissionLevelText' on ${CalendarOwner}'s calendar."
            }
            catch {
                Write-Status -Level Error -Message "Failed to set default permissions for ${CalendarOwner}: $_"
            }
        }

        if ($PermissionAction -eq 'specific') {
            Set-CalendarPermission -Email $CalendarOwner -User $UserToModifyPermission -Access $PermissionLevelText
        }

        if ($PermissionAction -eq 'remove') {
            try {
                Remove-MailboxFolderPermission -Identity "${CalendarOwner}:\Calendar" -User $UserToModifyPermission -ErrorAction Stop
                Write-Status -Level OK -Message "Removed custom permission for $UserToModifyPermission on ${CalendarOwner}'s calendar."
            }
            catch {
                Write-Status -Level Error -Message "Failed to remove permission for ${UserToModifyPermission}: $_"
            }
        }
    }

    Get-CalendarPermission $CalendarOwner

    # In single-shot mode, exit after one operation
    if ($singleShot) { break }

    # Reset for next iteration
    $CalendarOwner = $null

    $MakeMoreChanges = Read-Host -Prompt "Do you want to make more changes? [Y/N]"
} while ($MakeMoreChanges -in @('yes', 'Y', 'y'))

Invoke-SafeDisconnect
