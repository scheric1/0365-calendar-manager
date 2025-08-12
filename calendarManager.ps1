function User-Login {
    # Prompt for user's email
    $UserEmail = Read-Host -Prompt "Please enter your email address to login"

    # Connect to Microsoft Exchange
    Connect-ExchangeOnline -UserPrincipalName $UserEmail
}

function Set-Action {
    do {
        # Ask if they want to do
        Write-Host ""
        Write-Host "Choose an action:"
        Write-Host "1. Modify default permissions"
        Write-Host "2. Update permissions for a specific user"
        Write-Host "3. Remove a custom permission for a specific user"

        $PermissionActionChoice = Read-Host -Prompt "Enter the number corresponding to the action you want to perform"

        $PermissionAction = switch ($PermissionActionChoice) {
            "1" { "default" }
            "2" { "specific" }
            "3" { "remove" }
            default {
                Write-Host "Invalid selection. Please enter 1, 2, or 3." -ForegroundColor Red
                $null
            }
        }
    } while (-not $PermissionAction)

    return $PermissionAction
}

function Get-Permissions {
    param(
        [string]$CalendarOwner
    )
    Write-Host ""
    Get-MailboxFolderPermission -Identity "${CalendarOwner}:\Calendar"
    Write-Host ""
}

function Add-Or-Update-Permission {
    param(
        [string]$Email,
        [string]$User,
        [string]$Access
    )

    # Check if the permission exists
        $existingPermission = Get-MailboxFolderPermission -Identity "${Email}:\Calendar" -User $User -ErrorAction SilentlyContinue

        if ($null -eq $ExistingPermission) {
            # Add the permission if it doesn't exist
            Add-MailboxFolderPermission -Identity "${Email}:\Calendar" -User $User -AccessRights $Access
        } else {
            # Update the existing permission
            Set-MailboxFolderPermission -Identity "${Email}:\Calendar" -User $User -AccessRights $Access
        }
}

User-Login


do {
    # Prompt for calendar owner's email
    $CalendarOwner = Read-Host -Prompt "Please enter the email address of the calendar owner you want to view or modify"
    Get-Permissions $CalendarOwner

    $PermissionAction = Set-Action


    # If specific permission is needed what email is it
    if ($PermissionAction -eq 'specific') {
        $UserToModifyPermission = Read-Host -Prompt "Please enter the email address of the user permission you want to add or modify"
    }

    # If Permissions is modified what should we do
    if ($PermissionAction -eq 'specific' -or $PermissionAction -eq 'default'){
        do {
            # List permission levels
            Write-Host ""
            Write-Host "Please choose the permission level you'd like to set:"
            Write-Host "1. Owner"
            Write-Host "2. PublishingEditor"
            Write-Host "3. Editor"
            Write-Host "4. PublishingAuthor"
            Write-Host "5. Author"
            Write-Host "6. NonEditingAuthor"
            Write-Host "7. Reviewer"
            Write-Host "8. Contributor"
            Write-Host "9. AvailabilityOnly"
            Write-Host "10. LimitedDetails"

            # Prompt user to choose permission level
            $PermissionLevel = Read-Host -Prompt "Enter the number corresponding to the permission level"

            # Convert the number to a permission level string
            $PermissionLevelText = switch ($PermissionLevel) {
                "1" { "Owner" }
                "2" { "PublishingEditor" }
                "3" { "Editor" }
                "4" { "PublishingAuthor" }
                "5" { "Author" }
                "6" { "NonEditingAuthor" }
                "7" { "Reviewer" }
                "8" { "Contributor" }
                "9" { "AvailabilityOnly" }
                "10" { "LimitedDetails" }
                default {
                    Write-Host "Invalid selection. Please enter a number from 1 to 10." -ForegroundColor Red
                    $null
                }
            }
        } while (-not $PermissionLevelText)
    }


    # Perform the default action
    if ($PermissionAction -eq 'default') {
        Set-MailboxFolderPermission -Identity "${CalendarOwner}:\Calendar" -User Default -AccessRights $PermissionLevelText
    } 

    # Perform Specific Action    
    if ($PermissionAction -eq 'specific') {
        Add-Or-Update-Permission -Email $CalendarOwner -User $UserToModifyPermission  -Access $PermissionLevelText
    }

    # Perform Remove Action
    if ($PermissionAction -eq 'remove') {
        $UserToModifyPermission = Read-Host -Prompt "Please enter the email address of the user permission you want to remove"
        Remove-MailboxFolderPermission -Identity "${CalendarOwner}:\Calendar" -User $UserToModifyPermission
    }

    Get-Permissions $CalendarOwner

    # Ask if the user wants to make more changes
    $MakeMoreChanges = Read-Host -Prompt "Do you want to make more changes? [Y/N]"
} while ($MakeMoreChanges -in @('yes', 'Y', 'y'))

# Disconnect from Microsoft Exchange
Disconnect-ExchangeOnline
