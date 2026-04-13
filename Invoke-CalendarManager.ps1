$ScriptVersion = "1.5.0"
$RepoUrl = "https://github.com/scheric1/0365-calendar-manager"

function Test-ScriptVersion {
    $RawUrl = "https://raw.githubusercontent.com/scheric1/0365-calendar-manager/main/Invoke-CalendarManager.ps1"
    try {
        $response = Invoke-WebRequest -Uri $RawUrl -UseBasicParsing -ErrorAction Stop
        $remoteVersion = [regex]::Match($response.Content, '\$ScriptVersion\s*=\s*"([^"]+)"').Groups[1].Value
        if ($remoteVersion -and ([version]$remoteVersion -gt [version]$ScriptVersion)) {
            Write-Host "A newer version ($remoteVersion) is available. Download it from $RepoUrl" -ForegroundColor Yellow
        } else {
            Write-Host "Running latest version ($ScriptVersion)."
        }
    } catch {
        Write-Host "Unable to check for updates: $_" -ForegroundColor Yellow
    }
}

function Connect-ExchangeSession {
    $UserEmail = Read-Host -Prompt "Please enter your email address to login"

    try {
        Connect-ExchangeOnline -UserPrincipalName $UserEmail -ErrorAction Stop
    }
    catch {
        Write-Error -Message "Failed to connect to Exchange Online: $_" -ErrorId 1001
        exit
    }
}

function Read-ActionChoice {
    do {
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
        Write-Host "Please choose the permission level you'd like to set:"
        foreach ($k in $permissionLevels.Keys) {
            Write-Host "$k. $($permissionLevels[$k])"
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
        Write-Error -Message "Failed to retrieve permissions for ${CalendarOwner}: $_" -ErrorId 1002
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
        Write-Error -Message "Failed to retrieve existing permission for ${User}: $_" -ErrorId 1003
        return
    }

    if ($null -eq $existingPermission) {
        try {
            Add-MailboxFolderPermission -Identity "${Email}:\Calendar" -User $User -AccessRights $Access -ErrorAction Stop
        }
        catch {
            Write-Error -Message "Failed to add permission for ${User}: $_" -ErrorId 1004
        }
    } else {
        try {
            Set-MailboxFolderPermission -Identity "${Email}:\Calendar" -User $User -AccessRights $Access -ErrorAction Stop
        }
        catch {
            Write-Error -Message "Failed to update permission for ${User}: $_" -ErrorId 1005
        }
    }
}

Test-ScriptVersion
Connect-ExchangeSession


do {
    $CalendarOwner = Read-Host -Prompt "Please enter the email address of the calendar owner you want to view or modify"
    Get-CalendarPermission $CalendarOwner

    $PermissionAction = Read-ActionChoice

    if ($PermissionAction -eq 'specific') {
        $UserToModifyPermission = Read-Host -Prompt "Please enter the email address of the user permission you want to add or modify"
    }

    if ($PermissionAction -eq 'specific' -or $PermissionAction -eq 'default') {
        $PermissionLevelText = Read-PermissionLevel
    }

    # Perform the default action
    if ($PermissionAction -eq 'default') {
        try {
            Set-MailboxFolderPermission -Identity "${CalendarOwner}:\Calendar" -User Default -AccessRights $PermissionLevelText -ErrorAction Stop
        }
        catch {
            Write-Error -Message "Failed to set default permissions for ${CalendarOwner}: $_" -ErrorId 1006
        }
    }

    # Perform specific action
    if ($PermissionAction -eq 'specific') {
        Set-CalendarPermission -Email $CalendarOwner -User $UserToModifyPermission -Access $PermissionLevelText
    }

    # Perform remove action
    if ($PermissionAction -eq 'remove') {
        $UserToModifyPermission = Read-Host -Prompt "Please enter the email address of the user permission you want to remove"
        try {
            Remove-MailboxFolderPermission -Identity "${CalendarOwner}:\Calendar" -User $UserToModifyPermission -ErrorAction Stop
        }
        catch {
            Write-Error -Message "Failed to remove permission for ${UserToModifyPermission}: $_" -ErrorId 1007
        }
    }

    Get-CalendarPermission $CalendarOwner

    $MakeMoreChanges = Read-Host -Prompt "Do you want to make more changes? [Y/N]"
} while ($MakeMoreChanges -in @('yes', 'Y', 'y'))

# Disconnect from Microsoft Exchange
try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
}
catch {
    Write-Error -Message "Failed to disconnect from Exchange Online: $_" -ErrorId 1008
}
