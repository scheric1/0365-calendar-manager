# _Shared.ps1 — Shared helpers for O365 Calendar Manager
# Dot-sourced by Invoke-CalendarManager.ps1 and setup scripts

$script:ScriptVersion = "2.0.0"
$script:RepoUrl = "https://github.com/scheric1/0365-calendar-manager"

function Write-Status {
    param(
        [Parameter(Mandatory)]
        [ValidateSet("Info", "OK", "Warning", "Error")]
        [string]$Level,

        [Parameter(Mandatory)]
        [string]$Message
    )

    $prefix = switch ($Level) {
        "Info"    { "[INFO]" }
        "OK"      { "[OK]" }
        "Warning" { "[WARNING]" }
        "Error"   { "[ERROR]" }
    }

    $colour = switch ($Level) {
        "Info"    { "White" }
        "OK"      { "Green" }
        "Warning" { "Yellow" }
        "Error"   { "Red" }
    }

    Write-Host "$prefix $Message" -ForegroundColor $colour
}

function Write-Header {
    param(
        [Parameter(Mandatory)]
        [string]$Title
    )

    $separator = "=" * ($Title.Length + 6)
    Write-Host ""
    Write-Host $separator -ForegroundColor Cyan
    Write-Host "   $Title" -ForegroundColor Cyan
    Write-Host $separator -ForegroundColor Cyan
}

function Read-HostWithDefault {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,

        [string]$Default
    )

    if ($Default) {
        $response = Read-Host -Prompt "$Prompt [$Default]"
        if ([string]::IsNullOrWhiteSpace($response)) { return $Default }
        return $response
    } else {
        return Read-Host -Prompt $Prompt
    }
}

function Test-ExoModuleInstalled {
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
    if ($module) {
        Write-Status -Level Info -Message "ExchangeOnlineManagement v$($module.Version) found."
        return $true
    } else {
        Write-Status -Level Error -Message "ExchangeOnlineManagement module is not installed."
        Write-Status -Level Info -Message "Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
        return $false
    }
}

function Test-ScriptVersion {
    $RawUrl = "https://raw.githubusercontent.com/scheric1/0365-calendar-manager/main/Invoke-CalendarManager.ps1"
    try {
        $response = Invoke-WebRequest -Uri $RawUrl -UseBasicParsing -ErrorAction Stop
        $remoteVersion = [regex]::Match($response.Content, '\$script:ScriptVersion\s*=\s*"([^"]+)"').Groups[1].Value
        if ($remoteVersion -and ([version]$remoteVersion -gt [version]$script:ScriptVersion)) {
            Write-Status -Level Warning -Message "A newer version ($remoteVersion) is available. Download it from $script:RepoUrl"
        } else {
            Write-Status -Level Info -Message "Running latest version ($script:ScriptVersion)."
        }
    } catch {
        Write-Status -Level Warning -Message "Unable to check for updates: $_"
    }
}

function Connect-ExchangeSession {
    if (-not (Test-ExoModuleInstalled)) {
        exit 1
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $UserEmail = Read-Host -Prompt "Please enter your email address to login"

    Write-Status -Level Info -Message "Connecting to Exchange Online..."
    try {
        Connect-ExchangeOnline -UserPrincipalName $UserEmail -ErrorAction Stop
        Write-Status -Level OK -Message "Connected to Exchange Online."
    }
    catch {
        Write-Status -Level Error -Message "Failed to connect to Exchange Online: $_"
        exit 1
    }
}

function Invoke-SafeDisconnect {
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Status -Level OK -Message "Disconnected from Exchange Online."
    }
    catch {
        Write-Status -Level Warning -Message "Could not disconnect cleanly: $_"
    }
}
