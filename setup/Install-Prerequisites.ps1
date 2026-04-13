# Install-Prerequisites.ps1 — Install required modules for O365 Calendar Manager

param(
    [ValidateSet("CurrentUser", "AllUsers")]
    [string]$Scope = "CurrentUser"
)

# Dot-source shared helpers
. "$PSScriptRoot\..\_Shared.ps1"

Write-Header "Install Prerequisites"

# Phase 1: PowerShell Version Check
Write-Status -Level Info -Message "Phase 1: Checking PowerShell version..."

$psVersion = $PSVersionTable.PSVersion
$psEdition = $PSVersionTable.PSEdition

if ($psEdition -eq "Desktop" -and $psVersion.Major -ge 5) {
    Write-Status -Level OK -Message "PowerShell Desktop $psVersion — compatible."
} elseif ($psEdition -eq "Core" -and $psVersion.Major -ge 7) {
    Write-Status -Level OK -Message "PowerShell Core $psVersion — compatible."
} elseif ($psEdition -eq "Core") {
    Write-Status -Level Warning -Message "PowerShell Core $psVersion detected. Version 7+ is recommended."
} else {
    Write-Status -Level Error -Message "PowerShell $psVersion is not supported. Please upgrade to 5.1+ or 7+."
    exit 1
}

# Phase 2: Configure PowerShell Gallery
Write-Status -Level Info -Message "Phase 2: Configuring PowerShell Gallery..."

$gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
if ($gallery -and $gallery.InstallationPolicy -eq "Trusted") {
    Write-Status -Level OK -Message "PSGallery is already trusted."
} else {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Write-Status -Level OK -Message "PSGallery set to Trusted."
}

# Phase 3: Install ExchangeOnlineManagement
Write-Status -Level Info -Message "Phase 3: Installing ExchangeOnlineManagement module..."

if ($Scope -eq "AllUsers") {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Status -Level Error -Message "AllUsers scope requires administrator privileges. Run as admin or use -Scope CurrentUser."
        exit 1
    }
}

$existing = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1

if ($existing) {
    Write-Status -Level Info -Message "ExchangeOnlineManagement v$($existing.Version) is already installed."
    Write-Status -Level Info -Message "Checking for updates..."
    try {
        Update-Module -Name ExchangeOnlineManagement -Scope $Scope -ErrorAction Stop
        $updated = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
        if ($updated.Version -gt $existing.Version) {
            Write-Status -Level OK -Message "Updated to v$($updated.Version)."
        } else {
            Write-Status -Level OK -Message "Already on latest version."
        }
    } catch {
        Write-Status -Level Warning -Message "Could not check for updates: $_"
    }
} else {
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope $Scope -Force -AllowClobber -ErrorAction Stop
        Write-Status -Level OK -Message "ExchangeOnlineManagement installed successfully."
    } catch {
        Write-Status -Level Error -Message "Failed to install ExchangeOnlineManagement: $_"
        exit 1
    }
}

# Phase 4: Verification
Write-Status -Level Info -Message "Phase 4: Verifying installation..."

try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    $cmdCount = (Get-Command -Module ExchangeOnlineManagement).Count
    Write-Status -Level OK -Message "Module loaded successfully ($cmdCount commands available)."
} catch {
    Write-Status -Level Error -Message "Module installed but could not be imported: $_"
    exit 1
}

# Summary
Write-Header "Installation Complete"

$finalModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1

Write-Host ""
Write-Host "  Component                    Status" -ForegroundColor White
Write-Host "  ---------                    ------"
Write-Host "  PowerShell $psVersion" -NoNewline; Write-Host "              [OK]" -ForegroundColor Green
Write-Host "  PSGallery                    " -NoNewline; Write-Host "[OK]" -ForegroundColor Green
Write-Host "  ExchangeOnlineManagement v$($finalModule.Version) " -NoNewline; Write-Host "[OK]" -ForegroundColor Green
Write-Host ""

Write-Status -Level Info -Message "Next steps:"
Write-Host "  1. Run ./setup/Test-Installation.ps1 to verify everything works" -ForegroundColor Gray
Write-Host "  2. Run ./Invoke-CalendarManager.ps1 to start managing calendar permissions" -ForegroundColor Gray
Write-Host ""
