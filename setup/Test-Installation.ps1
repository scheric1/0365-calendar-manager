# Test-Installation.ps1 — Diagnostic checks for O365 Calendar Manager

# Dot-source shared helpers
. "$PSScriptRoot\..\_Shared.ps1"

Write-Header "Installation Diagnostics"

# System Information
Write-Host ""
Write-Host "  System Information" -ForegroundColor White
Write-Host "  ------------------"
Write-Host "  OS:               $([System.Runtime.InteropServices.RuntimeInformation]::OSDescription)" -ForegroundColor Gray
Write-Host "  PS Edition:       $($PSVersionTable.PSEdition)" -ForegroundColor Gray
Write-Host "  PS Version:       $($PSVersionTable.PSVersion)" -ForegroundColor Gray
Write-Host "  Execution Policy: $(Get-ExecutionPolicy -Scope CurrentUser)" -ForegroundColor Gray

$executionPolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($executionPolicy -eq "Restricted") {
    Write-Status -Level Warning -Message "Execution policy is 'Restricted'. Run: Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned"
}
Write-Host ""

# Test Suite
$tests = @()

# Test 1: PowerShell Version
$psVersion = $PSVersionTable.PSVersion
$psOk = ($PSVersionTable.PSEdition -eq "Desktop" -and $psVersion.Major -ge 5) -or
        ($PSVersionTable.PSEdition -eq "Core" -and $psVersion.Major -ge 7)
$tests += @{ Name = "PowerShell version ($psVersion)"; Passed = $psOk; Detail = if (-not $psOk) { "Requires 5.1+ (Desktop) or 7+ (Core)" } else { $null } }

# Test 2: ExchangeOnlineManagement Module
$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
$moduleOk = $null -ne $module
$tests += @{ Name = "ExchangeOnlineManagement module"; Passed = $moduleOk; Detail = if ($moduleOk) { "v$($module.Version)" } else { "Not installed. Run ./setup/Install-Prerequisites.ps1" } }

# Test 3: Module Import
$importOk = $false
if ($moduleOk) {
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        $importOk = $true
    } catch {
        # Import failed
    }
}
$tests += @{ Name = "Module import"; Passed = $importOk; Detail = if (-not $importOk -and $moduleOk) { "Module found but could not be imported" } elseif (-not $moduleOk) { "Skipped (module not installed)" } else { $null } }

# Test 4: Script Files
$scriptRoot = Split-Path $PSScriptRoot -Parent
$requiredFiles = @("Invoke-CalendarManager.ps1", "_Shared.ps1")
$missingFiles = $requiredFiles | Where-Object { -not (Test-Path (Join-Path $scriptRoot $_)) }
$filesOk = $missingFiles.Count -eq 0
$tests += @{ Name = "Required script files"; Passed = $filesOk; Detail = if (-not $filesOk) { "Missing: $($missingFiles -join ', ')" } else { $null } }

# Test 5: Write Permissions
$writeOk = $false
try {
    $testFile = Join-Path $scriptRoot ".write-test-$(Get-Random).tmp"
    [System.IO.File]::WriteAllText($testFile, "test")
    Remove-Item $testFile -Force
    $writeOk = $true
} catch {
    # Write failed
}
$tests += @{ Name = "Write permissions"; Passed = $writeOk; Detail = if (-not $writeOk) { "Cannot write to script directory" } else { $null } }

# Test 6: Network Connectivity
$networkOk = $false
try {
    $null = Invoke-WebRequest -Uri "https://www.powershellgallery.com" -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
    $networkOk = $true
} catch {
    # Network unreachable
}
$tests += @{ Name = "Network (PowerShell Gallery)"; Passed = $networkOk; Detail = if (-not $networkOk) { "Cannot reach powershellgallery.com" } else { $null } }

# Results
Write-Host "  Test Results" -ForegroundColor White
Write-Host "  ------------"

$passed = 0
$failed = 0

foreach ($test in $tests) {
    if ($test.Passed) {
        Write-Host "  [PASS] $($test.Name)" -ForegroundColor Green
        $passed++
    } else {
        Write-Host "  [FAIL] $($test.Name)" -ForegroundColor Red
        $failed++
    }
    if ($test.Detail) {
        Write-Host "         $($test.Detail)" -ForegroundColor Gray
    }
}

# Summary
Write-Host ""
Write-Host "  Total: $($tests.Count)  Passed: $passed  Failed: $failed" -ForegroundColor $(if ($failed -gt 0) { "Yellow" } else { "Green" })
Write-Host ""

if ($failed -gt 0) {
    Write-Status -Level Warning -Message "Some tests failed. Run ./setup/Install-Prerequisites.ps1 to fix."
    exit 1
} else {
    Write-Status -Level OK -Message "All tests passed. You're ready to go!"
    exit 0
}
