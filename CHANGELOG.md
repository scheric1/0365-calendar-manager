# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/), and this project adheres to [Semantic Versioning](https://semver.org/).

## [2.1.1] - 2026-04-22

### Fixed
- `Test-ScriptVersion` fetched the wrong file (`Invoke-CalendarManager.ps1`) when checking for updates. Since `$script:ScriptVersion` lives in `_Shared.ps1`, the regex never matched and the update check always reported "running latest version". Now points at `_Shared.ps1`.

### Changed
- README replaces hardcoded version text with shields.io badges (release, PowerShell, platform). The release badge reads live from the GitHub Releases API, so future version bumps no longer require a README edit.

## [2.1.0] - 2026-04-22

### Added
- `-AutoInstall` switch parameter to automatically install the `ExchangeOnlineManagement` module without prompting

### Fixed
- `Set-CalendarPermission` failing to add users with no existing calendar permission. `Get-MailboxFolderPermission` throws rather than returning `$null` when a user has no entry, so the add path never ran — now uses `-ErrorAction SilentlyContinue` with an explicit null check

## [2.0.0] - 2026-04-13

### Added
- Project restructure: `setup/` directory, proper Verb-Noun naming
- Shared helpers (`_Shared.ps1`) with `Write-Status`, `Write-Header`, `Read-HostWithDefault`, `Test-ExoModuleInstalled`, `Invoke-SafeDisconnect`
- Prerequisite installer (`setup/Install-Prerequisites.ps1`) with multi-phase installation and verification
- Installation diagnostics (`setup/Test-Installation.ps1`) with 6-test suite and colour-coded results
- Colour-coded output with `[INFO]`/`[OK]`/`[WARNING]`/`[ERROR]` prefixes throughout
- Confirmation prompts before all permission changes with summary display
- CLI parameter support (`-CalendarOwner`, `-Action`, `-User`, `-PermissionLevel`) to skip interactive prompts
- Module availability check before connecting (with install instructions on failure)
- Success messages after each permission change

### Changed
- Main script renamed from `calendarManager.ps1` to `Invoke-CalendarManager.ps1`
- All functions renamed to Verb-Noun convention:
  - `Check-Version` → `Test-ScriptVersion`
  - `User-Login` → `Connect-ExchangeSession`
  - `Set-Action` → `Read-ActionChoice`
  - `Get-PermissionLevel` → `Read-PermissionLevel`
  - `Get-Permissions` → `Get-CalendarPermission`
  - `Add-Or-Update-Permission` → `Set-CalendarPermission`
- Version check now matches `$script:ScriptVersion` pattern (supports dot-sourced variable)
- Disconnect uses `Invoke-SafeDisconnect` with error suppression
- README rewritten with comprehensive documentation

### Breaking
- Main script renamed — update any shortcuts or aliases pointing to `calendarManager.ps1`

## [1.5.0] - 2025-08-12

### Added
- Version check against GitHub on startup
- Interactive calendar permission management
- Default, user-specific, and remove permission actions
- Permission level selection menu (10 levels)
- Multi-operation session support (loop until user exits)
- Exchange Online connection and disconnect handling
