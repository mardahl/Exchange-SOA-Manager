<#
.SYNOPSIS
    Exchange SOA Manager - a terminal UI for managing Source of Authority (SOA)
    in Exchange Online hybrid environments.

.DESCRIPTION
    A dependency-light PowerShell TUI to switch the Source of Authority for:

      * User mailboxes  - toggles IsExchangeCloudManaged via Set-Mailbox
                          (Exchange attribute SOA for dir-synced mailboxes)
      * Groups          - toggles isCloudManaged via the Microsoft Graph
                          onPremisesSyncBehavior API (object-level SOA)
      * Contacts        - toggles isCloudManaged via the Microsoft Graph
                          onPremisesSyncBehavior API (object-level SOA)
      * Organization    - tenant-wide default for new dir-synced mailboxes
                          (Set-OrganizationConfig -ExchangeAttributesCloudManagedByDefault)

    Features:
      * Pure PowerShell terminal UI (VT/ANSI) - no WinForms, no DLLs.
      * Works on Windows PowerShell 5.1 and PowerShell 7+ (incl. macOS/Linux).
      * Attribute backup (JSON) before every conversion.
      * Group rollback safety check - detects cloud-only members that would
        break sync before converting a group back to on-premises.
      * CSV export of the current view and CSV/TXT bulk import selection.
      * Search, status filters, sorting, multi-select, batch operations.
      * Timestamped log file and in-app log viewer.
      * Demo mode with generated data for evaluation and screenshots.

    Required modules (installed on demand, CurrentUser scope):
      * ExchangeOnlineManagement      (mailboxes + organization config)
      * Microsoft.Graph.Authentication (groups + contacts, lightweight REST)

    Graph is run out-of-process to avoid MSAL assembly conflicts with
    ExchangeOnlineManagement. See Connect-GraphService for details.

    Required roles / permissions:
      * Mailboxes / Org : Exchange Administrator (or Global Administrator)
      * Groups / Contacts: Hybrid Identity Administrator + admin consent for
        Group.Read.All, GroupMember.Read.All, User.Read.All,
        Group-OnPremisesSyncBehavior.ReadWrite.All, OrgContact.Read.All,
        Contacts-OnPremisesSyncBehavior.ReadWrite.All

.PARAMETER Demo
    Run with generated sample data. No connections are made and nothing is
    changed. Useful to explore the UI and take screenshots.

.PARAMETER Ascii
    Use plain ASCII glyphs instead of Unicode box drawing characters.
    Helpful for legacy consoles with raster fonts.

.PARAMETER NoDisconnect
    Keep the Exchange Online and Graph sessions alive when the tool exits.

.EXAMPLE
    .\SOA-Manager.ps1

.EXAMPLE
    .\SOA-Manager.ps1 -Demo

.NOTES
    Version : 1.1.0
    License : MIT
    Inspired by codeandersen's Exchange-SOA-Conversion-tool (WinForms, users
    only) - this tool adds groups, contacts, the tenant-wide default switch
    and a portable terminal UI.

.LINK
    https://learn.microsoft.com/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

.LINK
    https://learn.microsoft.com/entra/identity/hybrid/how-to-group-source-of-authority-configure
#>

#Requires -Version 5.1

[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Terminal UI: Console.Write is the renderer; Write-Host is used for main-buffer auth prompts.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '', Justification = 'Best-effort cleanup paths (console restore, disconnect, log fallback).')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helpers; the TUI has its own confirmation dialogs.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Internal helpers, not exported cmdlets.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Ascii', Justification = 'Consumed in dot-sourced src/00-globals.ps1 at script scope; not visible to per-file static analysis.')]
param(
    [switch]$Demo,
    [switch]$Ascii,
    [switch]$NoDisconnect
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:Root = $PSScriptRoot

# ============================================================================
# Load modules. Dot-sourced with the `foreach` statement (NOT ForEach-Object)
# so they run at script scope: functions and $script: state land in the shared
# scope, exactly as when this was one file. Load order is fixed by file name.
# ============================================================================
foreach ($f in (Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter '*.ps1' | Sort-Object Name)) {
    . $f.FullName
}

# ============================================================================
#region Main
# ============================================================================
$script:DemoMode = [bool]$Demo
$script:KeepSessions = [bool]$NoDisconnect

Write-SoaLog -Message ('=' * 60)
Write-SoaLog -Message ("Exchange SOA Manager v{0} started (PS {1}, Demo={2})" -f $script:Version, $PSVersionTable.PSVersion, $script:DemoMode)
Write-SoaLog -Message ('=' * 60)

if ($script:DemoMode) {
    Write-SoaLog -Message 'DEMO MODE - no connections are made, no changes are written.' -Level WARN
}

try {
    Enter-Tui
    $lastW = 0; $lastH = 0
    while (-not $script:UI.Quit) {
        $size = Get-ConsoleSize
        if ($size[0] -ne $lastW -or $size[1] -ne $lastH) {
            $lastW = $size[0]; $lastH = $size[1]
            [Console]::Write("$script:ESC[2J")
            $script:UI.Dirty = $true
        }
        if ($script:UI.Dirty) {
            Write-Screen
            $script:UI.Dirty = $false
        }
        if ([Console]::KeyAvailable) {
            while ([Console]::KeyAvailable -and -not $script:UI.Quit) {
                $k = [Console]::ReadKey($true)
                Invoke-KeyDispatch -K $k
            }
            $script:UI.Dirty = $true
        } else {
            Start-Sleep -Milliseconds 25
        }
    }
} finally {
    Exit-Tui
    if (-not $script:KeepSessions) {
        Write-Host 'Closing sessions...' -ForegroundColor DarkGray
        Disconnect-AllServices
    } else {
        Write-SoaLog -Message 'Sessions left open (-NoDisconnect).'
    }
    Write-SoaLog -Message 'Exchange SOA Manager ended.'
    Write-Host ''
    Write-Host 'Exchange SOA Manager closed.' -ForegroundColor Cyan
    Write-Host ("  Log     : " + $script:LogFile) -ForegroundColor DarkGray
    if (Test-Path -LiteralPath $script:BackupDir)  { Write-Host ("  Backups : " + $script:BackupDir) -ForegroundColor DarkGray }
    if (Test-Path -LiteralPath $script:ExportDir)  { Write-Host ("  Exports : " + $script:ExportDir) -ForegroundColor DarkGray }
}

#endregion
