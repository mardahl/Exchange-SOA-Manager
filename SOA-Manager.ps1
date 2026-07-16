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
param(
    [switch]$Demo,
    [switch]$Ascii,
    [switch]$NoDisconnect
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# ============================================================================
#region Globals & State
# ============================================================================

$script:Version = '1.3.8'
$script:ESC     = [char]27
$script:IsWin   = ($PSVersionTable.PSVersion.Major -lt 6) -or ($null -ne (Get-Variable -Name IsWindows -ErrorAction SilentlyContinue) -and $IsWindows)

if ($PSScriptRoot) { $script:Root = $PSScriptRoot } else { $script:Root = (Get-Location).Path }

$script:LogFile   = Join-Path $script:Root ("SOA-Manager_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$script:BackupDir = Join-Path $script:Root 'SOA-Backups'
$script:ExportDir = Join-Path $script:Root 'SOA-Exports'
$script:LogBuffer = New-Object System.Collections.ArrayList
$script:Spinner   = $null   # background-runspace spinner (see Start-LoadSpinner)

# Connection state
$script:Conn = @{
    Exo          = $false
    Graph        = $false
    ExoAccount   = ''
    GraphAccount = ''
}

# Out-of-process Graph worker state
$script:GraphWorker = @{
    Process    = $null
    StdIn      = $null   # StreamWriter
    StdOut     = $null   # StreamReader
    StdErr     = $null   # StreamReader
    LastError  = ''
    WorkerPath = ''
    Account    = ''
}

# Organization config state
$script:Org = @{
    Loaded      = $false
    CloudDefault= $false
    CheckedAt   = $null
    TenantName  = ''
}

# UI state
$script:UI = @{
    Quit       = $false
    Dirty      = $true
    W          = 0
    H          = 0
    Tab        = 0          # index into $script:Tabs
    SearchMode = $false
    LogScroll  = 0          # 0 = pinned to bottom
}

# Demo org state
$script:DemoOrgCloudDefault = $false

function New-ListTab {
    param([string]$Name, [string]$Conn, [string]$Noun)
    return @{
        Kind       = 'List'
        Name       = $Name
        Conn       = $Conn        # 'Exo' | 'Graph'
        Noun       = $Noun        # 'mailboxes' | 'groups' | 'contacts'
        Items      = @()          # full data set
        View       = @()          # filtered + sorted projection
        Loaded     = $false
        Cursor     = 0
        Scroll     = 0
        Search     = ''
        Filter     = 'All'        # All | Cloud | OnPrem | Pending
        SortCol    = 'Name'       # Name | Email | Soa
        SortDesc   = $false
    }
}

$script:Tabs = @(
    (New-ListTab -Name 'Mailboxes' -Conn 'Exo'   -Noun 'mailboxes'),
    (New-ListTab -Name 'Groups'    -Conn 'Graph' -Noun 'groups'),
    (New-ListTab -Name 'Contacts'  -Conn 'Graph' -Noun 'contacts'),
    @{ Kind = 'Org'; Name = 'Organization'; Conn = 'Exo' },
    @{ Kind = 'Log'; Name = 'Log';          Conn = '' }
)

# Glyphs (Unicode with ASCII fallback)
if ($Ascii) {
    $script:G = @{
        H='-'; V='|'; TL='+'; TR='+'; BL='+'; BR='+'
        Dot='*'; Half='~'; Ring='o'
        BarOn='#'; BarOff='-'
        Up='^'; Down='v'; Ell='..'
        ChkOn='[x]'; ChkOff='[ ]'
        Arrow='->'
    }
} else {
    $script:G = @{
        H=([char]0x2500); V=([char]0x2502); TL=([char]0x250C); TR=([char]0x2510); BL=([char]0x2514); BR=([char]0x2518)
        Dot=([char]0x25CF); Half=([char]0x25D0); Ring=([char]0x25CB)
        BarOn=([char]0x2588); BarOff=([char]0x2591)
        Up=([char]0x2191); Down=([char]0x2193); Ell=([char]0x2026)
        ChkOn=('[' + [char]0x25A0 + ']'); ChkOff='[ ]'
        Arrow=([char]0x2192)
    }
}

# Theme (256-color SGR sequences)
$e = $script:ESC
$script:T = @{
    Reset      = "$e[0m"
    TitleBg    = "$e[48;5;236m"
    TitleApp   = "$e[1;38;5;45;48;5;236m"
    TitleDim   = "$e[38;5;245;48;5;236m"
    TitleOk    = "$e[38;5;42;48;5;236m"
    TitleOff   = "$e[38;5;240;48;5;236m"
    TitleDemo  = "$e[1;38;5;213;48;5;236m"
    TabOn      = "$e[1;38;5;16;48;5;45m"
    TabOff     = "$e[38;5;248;48;5;236m"
    TabBg      = "$e[48;5;236m"
    Ctx        = "$e[38;5;245m"
    CtxHi      = "$e[38;5;252m"
    ColHead    = "$e[1;38;5;250m"
    Row        = "$e[38;5;252m"
    RowDim     = "$e[38;5;245m"
    CursorBg   = "$e[48;5;24m"
    CursorFg   = "$e[38;5;231;48;5;24m"
    SelMark    = "$e[1;38;5;45m"
    Cloud      = "$e[38;5;45m"
    OnPrem     = "$e[38;5;214m"
    Pending    = "$e[38;5;171m"
    Good       = "$e[38;5;42m"
    Warn       = "$e[38;5;220m"
    Danger     = "$e[1;38;5;196m"
    Muted      = "$e[38;5;245m"
    FootBg     = "$e[48;5;236m"
    FootKey    = "$e[1;38;5;45;48;5;236m"
    FootTxt    = "$e[38;5;245;48;5;236m"
    Border     = "$e[38;5;45m"
    BorderWarn = "$e[38;5;220m"
    BorderErr  = "$e[38;5;196m"
    ModalTitle = "$e[1;38;5;231m"
    Input      = "$e[38;5;231;48;5;238m"
    BarOn      = "$e[38;5;45m"
    BarOff     = "$e[38;5;238m"
}

#endregion

# ============================================================================
#region Logging
# ============================================================================

function Write-SoaLog {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK')][string]$Level = 'INFO'
    )
    $stamp = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
    $entry = "[$stamp] [$Level] $Message"
    [void]$script:LogBuffer.Add(@{ Stamp=$stamp; Level=$Level; Message=$Message })
    try { Add-Content -Path $script:LogFile -Value $entry -Encoding UTF8 } catch { }
}

#endregion

# ============================================================================
#region Console / VT engine
# ============================================================================

$script:SavedOutputEncoding = $null
$script:SavedCtrlC          = $null
$script:TuiActive           = $false

function Enable-VirtualTerminal {
    if (-not $script:IsWin) { return $true }
    try {
        if (-not ('SoaTui.Native' -as [type])) {
            Add-Type -Namespace SoaTui -Name Native -MemberDefinition @'
[DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr GetStdHandle(int nStdHandle);
[DllImport("kernel32.dll", SetLastError=true)] public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
[DllImport("kernel32.dll", SetLastError=true)] public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
'@
        }
        $handle = [SoaTui.Native]::GetStdHandle(-11)   # STD_OUTPUT_HANDLE
        $mode = 0
        if ([SoaTui.Native]::GetConsoleMode($handle, [ref]$mode)) {
            # ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x4
            [void][SoaTui.Native]::SetConsoleMode($handle, $mode -bor 4)
        }
        return $true
    } catch {
        return $false
    }
}

function Enter-Tui {
    if ($script:TuiActive) { return }
    $script:SavedOutputEncoding = [Console]::OutputEncoding
    try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }
    [void](Enable-VirtualTerminal)
    try {
        $script:SavedCtrlC = [Console]::TreatControlCAsInput
        [Console]::TreatControlCAsInput = $true
    } catch { }
    [Console]::Write("$script:ESC[?1049h")   # alternate screen buffer
    [Console]::Write("$script:ESC[?25l")     # hide cursor
    [Console]::Write("$script:ESC[2J")
    $script:TuiActive = $true
    $script:UI.Dirty = $true
}

function Exit-Tui {
    if (-not $script:TuiActive) { return }
    [Console]::Write("$script:ESC[0m")
    [Console]::Write("$script:ESC[?25h")     # show cursor
    [Console]::Write("$script:ESC[?1049l")   # back to main buffer
    try {
        if ($null -ne $script:SavedCtrlC) { [Console]::TreatControlCAsInput = $script:SavedCtrlC }
    } catch { }
    try {
        if ($null -ne $script:SavedOutputEncoding) { [Console]::OutputEncoding = $script:SavedOutputEncoding }
    } catch { }
    $script:TuiActive = $false
}

function Invoke-OnMainBuffer {
    # Temporarily leave the TUI (for interactive auth, module install, ...)
    param([Parameter(Mandatory=$true)][scriptblock]$Action)
    $wasActive = $script:TuiActive
    if ($wasActive) { Exit-Tui }
    try {
        & $Action
    } finally {
        if ($wasActive) { Enter-Tui }
    }
}

function Get-ConsoleSize {
    $w = 80; $h = 24
    try { $w = [Console]::WindowWidth; $h = [Console]::WindowHeight } catch { }
    if ($w -lt 1) { $w = 80 }
    if ($h -lt 1) { $h = 24 }
    return @($w, $h)
}

#endregion

# ============================================================================
#region Drawing primitives
# ============================================================================

function Get-PadCell {
    # Truncate (with ellipsis) or right-pad plain text to an exact width.
    param([string]$Text, [int]$Width, [switch]$AlignRight)
    if ($Width -le 0) { return '' }
    if ($null -eq $Text) { $Text = '' }
    if ($Text.Length -gt $Width) {
        $ell = [string]$script:G.Ell
        if ($Width -le $ell.Length) { return $Text.Substring(0, $Width) }
        return $Text.Substring(0, $Width - $ell.Length) + $ell
    }
    if ($AlignRight) { return $Text.PadLeft($Width) }
    return $Text.PadRight($Width)
}

function Add-FrameLine {
    # Append one full screen line (absolute row) to the frame builder.
    param([System.Text.StringBuilder]$Sb, [int]$Row, [string]$Content)
    [void]$Sb.Append("$script:ESC[$Row;1H")
    [void]$Sb.Append($Content)
    [void]$Sb.Append("$($script:T.Reset)$script:ESC[K")
}

function Get-SoaBadge {
    param([string]$Soa, [int]$Width = 10)
    $t = $script:T; $g = $script:G
    switch ($Soa) {
        'Cloud'   { return ($t.Cloud   + (Get-PadCell ("$($g.Dot) Cloud")   $Width)) }
        'OnPrem'  { return ($t.OnPrem  + (Get-PadCell ("$($g.Dot) On-prem") $Width)) }
        'Pending' { return ($t.Pending + (Get-PadCell ("$($g.Half) Pending") $Width)) }
        default   { return ($t.Muted   + (Get-PadCell '?' $Width)) }
    }
}

function Get-FooterBar {
    # Render key hints: array of @(key,label) pairs, truncated to width.
    param([object[]]$Hints, [int]$Width)
    $t = $script:T
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append($t.FootBg)
    $plainLen = 0
    foreach ($pair in $Hints) {
        $k = [string]$pair[0]; $l = [string]$pair[1]
        $piece = ' ' + $k + ' ' + $l + ' '
        if (($plainLen + $piece.Length) -gt $Width) { break }
        [void]$sb.Append($t.FootKey).Append(' ').Append($k).Append($t.FootTxt).Append(' ').Append($l).Append(' ')
        $plainLen += $piece.Length
    }
    if ($plainLen -lt $Width) {
        [void]$sb.Append($t.FootBg).Append((' ' * ($Width - $plainLen)))
    }
    return $sb.ToString()
}

function Write-CenteredPanel {
    # Render centered lines in the content area of the frame.
    # Lines: array of @(styledText, plainLength) or plain strings.
    param([System.Text.StringBuilder]$Sb, [object[]]$Lines, [int]$Top, [int]$Bottom, [int]$Width)
    $count = $Lines.Count
    $area = $Bottom - $Top + 1
    $startRow = $Top + [Math]::Max(0, [int](($area - $count) / 2))
    $row = $startRow
    foreach ($ln in $Lines) {
        if ($row -gt $Bottom) { break }
        $styled = ''; $plain = 0
        if ($ln -is [array]) { $styled = [string]$ln[0]; $plain = [int]$ln[1] }
        else { $styled = [string]$ln; $plain = ([string]$ln).Length }
        $padLeft = [Math]::Max(0, [int](($Width - $plain) / 2))
        Add-FrameLine -Sb $Sb -Row $row -Content ((' ' * $padLeft) + $styled)
        $row++
    }
    return $row
}

#endregion

# ============================================================================
#region Modals
# ============================================================================

function Split-TextLines {
    # Word-wrap plain text to a width, preserving leading indentation.
    param([string]$Text, [int]$Width)
    $out = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrEmpty($Text)) { return ,@('') }
    foreach ($para in ($Text -split "`n")) {
        if ($para.Length -eq 0) { [void]$out.Add(''); continue }
        if ($para.Length -le $Width) { [void]$out.Add($para); continue }
        $indent = ''
        $m = [regex]::Match($para, '^\s+')
        if ($m.Success) { $indent = $m.Value }
        $bodyW = [Math]::Max(8, $Width - $indent.Length)
        $rest = $para.TrimStart()
        if ($rest.Length -eq 0) { [void]$out.Add($para); continue }
        $line = ''
        foreach ($word in ($rest -split ' ')) {
            while ($word.Length -gt $bodyW) {
                # hard-break tokens longer than the line (e.g. file paths)
                if ($line.Length -gt 0) { [void]$out.Add($indent + $line); $line = '' }
                [void]$out.Add($indent + $word.Substring(0, $bodyW))
                $word = $word.Substring($bodyW)
            }
            if ($word.Length -eq 0) { continue }
            if ($line.Length -eq 0) { $line = $word }
            elseif (($line.Length + 1 + $word.Length) -le $bodyW) { $line = $line + ' ' + $word }
            else { [void]$out.Add($indent + $line); $line = $word }
        }
        if ($line.Length -gt 0) { [void]$out.Add($indent + $line) }
    }
    return ,$out.ToArray()
}

function ConvertTo-ModalLines {
    # Normalize: items may be [string] or @($styleSgr, $text). Wrap strings.
    param([object[]]$Lines, [int]$Width)
    $out = New-Object System.Collections.ArrayList
    foreach ($ln in $Lines) {
        if ($ln -is [array]) {
            $style = [string]$ln[0]; $text = [string]$ln[1]
            foreach ($w in (Split-TextLines -Text $text -Width $Width)) {
                [void]$out.Add(@($style, $w))
            }
        } else {
            foreach ($w in (Split-TextLines -Text ([string]$ln) -Width $Width)) {
                [void]$out.Add(@($script:T.Row, $w))
            }
        }
    }
    return ,$out.ToArray()
}

function Write-ModalFrame {
    # Draw a bordered box with title; returns geometry hashtable.
    param(
        [string]$Title,
        [object[]]$BodyLines,      # normalized @($style,$text) pairs
        [string]$FooterHint,
        [string]$BorderStyle,
        [int]$MinWidth = 68,
        [int]$FixedBodyHeight = 0, # 0 = size to content
        [int]$BodyScroll = 0
    )
    $t = $script:T; $g = $script:G
    $size = Get-ConsoleSize; $W = $size[0]; $H = $size[1]
    if (-not $BorderStyle) { $BorderStyle = $t.Border }

    $boxW = [Math]::Min([Math]::Max($MinWidth, $Title.Length + 8), $W - 4)
    $innerW = $boxW - 4

    $bodyH = $BodyLines.Count
    if ($FixedBodyHeight -gt 0) { $bodyH = $FixedBodyHeight }
    $maxBodyH = $H - 8
    if ($maxBodyH -lt 3) { $maxBodyH = 3 }
    $scrollable = $false
    if ($bodyH -gt $maxBodyH) { $bodyH = $maxBodyH; $scrollable = $true }

    $boxH = $bodyH + 4   # top border, blank-ish padding handled in body, footer hint, bottom border
    $x = [Math]::Max(1, [int](($W - $boxW) / 2) + 1)
    $y = [Math]::Max(1, [int](($H - $boxH) / 2) + 1)

    $sb = New-Object System.Text.StringBuilder
    $hChar = [string]$g.H

    # top border with title
    $tt = " $Title "
    if ($tt.Length -gt ($boxW - 4)) { $tt = Get-PadCell $tt ($boxW - 4) }
    $dashTotal = $boxW - 2 - $tt.Length
    $dashL = 1
    $dashR = [Math]::Max(0, $dashTotal - $dashL)
    [void]$sb.Append("$script:ESC[$y;$($x)H")
    [void]$sb.Append($BorderStyle).Append([string]$g.TL).Append($hChar * $dashL)
    [void]$sb.Append($t.ModalTitle).Append($tt).Append($t.Reset).Append($BorderStyle)
    [void]$sb.Append($hChar * $dashR).Append([string]$g.TR).Append($t.Reset)

    # body rows
    $visible = $BodyLines
    if ($scrollable -or ($FixedBodyHeight -gt 0 -and $BodyLines.Count -gt $bodyH)) {
        $start = [Math]::Max(0, [Math]::Min($BodyScroll, $BodyLines.Count - $bodyH))
        $visible = $BodyLines[$start..([Math]::Min($BodyLines.Count - 1, $start + $bodyH - 1))]
    }
    $row = $y + 1
    for ($i = 0; $i -lt $bodyH; $i++) {
        $style = $t.Row; $text = ''
        if ($i -lt $visible.Count) {
            $pair = $visible[$i]
            $style = [string]$pair[0]; $text = [string]$pair[1]
        }
        [void]$sb.Append("$script:ESC[$row;$($x)H")
        [void]$sb.Append($BorderStyle).Append([string]$g.V).Append($t.Reset).Append(' ')
        [void]$sb.Append($style).Append((Get-PadCell $text $innerW)).Append($t.Reset)
        [void]$sb.Append(' ').Append($BorderStyle).Append([string]$g.V).Append($t.Reset)
        $row++
    }

    # footer hint row
    [void]$sb.Append("$script:ESC[$row;$($x)H")
    [void]$sb.Append($BorderStyle).Append([string]$g.V).Append($t.Reset).Append(' ')
    [void]$sb.Append($t.Muted).Append((Get-PadCell $FooterHint $innerW -AlignRight)).Append($t.Reset)
    [void]$sb.Append(' ').Append($BorderStyle).Append([string]$g.V).Append($t.Reset)
    $row++

    # bottom border
    [void]$sb.Append("$script:ESC[$row;$($x)H")
    [void]$sb.Append($BorderStyle).Append([string]$g.BL).Append($hChar * ($boxW - 2)).Append([string]$g.BR).Append($t.Reset)

    [Console]::Write($sb.ToString())
    return @{ X=$x; Y=$y; W=$boxW; H=$boxH; InnerW=$innerW; BodyH=$bodyH; Total=$BodyLines.Count; Scrollable=$scrollable }
}

function Read-ModalKey {
    while ($true) {
        if ([Console]::KeyAvailable) { return [Console]::ReadKey($true) }
        Start-Sleep -Milliseconds 20
    }
}

function Show-MsgModal {
    param([string]$Title, [object[]]$Lines, [ValidateSet('Info','Warn','Error')][string]$Kind = 'Info')
    $border = $script:T.Border
    if ($Kind -eq 'Warn')  { $border = $script:T.BorderWarn }
    if ($Kind -eq 'Error') { $border = $script:T.BorderErr }
    $norm = ConvertTo-ModalLines -Lines $Lines -Width 64
    $scroll = 0
    while ($true) {
        Write-Screen
        $geo = Write-ModalFrame -Title $Title -BodyLines $norm -FooterHint 'Enter close' -BorderStyle $border -BodyScroll $scroll
        $k = Read-ModalKey
        switch ($k.Key) {
            'Enter'      { $script:UI.Dirty = $true; return }
            'Escape'     { $script:UI.Dirty = $true; return }
            'UpArrow'    { if ($scroll -gt 0) { $scroll-- } }
            'DownArrow'  { if ($geo.Scrollable -and $scroll -lt ($geo.Total - $geo.BodyH)) { $scroll++ } }
            'PageUp'     { $scroll = [Math]::Max(0, $scroll - $geo.BodyH) }
            'PageDown'   { if ($geo.Scrollable) { $scroll = [Math]::Min($geo.Total - $geo.BodyH, $scroll + $geo.BodyH) } }
        }
    }
}

function Show-ConfirmModal {
    param([string]$Title, [object[]]$Lines, [switch]$Danger)
    $border = $script:T.Border
    if ($Danger) { $border = $script:T.BorderErr }
    $norm = ConvertTo-ModalLines -Lines $Lines -Width 64
    $scroll = 0
    while ($true) {
        Write-Screen
        $geo = Write-ModalFrame -Title $Title -BodyLines $norm -FooterHint 'Y yes   N/Esc no' -BorderStyle $border -BodyScroll $scroll
        $k = Read-ModalKey
        if ($k.Key -eq 'UpArrow')   { if ($scroll -gt 0) { $scroll-- }; continue }
        if ($k.Key -eq 'DownArrow') { if ($geo.Scrollable -and $scroll -lt ($geo.Total - $geo.BodyH)) { $scroll++ }; continue }
        if ($k.Key -eq 'Escape' -or [char]::ToUpper($k.KeyChar) -eq 'N') { $script:UI.Dirty = $true; return $false }
        if ([char]::ToUpper($k.KeyChar) -eq 'Y') { $script:UI.Dirty = $true; return $true }
        if (($k.Modifiers -band [ConsoleModifiers]::Control) -and $k.Key -eq 'C') { $script:UI.Dirty = $true; return $false }
    }
}

function Show-TypedConfirmModal {
    # Requires the operator to type an exact word. Returns $true/$false.
    param([string]$Title, [object[]]$Lines, [string]$Word)
    $typed = ''
    while ($true) {
        Write-Screen
        $body = New-Object System.Collections.ArrayList
        foreach ($ln in (ConvertTo-ModalLines -Lines $Lines -Width 64)) { [void]$body.Add($ln) }
        [void]$body.Add(@($script:T.Row, ''))
        [void]$body.Add(@($script:T.CtxHi, "Type $Word and press Enter to proceed:"))
        $field = $typed + '_'
        [void]$body.Add(@($script:T.Input, ('  ' + $field)))
        [void](Write-ModalFrame -Title $Title -BodyLines $body.ToArray() -FooterHint 'Enter confirm   Esc cancel' -BorderStyle $script:T.BorderErr)
        $k = Read-ModalKey
        if ($k.Key -eq 'Escape') { $script:UI.Dirty = $true; return $false }
        if (($k.Modifiers -band [ConsoleModifiers]::Control) -and $k.Key -eq 'C') { $script:UI.Dirty = $true; return $false }
        if ($k.Key -eq 'Enter') {
            $script:UI.Dirty = $true
            return ($typed -ceq $Word)
        }
        if ($k.Key -eq 'Backspace') {
            if ($typed.Length -gt 0) { $typed = $typed.Substring(0, $typed.Length - 1) }
            continue
        }
        if ($k.KeyChar -and -not [char]::IsControl($k.KeyChar) -and $typed.Length -lt 32) {
            $typed += $k.KeyChar
        }
    }
}

function Show-InputModal {
    # Free-text input. Returns the string, or $null when cancelled.
    param([string]$Title, [string]$Prompt, [string]$Default = '')
    $typed = $Default
    while ($true) {
        Write-Screen
        $body = New-Object System.Collections.ArrayList
        foreach ($ln in (ConvertTo-ModalLines -Lines @($Prompt) -Width 60)) { [void]$body.Add($ln) }
        [void]$body.Add(@($script:T.Row, ''))
        $shown = $typed
        if ($shown.Length -gt 58) { $shown = [string]$script:G.Ell + $shown.Substring($shown.Length - 55) }
        [void]$body.Add(@($script:T.Input, ('  ' + $shown + '_')))
        [void](Write-ModalFrame -Title $Title -BodyLines $body.ToArray() -FooterHint 'Enter accept   Esc cancel' -BorderStyle $script:T.Border -MinWidth 66)
        $k = Read-ModalKey
        if ($k.Key -eq 'Escape') { $script:UI.Dirty = $true; return $null }
        if (($k.Modifiers -band [ConsoleModifiers]::Control) -and $k.Key -eq 'C') { $script:UI.Dirty = $true; return $null }
        if ($k.Key -eq 'Enter') { $script:UI.Dirty = $true; return $typed }
        if ($k.Key -eq 'Backspace') {
            if ($typed.Length -gt 0) { $typed = $typed.Substring(0, $typed.Length - 1) }
            continue
        }
        if ($k.KeyChar -and -not [char]::IsControl($k.KeyChar) -and $typed.Length -lt 400) {
            $typed += $k.KeyChar
        }
    }
}

function Show-ReportModal {
    # Scrollable result list. $Lines are strings or @($style,$text) pairs.
    param([string]$Title, [object[]]$Lines, [string]$Hint = 'Up/Down scroll   Enter close')
    $norm = ConvertTo-ModalLines -Lines $Lines -Width 72
    $scroll = 0
    while ($true) {
        Write-Screen
        $geo = Write-ModalFrame -Title $Title -BodyLines $norm -FooterHint $Hint -BorderStyle $script:T.Border -MinWidth 78 -BodyScroll $scroll
        $k = Read-ModalKey
        switch ($k.Key) {
            'Enter'     { $script:UI.Dirty = $true; return }
            'Escape'    { $script:UI.Dirty = $true; return }
            'UpArrow'   { if ($scroll -gt 0) { $scroll-- } }
            'DownArrow' { if ($geo.Scrollable -and $scroll -lt ($geo.Total - $geo.BodyH)) { $scroll++ } }
            'PageUp'    { $scroll = [Math]::Max(0, $scroll - $geo.BodyH) }
            'PageDown'  { if ($geo.Scrollable) { $scroll = [Math]::Min([Math]::Max(0,$geo.Total - $geo.BodyH), $scroll + $geo.BodyH) } }
            'Home'      { $scroll = 0 }
            'End'       { if ($geo.Scrollable) { $scroll = [Math]::Max(0, $geo.Total - $geo.BodyH) } }
        }
    }
}

function Start-LoadSpinner {
    # Animates the indeterminate-progress spinner from a background runspace
    # so it keeps moving while the main thread is blocked inside a cmdlet
    # pipeline (e.g. waiting for the first page of Get-Mailbox results).
    # Write-ProgressModal publishes the spinner cell coordinates into State;
    # X = 0 means hidden. Each [Console]::Write is a single synchronized call
    # writing a complete, absolutely positioned sequence, so the background
    # writes never tear against the main thread's full-modal repaints.
    if ($script:Spinner) { return }
    try {
        $state = [hashtable]::Synchronized(@{
            X = 0; Y = 0; Run = $true
            Style = [string]$script:T.Row; Reset = [string]$script:T.Reset
        })
        $ps = [powershell]::Create()
        [void]$ps.AddScript({
            param($state)
            $esc = [char]27
            $frames = '|', '/', '-', '\'
            while ($state.Run) {
                $x = $state.X; $y = $state.Y
                if ($x -gt 0 -and $y -gt 0) {
                    # Same frame formula as Write-ProgressModal so background
                    # ticks and full repaints stay in phase.
                    $f = $frames[[int](([Environment]::TickCount -band 0x7FFFFFFF) / 120) % 4]
                    [Console]::Write(('{0}[{1};{2}H{3}{4}{5}' -f $esc, $y, $x, $state.Style, $f, $state.Reset))
                }
                Start-Sleep -Milliseconds 120
            }
        }).AddArgument($state)
        $script:Spinner = @{ PS = $ps; Handle = $ps.BeginInvoke(); State = $state }
    } catch {
        Write-SoaLog -Message ("Load spinner unavailable: {0}" -f $_.Exception.Message) -Level WARN
        $script:Spinner = $null
    }
}

function Stop-LoadSpinner {
    # Idempotent; joins the runspace so no stray writes can land after return.
    if (-not $script:Spinner) { return }
    $sp = $script:Spinner
    $script:Spinner = $null
    try {
        $sp.State.Run = $false
        [void]$sp.PS.EndInvoke($sp.Handle)
        $sp.PS.Dispose()
    } catch { }
}

function Write-ProgressModal {
    # Stateless render of a progress modal; caller invokes repeatedly.
    # Total > 0  : determinate - percent bar, "Processing X of Y".
    # Total <= 0 : indeterminate - marquee bar + spinner, "Retrieved X so far";
    #              used while streaming results whose total is not known upfront.
    param([string]$Title, [int]$Done, [int]$Total, [string]$Label, [int]$Ok, [int]$Failed)
    $t = $script:T; $g = $script:G
    $innerW = 60
    $barW = $innerW - 7
    $hint = 'working...'
    if ($Total -gt 0) {
        $pct = [int](100 * $Done / $Total)
        $fill = [int]($barW * $Done / $Total)
        if ($fill -gt $barW) { $fill = $barW }
        $bar = $t.BarOn + ([string]$g.BarOn * $fill) + $t.BarOff + ([string]$g.BarOff * ($barW - $fill)) + $t.Reset + $t.Row + (' {0,3}%' -f $pct)
        $head = "Processing $Done of $Total"
    } else {
        # Marquee: short segment bouncing across the bar; frame from the clock
        # so every repaint advances it. Suffix is 5 visible chars, like ' 100%'.
        $segW = 8
        $span = $barW - $segW
        $frame = [int](([Environment]::TickCount -band 0x7FFFFFFF) / 120)
        $phase = $frame % (2 * $span)
        $pos = $phase
        if ($phase -gt $span) { $pos = (2 * $span) - $phase }
        $spin = @('|', '/', '-', '\')[$frame % 4]
        $bar = $t.BarOff + ([string]$g.BarOff * $pos) + $t.BarOn + ([string]$g.BarOn * $segW) + $t.BarOff + ([string]$g.BarOff * ($span - $pos)) + $t.Reset + $t.Row + ('   {0} ' -f $spin)
        $head = "Retrieved $Done so far"
        if ($Done -le 0) { $head = 'Working - this can take a while...' }
        $hint = 'Esc cancels - working...'
    }
    $body = @(
        @($t.Row,   $head),
        @($t.Row,   ''),
        @('RAWBAR', $bar),
        @($t.Row,   ''),
        @($t.CtxHi, (Get-PadCell $Label $innerW)),
        @($t.Good,  ("  OK: $Ok    " )),
        @($t.Danger,("  Failed: $Failed"))
    )
    # RAWBAR lines carry their own styling; Write-ModalFrame pads plain text,
    # so pre-pad: the bar already has fixed visible width (barW + 5).
    $norm = New-Object System.Collections.ArrayList
    foreach ($pair in $body) {
        if ($pair[0] -eq 'RAWBAR') {
            [void]$norm.Add(@('', ''))  # placeholder; replaced below
        } else {
            [void]$norm.Add($pair)
        }
    }
    # Render frame manually to keep the styled bar intact
    $size = Get-ConsoleSize; $W = $size[0]; $H = $size[1]
    $boxW = [Math]::Min(66, $W - 4); $innerBox = $boxW - 4
    $boxH = $body.Count + 4
    $x = [Math]::Max(1, [int](($W - $boxW) / 2) + 1)
    $y = [Math]::Max(1, [int](($H - $boxH) / 2) + 1)
    # Publish the spinner cell to the background spinner runspace (if any):
    # the suffix slot after the marquee bar. Determinate modals hide it.
    if ($script:Spinner) {
        if ($Total -gt 0) {
            $script:Spinner.State.X = 0
        } else {
            $script:Spinner.State.Y = $y + 3
            $script:Spinner.State.X = $x + 2 + $barW + 3
        }
    }
    $sb = New-Object System.Text.StringBuilder
    $hChar = [string]$g.H
    $tt = " $Title "
    $dashTotal = $boxW - 2 - $tt.Length
    if ($dashTotal -lt 0) { $tt = Get-PadCell $tt ($boxW - 2); $dashTotal = 0 }
    [void]$sb.Append("$script:ESC[$y;$($x)H").Append($t.Border).Append([string]$g.TL).Append($hChar * 1)
    [void]$sb.Append($t.ModalTitle).Append($tt).Append($t.Reset).Append($t.Border)
    [void]$sb.Append($hChar * [Math]::Max(0,($dashTotal - 1))).Append([string]$g.TR).Append($t.Reset)
    $row = $y + 1
    foreach ($pair in $body) {
        [void]$sb.Append("$script:ESC[$row;$($x)H").Append($t.Border).Append([string]$g.V).Append($t.Reset).Append(' ')
        if ($pair[0] -eq 'RAWBAR') {
            [void]$sb.Append([string]$pair[1])
            $visLen = $barW + 5
            if ($visLen -lt $innerBox) { [void]$sb.Append(' ' * ($innerBox - $visLen)) }
        } else {
            [void]$sb.Append([string]$pair[0]).Append((Get-PadCell ([string]$pair[1]) $innerBox)).Append($t.Reset)
        }
        [void]$sb.Append(' ').Append($t.Border).Append([string]$g.V).Append($t.Reset)
        $row++
    }
    [void]$sb.Append("$script:ESC[$row;$($x)H").Append($t.Border).Append([string]$g.V).Append($t.Reset).Append(' ')
    [void]$sb.Append($t.Muted).Append((Get-PadCell $hint $innerBox -AlignRight)).Append($t.Reset)
    [void]$sb.Append(' ').Append($t.Border).Append([string]$g.V).Append($t.Reset)
    $row++
    [void]$sb.Append("$script:ESC[$row;$($x)H").Append($t.Border).Append([string]$g.BL).Append($hChar * ($boxW - 2)).Append([string]$g.BR).Append($t.Reset)
    [Console]::Write($sb.ToString())
}

function Show-HelpModal {
    $t = $script:T
    $lines = @(
        @($t.ModalTitle, 'Navigation'),
        @($t.Row, '  Tab / Shift+Tab      switch tab            1-5  jump to tab'),
        @($t.Row, '  Up / Down            move cursor           PgUp / PgDn  page'),
        @($t.Row, '  Home / End           jump to first / last entry'),
        @($t.CtxHi, '  Q  or  Ctrl+C        quit the application'),
        @($t.Row, ''),
        @($t.ModalTitle, 'Selection'),
        @($t.Row, '  Space                toggle selection on current row'),
        @($t.Row, '  A                    select everything in the current view'),
        @($t.Row, '  N                    clear all selections'),
        @($t.Row, ''),
        @($t.ModalTitle, 'List tools'),
        @($t.Row, '  /                    live search (Enter keep, Esc clear)'),
        @($t.Row, '  F                    cycle status filter All/Cloud/On-prem/Pending'),
        @($t.Row, '  S                    cycle sort column     D  flip sort direction'),
        @($t.Row, '  R                    reload data from the service'),
        @($t.Row, '  Esc                  cancel a load in progress'),
        @($t.Row, '  E                    export current view to CSV'),
        @($t.Row, '  I                    import CSV/TXT and select matching entries'),
        @($t.Row, ''),
        @($t.ModalTitle, 'Conversions'),
        @($t.Row, '  C                    convert selection to CLOUD managed SOA'),
        @($t.Row, '  O                    convert selection back to ON-PREM managed SOA'),
        @($t.Row, '                       (attribute backup JSON is written first;'),
        @($t.Row, '                       group rollbacks run a cloud-member check)'),
        @($t.Row, ''),
        @($t.ModalTitle, 'Organization tab'),
        @($t.Row, '  E                    enable cloud-managed default for new mailboxes'),
        @($t.Row, '  D                    revert to server-managed default'),
        @($t.Row, ''),
        @($t.ModalTitle, 'Required roles (PIM users: activate BEFORE connecting)'),
        @($t.Row, '  Mailboxes            Exchange Admin (or Hybrid Identity / Global)'),
        @($t.Row, '  Groups / Contacts    Hybrid Identity Admin + consented Graph scopes'),
        @($t.Row, '  Organization         Hybrid Identity or Global Administrator'),
        @($t.Row, '  Activated late?      W disconnects - reconnect for a fresh token'),
        @($t.Row, ''),
        @($t.ModalTitle, 'Misc'),
        @($t.Row, '  W                    disconnect EXO + Graph sessions'),
        @($t.Row, '  ?                    this help'),
        @($t.Row, '  Q  or  Ctrl+C        quit the application'),
        @($t.Row, ''),
        @($t.Muted, ("  Log file: " + $script:LogFile)),
        @($t.Muted, ("  Backups : " + $script:BackupDir)),
        @($t.Muted, ("  Exports : " + $script:ExportDir))
    )
    Show-ReportModal -Title "Help - Exchange SOA Manager v$script:Version" -Lines $lines
}

#endregion

# ============================================================================
#region Data helpers
# ============================================================================

function Get-PropSafe {
    # StrictMode-safe property access for dynamic/API objects.
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Confirm-SoaDirectory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        [void](New-Item -ItemType Directory -Path $Path -Force)
    }
}

function Test-SoaModule {
    param([string]$Name)
    return ($null -ne (Get-Module -ListAvailable -Name $Name | Select-Object -First 1))
}

function Install-SoaModule {
    param([string]$Name)
    $ok = Show-ConfirmModal -Title 'Module required' -Lines @(
        "The PowerShell module '$Name' is not installed.",
        '',
        'Install it now for the current user?',
        @($script:T.Muted, "Install-Module $Name -Scope CurrentUser")
    )
    if (-not $ok) { return $false }
    Invoke-OnMainBuffer -Action {
        Write-Host ''
        Write-Host "Installing $Name (CurrentUser scope)..." -ForegroundColor Cyan
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Host "Installed $Name." -ForegroundColor Green
        } catch {
            Write-Host "Install failed: $($_.Exception.Message)" -ForegroundColor Red
            Start-Sleep -Seconds 3
        }
    }
    $success = Test-SoaModule -Name $Name
    if ($success) { Write-SoaLog -Message "Installed module $Name." -Level OK }
    else { Write-SoaLog -Message "Failed to install module $Name." -Level ERROR }
    return $success
}

function Test-RsatAd {
    return ($null -ne (Get-Module -ListAvailable -Name 'ActiveDirectory' | Select-Object -First 1))
}

function Install-RsatAd {
    # RSAT ActiveDirectory is a Windows capability/feature, not a Gallery module.
    # Client SKUs (ProductType 1) use Add-WindowsCapability; servers use the feature.
    $isClient = $true
    try { $isClient = ([int](Get-CimInstance Win32_OperatingSystem).ProductType -eq 1) } catch { $isClient = $true }
    $cmd = if ($isClient) {
        'Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'
    } else {
        'Install-WindowsFeature RSAT-AD-PowerShell'
    }
    $ok = Show-ConfirmModal -Title 'RSAT required' -Lines @(
        'The RSAT ActiveDirectory PowerShell module is required for the forward audit.',
        '',
        'Install it now? (requires an elevated / admin session)',
        @($script:T.Muted, "  $cmd")
    )
    if (-not $ok) { return $false }
    Invoke-OnMainBuffer -Action {
        Write-Host ''
        Write-Host 'Installing RSAT ActiveDirectory tools...' -ForegroundColor Cyan
        try {
            if ($isClient) {
                Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -ErrorAction Stop | Out-Null
            } else {
                Install-WindowsFeature -Name 'RSAT-AD-PowerShell' -ErrorAction Stop | Out-Null
            }
            Write-Host 'Installed RSAT ActiveDirectory tools.' -ForegroundColor Green
        } catch {
            Write-Host "Install failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host 'You may need an elevated session. Run the command shown above manually.' -ForegroundColor Yellow
            Start-Sleep -Seconds 3
        }
    }
    $success = Test-RsatAd
    if ($success) { Write-SoaLog -Message 'Installed RSAT ActiveDirectory module.' -Level OK }
    else { Write-SoaLog -Message 'Failed to install RSAT ActiveDirectory module.' -Level ERROR }
    return $success
}

function Test-AuditPrerequisite {
    # Returns Ok=$true when the forward audit can run (or in demo mode).
    if ($script:DemoMode) { return [pscustomobject]@{ Ok = $true; Reason = '' } }
    if (-not $script:IsWin) {
        return [pscustomobject]@{ Ok = $false; Reason = 'The forward audit requires a domain-joined Windows host. Active Directory is not reachable from this platform.' }
    }
    if (-not (Test-RsatAd)) {
        if (-not (Install-RsatAd)) {
            return [pscustomobject]@{ Ok = $false; Reason = 'The RSAT ActiveDirectory module is required and is not installed.' }
        }
    }
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        [void](Get-ADDomain -ErrorAction Stop)
    } catch {
        return [pscustomobject]@{ Ok = $false; Reason = ('Active Directory is not reachable: ' + $_.Exception.Message) }
    }
    return [pscustomobject]@{ Ok = $true; Reason = '' }
}

#endregion

# ============================================================================
#region Connections
# ============================================================================

function Connect-ExoService {
    if ($script:Conn.Exo) { return $true }
    if ($script:DemoMode) {
        $script:Conn.Exo = $true
        $script:Conn.ExoAccount = 'demo-admin@contoso.com'
        Write-SoaLog -Message 'Demo: simulated Exchange Online connection.' -Level OK
        return $true
    }
    if (-not (Test-SoaModule -Name 'ExchangeOnlineManagement')) {
        if (-not (Install-SoaModule -Name 'ExchangeOnlineManagement')) { return $false }
    }
    # MSAL assembly conflict root cause:
    # ExchangeOnlineManagement and Microsoft.Graph.Authentication ship different
    # versions of Microsoft.Identity.Client.dll. .NET loads an assembly by
    # simple name once per process, so whichever module authenticates first pins
    # its MSAL and the other module may throw MissingMethodException.
    # Prior workarounds (Graph-first, EXO-first, dummy credential warmup) all
    # break under newer releases because the DLL signatures keep diverging.
    #
    # Fix: leave Graph completely out of this process. Microsoft Graph calls run
    # in a dedicated child PowerShell process (Start-GraphWorker). EXO owns the
    # parent process MSAL context; the child process owns Graph's MSAL context.
    # They communicate over stdin/stdout JSON envelopes.
    Write-SoaLog -Message 'Connecting to Exchange Online...'
    $script:LastConnectError = $null
    Invoke-OnMainBuffer -Action {
        Write-Host ''
        Write-Host 'Connecting to Exchange Online - complete sign-in in your browser...' -ForegroundColor Cyan
        Write-Host '  Required role : Exchange Administrator (mailbox SOA; Hybrid Identity / Global Admin also work)' -ForegroundColor DarkGray
        Write-Host '                  Hybrid Identity or Global Administrator (tenant-wide default)' -ForegroundColor DarkGray
        Write-Host '  Using PIM?    : activate the role BEFORE completing sign-in.' -ForegroundColor Yellow
        try {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        } catch {
            $script:LastConnectError = $_.Exception.Message
        }
    }
    # Verify the session actually exists.
    $info = $null
    try { $info = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1 } catch { $info = $null }
    if ($null -eq $info) {
        $msg = 'Could not establish an Exchange Online session.'
        if ($script:LastConnectError) { $msg = [string]$script:LastConnectError }
        Write-SoaLog -Message "EXO connection failed: $msg" -Level ERROR
        Show-MsgModal -Title 'Connection failed' -Lines @($msg) -Kind Error
        return $false
    }
    $script:Conn.Exo = $true
    $acct = Get-PropSafe $info 'UserPrincipalName'
    if ($acct) { $script:Conn.ExoAccount = [string]$acct }
    Write-SoaLog -Message ("Connected to Exchange Online as {0}." -f $script:Conn.ExoAccount) -Level OK
    return $true
}


function Connect-GraphService {
    if ($script:Conn.Graph) { return $true }
    if ($script:DemoMode) {
        $script:Conn.Graph = $true
        $script:Conn.GraphAccount = 'demo-admin@contoso.com'
        Write-SoaLog -Message 'Demo: simulated Microsoft Graph connection.' -Level OK
        return $true
    }
    if (-not (Test-SoaModule -Name 'Microsoft.Graph.Authentication')) {
        if (-not (Install-SoaModule -Name 'Microsoft.Graph.Authentication')) { return $false }
    }
    Write-SoaLog -Message 'Starting out-of-process Microsoft Graph worker...'
    $script:LastConnectError = $null
    $scopes = @(
        'Group.Read.All'
        'GroupMember.Read.All'
        'User.Read.All'
        'Group-OnPremisesSyncBehavior.ReadWrite.All'
        'OrgContact.Read.All'
        'Contacts-OnPremisesSyncBehavior.ReadWrite.All'
    )
    # The interactive browser sign-in must happen in the child process. We
    # leave the TUI so the device-code / browser prompt is visible on the main
    # console, then launch the worker.
    $script:Conn.Graph = $false
    Invoke-OnMainBuffer -Action {
        Write-Host ''
        Write-Host 'Connecting to Microsoft Graph - complete sign-in in your browser...' -ForegroundColor Cyan
        Write-Host '  Required role : Hybrid Identity Administrator' -ForegroundColor DarkGray
        Write-Host '  Using PIM?    : activate the role BEFORE completing sign-in.' -ForegroundColor Yellow
        Write-Host 'Requested scopes:' -ForegroundColor DarkGray
        foreach ($s in $scopes) { Write-Host "  $s" -ForegroundColor DarkGray }
        try {
            $started = Start-GraphWorker -Scopes $scopes
            if ($started) { $script:Conn.Graph = $true }
        } catch {
            $script:LastConnectError = $_.Exception.Message
            Write-Host "Graph worker failed: $($_.Exception.Message)" -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    if (-not $script:Conn.Graph) {
        Write-SoaLog -Message 'Graph worker connection failed or was cancelled.' -Level ERROR
        Show-MsgModal -Title 'Connection failed' -Lines @(
            'Could not establish a Microsoft Graph session.',
            '',
            'Note: the Group/Contacts OnPremisesSyncBehavior scopes require admin consent in the tenant.'
        ) -Kind Error
        return $false
    }
    $script:Conn.GraphAccount = $script:GraphWorker.Account
    Write-SoaLog -Message ("Connected to Microsoft Graph as {0}." -f $script:Conn.GraphAccount) -Level OK
    return $true
}

function Stop-GraphWorker {
    # Idempotent cleanup of the out-of-process Graph worker.
    $gw = $script:GraphWorker
    if ($gw.StdIn) {
        try { $gw.StdIn.Close() } catch { }
        try { $gw.StdIn.Dispose() } catch { }
        $gw.StdIn = $null
    }
    if ($gw.StdOut) {
        try { $gw.StdOut.Close() } catch { }
        try { $gw.StdOut.Dispose() } catch { }
        $gw.StdOut = $null
    }
    if ($gw.StdErr) {
        try { $gw.StdErr.Close() } catch { }
        try { $gw.StdErr.Dispose() } catch { }
        $gw.StdErr = $null
    }
    if ($gw.Process) {
        try {
            if (-not $gw.Process.HasExited) {
                $gw.Process.Kill()
                $gw.Process.WaitForExit(2000) | Out-Null
            }
        } catch { }
        try { $gw.Process.Dispose() } catch { }
        $gw.Process = $null
    }
    if ($gw.WorkerPath) {
        try { Remove-Item -LiteralPath $gw.WorkerPath -ErrorAction SilentlyContinue } catch { }
        $gw.WorkerPath = $null
    }
    $gw.LastError = ''
}

function Start-GraphWorker {
    param([string[]]$Scopes)
    Stop-GraphWorker
    $gw = $script:GraphWorker
    # Write the worker to a temporary script file. A child file avoids
    # Base64 length limits and lets us pass scopes via a normal parameter.
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('SOA-GraphWorker_' + [Guid]::NewGuid().ToString('n') + '.ps1')
    @'
param([string]$ScopeList)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$scopes = $ScopeList -split '\|'
Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
$ctx = Get-MgContext
$acct = ''
if ($ctx) { $acct = [string]$ctx.Account }
@{ type='ready'; account=$acct } | ConvertTo-Json -Compress
$in = [Console]::In
while ($null -ne ($line = $in.ReadLine())) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $job = $line | ConvertFrom-Json
    try {
        switch ($job.method) {
            'GET' {
                $hdr = @{ }
                if ($job.headers) { foreach ($h in $job.headers.PSObject.Properties) { $hdr[$h.Name] = $h.Value } }
                $resp = Invoke-MgGraphRequest -Method GET -Uri $job.uri -Headers $hdr -OutputType PSObject -ErrorAction Stop
                @{ type='ok'; id=$job.id; value=$resp } | ConvertTo-Json -Depth 10 -Compress
            }
            'PATCH' {
                [void](Invoke-MgGraphRequest -Method PATCH -Uri $job.uri -Body $job.body -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop)
                @{ type='ok'; id=$job.id } | ConvertTo-Json -Compress
            }
            'POST' {
                $body = $job.body
                $ct = $job.contentType
                if (-not $ct) { $ct = 'application/json' }
                $resp = Invoke-MgGraphRequest -Method POST -Uri $job.uri -Body $body -ContentType $ct -OutputType PSObject -ErrorAction Stop
                @{ type='ok'; id=$job.id; value=$resp } | ConvertTo-Json -Depth 10 -Compress
            }
            default { throw "Unknown method $($job.method)" }
        }
    } catch {
        $msg = $_.Exception.Message
        if ($_.Exception -is [System.Management.Automation.MethodInvocationException] -and $_.Exception.InnerException) {
            $msg = $_.Exception.InnerException.Message
        }
        @{ type='err'; id=$job.id; message=$msg } | ConvertTo-Json -Compress
    }
}
'@ | Set-Content -LiteralPath $tmp -Encoding UTF8
    $gw.WorkerPath = $tmp
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = if ($PSVersionTable.PSEdition -eq 'Core' -and (Get-Command 'pwsh' -ErrorAction SilentlyContinue)) { 'pwsh' } else { 'powershell' }
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$tmp`" -ScopeList `"$($Scopes -join '|')`""
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $proc = [System.Diagnostics.Process]::Start($psi)
    $gw.Process = $proc
    $gw.StdIn = $proc.StandardInput
    $gw.StdOut = $proc.StandardOutput
    $gw.StdErr = $proc.StandardError
    $gw.Account = ''
    # Read the ready line (or error). The sign-in is synchronous in the child,
    # so this blocks until the user completes auth in the browser.
    $ready = $gw.StdOut.ReadLine()
    if (-not $ready) {
        Stop-GraphWorker
        throw 'Graph worker exited before signalling readiness.'
    }
    $readyObj = $ready | ConvertFrom-Json
    if ($readyObj.type -ne 'ready') {
        $err = $readyObj.message
        if (-not $err) { $err = 'Graph worker failed during authentication.' }
        Stop-GraphWorker
        throw $err
    }
    $gw.Account = [string]$readyObj.account
    return $true
}

function Invoke-GraphWorker {
    param([hashtable]$Job)
    $gw = $script:GraphWorker
    if (-not $gw.Process -or $gw.Process.HasExited) {
        throw 'Graph worker is not running. Reconnect to Microsoft Graph first.'
    }
    if (-not $Job.ContainsKey('id')) { $Job['id'] = [Guid]::NewGuid().ToString('n') }
    $line = $Job | ConvertTo-Json -Depth 5 -Compress
    $gw.StdIn.WriteLine($line)
    $resp = $gw.StdOut.ReadLine()
    if (-not $resp) {
        $stderr = ''
        try { $stderr = $gw.StdErr.ReadToEnd() } catch { }
        Stop-GraphWorker
        throw ('Graph worker closed the response stream. {0}' -f $stderr)
    }
    $obj = $resp | ConvertFrom-Json
    if ($obj.type -eq 'err') {
        throw [string]$obj.message
    }
    return $obj.value
}

function Disconnect-AllServices {
    if ($script:DemoMode) { return }
    if ($script:Conn.Exo) {
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
        Write-SoaLog -Message 'Disconnected from Exchange Online.'
    }
    if ($script:Conn.Graph) {
        Stop-GraphWorker
        Write-SoaLog -Message 'Disconnected from Microsoft Graph.'
    }
    $script:Conn.Exo = $false; $script:Conn.Graph = $false
    $script:Conn.ExoAccount = ''; $script:Conn.GraphAccount = ''
}

#endregion

# ============================================================================
#region Graph REST helpers
# ============================================================================

function Invoke-GraphGetAll {
    # Paged GET; returns array of PSObjects from .value.
    # -OnPage (optional) is invoked with the cumulative item count after each page.
    param([string]$Uri, [switch]$Advanced, [scriptblock]$OnPage)
    $headers = @{}
    if ($Advanced) { $headers['ConsistencyLevel'] = 'eventual' }
    $out = New-Object System.Collections.ArrayList
    $next = $Uri
    while ($next) {
        $resp = Invoke-GraphWorker -Job @{ method='GET'; uri=$next; headers=$headers }
        $val = Get-PropSafe $resp 'value'
        if ($null -ne $val) { foreach ($v in @($val)) { [void]$out.Add($v) } }
        $next = Get-PropSafe $resp '@odata.nextLink'
        if ($OnPage) { & $OnPage $out.Count }
    }
    return ,$out.ToArray()
}

function Get-SyncBehaviorMap {
    # Batched lookup of isCloudManaged for a set of object ids.
    # $Resource: 'groups' or 'contacts'. Returns hashtable id -> [bool]
    # -Progress (optional) is invoked per batch with (count, label).
    param([string[]]$Ids, [string]$Resource, [scriptblock]$Progress)
    $map = @{}
    if (-not $Ids -or $Ids.Count -eq 0) { return $map }
    $chunkSize = 20
    for ($i = 0; $i -lt $Ids.Count; $i += $chunkSize) {
        $end = [Math]::Min($i + $chunkSize - 1, $Ids.Count - 1)
        $chunk = $Ids[$i..$end]
        $requests = New-Object System.Collections.ArrayList
        $n = 1
        foreach ($id in $chunk) {
            [void]$requests.Add(@{
                id     = "$n"
                method = 'GET'
                url    = ('/' + $Resource + '/' + $id + '/onPremisesSyncBehavior?$select=isCloudManaged')
            })
            $n++
        }
        $body = @{ requests = $requests.ToArray() } | ConvertTo-Json -Depth 4
        $resp = Invoke-GraphWorker -Job @{ method='POST'; uri='https://graph.microsoft.com/v1.0/$batch'; body=$body; contentType='application/json' }
        $responses = Get-PropSafe $resp 'responses'
        if ($null -eq $responses) { continue }
        foreach ($r in @($responses)) {
            $rid = [int](Get-PropSafe $r 'id')
            $status = [int](Get-PropSafe $r 'status')
            $objId = $chunk[$rid - 1]
            if ($status -eq 200) {
                $rbody = Get-PropSafe $r 'body'
                $icm = Get-PropSafe $rbody 'isCloudManaged'
                $map[$objId] = [bool]$icm
            } else {
                Write-SoaLog -Message ("onPremisesSyncBehavior lookup failed for {0}/{1} (HTTP {2})." -f $Resource, $objId, $status) -Level WARN
            }
        }
        if ($Progress) {
            & $Progress ($end + 1) ('Checking SOA state - {0} of {1} converted {2}' -f ($end + 1), $Ids.Count, $Resource)
        }
    }
    return $map
}

#endregion

# ============================================================================
#region Demo data
# ============================================================================

$script:DemoFirst = @('Anna','Bjorn','Carla','David','Elena','Felix','Greta','Hugo','Ines','Jonas','Klara','Lars','Mona','Nils','Olga','Per','Quinn','Rosa','Sven','Tara','Ulf','Vera','Wim','Xena','Yara','Zane','Astrid','Bruno','Celine','Dario','Edith','Frans','Gilda','Henrik','Iris','Joost','Kira','Liam','Mara','Noor')
$script:DemoLast  = @('Andersen','Bergstrom','Carlsen','Dahl','Eriksen','Fischer','Gruber','Hansen','Iversen','Jansen','Koch','Lindgren','Meyer','Nielsen','Olsen','Petersen','Qvist','Rasmussen','Sorensen','Thomsen','Ulrich','Vogel','Weber','Xander','Ylva','Zimmermann','Abel','Brandt','Clausen','Dietrich','Engel','Falk','Gerber','Holm','Ibsen','Jung','Krause','Lund','Moller','Nygaard')

function New-DemoMailboxes {
    $rng = New-Object System.Random(42)
    $items = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt 340; $i++) {
        $fn = $script:DemoFirst[$rng.Next(0, $script:DemoFirst.Count)]
        $ln = $script:DemoLast[$rng.Next(0, $script:DemoLast.Count)]
        $name = "$fn $ln"
        $upn = ('{0}.{1}{2}@contoso.com' -f $fn.ToLower(), $ln.ToLower(), $i)
        $kind = 'UserMailbox'
        if (($i % 9) -eq 0) { $kind = 'SharedMailbox'; $name = "SM $ln team $i"; $upn = "shared.$($ln.ToLower())$i@contoso.com" }
        elseif (($i % 23) -eq 0) { $kind = 'RoomMailbox'; $name = "Room $ln $i"; $upn = "room.$($ln.ToLower())$i@contoso.com" }
        $soa = 'OnPrem'
        if ($rng.Next(0, 100) -lt 30) { $soa = 'Cloud' }
        [void]$items.Add([pscustomobject]@{
            Type='Mailbox'; Id=$upn; Name=$name; Email=$upn; Detail=$kind; Soa=$soa; Selected=$false
            Raw=[pscustomobject]@{ DisplayName=$name; UserPrincipalName=$upn; PrimarySmtpAddress=$upn; RecipientTypeDetails=$kind; IsDirSynced=$true }
        })
    }
    return ,($items.ToArray() | Sort-Object -Property Name)
}

function New-DemoGroups {
    $rng = New-Object System.Random(1337)
    $kinds = @('Security','Distribution','Mail-sec','Security','Distribution')
    $words = @('Finance','HR','Sales','Engineering','Support','Legal','Marketing','Ops','Research','Field','Branch','Procurement','Payroll','Helpdesk','Admins','Auditors','Interns','Leads')
    $items = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt 72; $i++) {
        $w1 = $words[$rng.Next(0, $words.Count)]
        $w2 = $words[$rng.Next(0, $words.Count)]
        $name = "GRP $w1 $w2 $i"
        $mailNick = ('grp.{0}.{1}{2}' -f $w1.ToLower(), $w2.ToLower(), $i)
        $kind = $kinds[$i % $kinds.Count]
        $mail = ''
        if ($kind -ne 'Security') { $mail = "$mailNick@contoso.com" }
        $roll = $rng.Next(0, 100)
        $soa = 'OnPrem'
        if ($roll -lt 22) { $soa = 'Cloud' } elseif ($roll -lt 30) { $soa = 'Pending' }
        [void]$items.Add([pscustomobject]@{
            Type='Group'; Id=([guid]::NewGuid().ToString()); Name=$name; Email=$mail; Detail=$kind; Soa=$soa; Selected=$false
            Raw=[pscustomobject]@{ displayName=$name; mail=$mail; mailEnabled=($kind -ne 'Security'); securityEnabled=($kind -ne 'Distribution'); onPremisesSyncEnabled=($soa -eq 'OnPrem') }
        })
    }
    return ,($items.ToArray() | Sort-Object -Property Name)
}

function New-DemoContacts {
    $rng = New-Object System.Random(7)
    $items = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt 26; $i++) {
        $fn = $script:DemoFirst[$rng.Next(0, $script:DemoFirst.Count)]
        $ln = $script:DemoLast[$rng.Next(0, $script:DemoLast.Count)]
        $name = "$fn $ln (ext)"
        $mail = ('{0}.{1}@partner{2}.example' -f $fn.ToLower(), $ln.ToLower(), ($i % 6))
        $soa = 'OnPrem'
        if ($rng.Next(0, 100) -lt 20) { $soa = 'Cloud' }
        [void]$items.Add([pscustomobject]@{
            Type='Contact'; Id=([guid]::NewGuid().ToString()); Name=$name; Email=$mail; Detail='Mail contact'; Soa=$soa; Selected=$false
            Raw=[pscustomobject]@{ displayName=$name; mail=$mail; onPremisesSyncEnabled=($soa -eq 'OnPrem') }
        })
    }
    return ,($items.ToArray() | Sort-Object -Property Name)
}

#endregion

# ============================================================================
#region Data fetchers
# ============================================================================

function Format-ElapsedTime {
    param([System.Diagnostics.Stopwatch]$Stopwatch)
    $s = [int]$Stopwatch.Elapsed.TotalSeconds
    if ($s -ge 60) { return ('{0}m {1:d2}s' -f [int][Math]::Floor($s / 60), ($s % 60)) }
    return ('{0}s' -f $s)
}

function Invoke-ExoBackgroundFetch {
    # Runs an EXO command in a background runspace so the main thread stays
    # free to repaint the modal (elapsed time) and honor Esc while the EXO v3
    # cmdlet buffers its entire result set. EXO v3 stores connections in a
    # process-wide static, so the worker runspace reuses the live connection
    # after importing the module. -Tick runs every poll; it may throw
    # OperationCanceledException (Esc), which stops the worker and rethrows.
    #
    # Get-Mailbox is NOT exported by ExchangeOnlineManagement - Connect
    # generates it into a tmpEXO_* module inside the connecting runspace, so
    # the worker must import that module's .psm1 (-ProxyModulePath) to see it.
    param([string]$Command, [scriptblock]$Tick, [string]$ProxyModulePath)
    $ps = [powershell]::Create()
    $bootstrap = 'Import-Module ExchangeOnlineManagement -ErrorAction Stop; '
    if ($ProxyModulePath) {
        $bootstrap += 'Import-Module ''' + $ProxyModulePath + ''' -ErrorAction Stop; '
    }
    [void]$ps.AddScript($bootstrap + $Command)
    $handle = $ps.BeginInvoke()
    try {
        while (-not $handle.IsCompleted) {
            if ($Tick) { & $Tick }
            Start-Sleep -Milliseconds 150
        }
        $out = $ps.EndInvoke($handle)
        # Non-terminating worker errors (e.g. CommandNotFound for a script
        # command) do NOT make EndInvoke throw - they land in the error
        # stream and the result is silently empty. Surface them so the
        # caller's fallback path actually runs.
        if ($ps.Streams.Error.Count -gt 0) { throw $ps.Streams.Error[0].Exception }
        return ,$out
    } catch [System.OperationCanceledException] {
        try { $ps.Stop() } catch { }
        throw
    } finally {
        $ps.Dispose()
    }
}

function Get-MailboxItems {
    # -Progress (optional) is invoked with (count, label) as results stream in.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'False positive: $Progress is used inside the $build scriptblock.')]
    param([scriptblock]$Progress)
    if ($script:DemoMode) { return ,(New-DemoMailboxes) }
    Write-SoaLog -Message 'Retrieving dir-synced mailboxes from Exchange Online...'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $items = New-Object System.Collections.ArrayList
    $stats = @{ Received = 0 }
    # Get-Mailbox has no -Properties parameter (and Get-EXOMailbox, which has
    # one, does not expose IsExchangeCloudManaged), so the full 200+ property
    # object always crosses the wire. Project each one to a slim record as
    # soon as it streams in so the heavy object can be garbage-collected
    # instead of being buffered - and kept alive in Raw - for the session.
    $build = {
        param($mb)
        $stats.Received++
        # EXO v3 cmdlets buffer the entire result set before the pipeline
        # releases anything, so this only fires once everything has arrived.
        # No point reporting a count; keep the call for the Esc-drain in the
        # callback and pass 0 so the modal stays in "can take a while" mode.
        if ($Progress) { & $Progress 0 ('This can take several minutes - {0} elapsed' -f (Format-ElapsedTime $sw)) }
        if (-not [bool](Get-PropSafe $mb 'IsDirSynced')) { return }
        $cloud = [bool](Get-PropSafe $mb 'IsExchangeCloudManaged')
        $slim = [pscustomobject]@{
            DisplayName               = [string](Get-PropSafe $mb 'DisplayName')
            UserPrincipalName         = [string](Get-PropSafe $mb 'UserPrincipalName')
            PrimarySmtpAddress        = [string](Get-PropSafe $mb 'PrimarySmtpAddress')
            RecipientTypeDetails      = [string](Get-PropSafe $mb 'RecipientTypeDetails')
            IsDirSynced               = $true
            IsExchangeCloudManaged    = $cloud
            ExternalDirectoryObjectId = [string](Get-PropSafe $mb 'ExternalDirectoryObjectId')
        }
        $soa = 'OnPrem'
        if ($cloud) { $soa = 'Cloud' }
        $id = $slim.ExternalDirectoryObjectId
        if ([string]::IsNullOrEmpty($id)) { $id = $slim.UserPrincipalName }
        [void]$items.Add([pscustomobject]@{
            Type     = 'Mailbox'
            Id       = $id
            Name     = $slim.DisplayName
            Email    = $slim.PrimarySmtpAddress
            Detail   = $slim.RecipientTypeDetails
            Soa      = $soa
            Selected = $false
            Raw      = $slim
        })
    }
    $tick = {
        if ($Progress) { & $Progress 0 ('This can take several minutes - {0} elapsed' -f (Format-ElapsedTime $sw)) }
    }
    try {
        # Retrieve only dir-synced mailboxes from the server side. This is
        # extremely fast and avoids fetching all cloud-only mailboxes (which
        # can cause timeouts or throttling in large tenants). Fetch runs in a
        # background runspace: EXO v3 buffers everything before returning, so
        # this keeps the main thread free for Esc-cancel and elapsed updates.
        # Filter is single-quoted so the worker passes a literal $true to
        # Exchange instead of interpolating it to 'True' first.
        $cmd = 'Get-Mailbox -Filter ''IsDirSynced -eq $true'' -ResultSize Unlimited -ErrorAction Stop'
        # Get-Mailbox lives in the tmpEXO_* proxy module Connect generated in
        # THIS runspace; hand its path to the worker so it can import it.
        $proxy = Get-Module | Where-Object { $_.Name -like 'tmpEXO*' -or $_.Name -like 'tmp_*' } | Select-Object -First 1
        $proxyPath = ''
        if ($proxy) { $proxyPath = [string]$proxy.Path }
        $raw = Invoke-ExoBackgroundFetch -Command $cmd -Tick $tick -ProxyModulePath $proxyPath
        foreach ($mb in $raw) { & $build $mb }
    } catch [System.OperationCanceledException] {
        throw
    } catch {
        # Fallback for filter issues or a worker runspace that cannot see the
        # EXO connection: blocking main-thread fetch, as before (no Esc).
        Write-SoaLog -Message ("Background filtered Get-Mailbox failed ({0}); trying blocking unrestricted Get-Mailbox." -f $_.Exception.Message) -Level WARN
        $items.Clear()
        $stats.Received = 0
        Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | ForEach-Object { & $build $_ }
    }
    Write-SoaLog -Message ("Retrieved {0} mailboxes; {1} are dir-synced (shown)." -f $stats.Received, $items.Count) -Level OK
    return ,($items.ToArray() | Sort-Object -Property Name)
}

function Get-GroupKind {
    param($G)
    $types = Get-PropSafe $G 'groupTypes'
    if ($types -and (@($types) -contains 'Unified')) { return 'M365' }
    $mailOn = [bool](Get-PropSafe $G 'mailEnabled')
    $secOn  = [bool](Get-PropSafe $G 'securityEnabled')
    if ($mailOn -and $secOn) { return 'Mail-sec' }
    if ($mailOn) { return 'Distribution' }
    if ($secOn) { return 'Security' }
    return 'Group'
}

function Get-GroupItems {
    # -Progress (optional) is invoked with (count, label) as results stream in.
    param([scriptblock]$Progress)
    if ($script:DemoMode) { return ,(New-DemoGroups) }
    Write-SoaLog -Message 'Retrieving groups from Microsoft Graph...'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $phase = @{ Text = 'synced groups' }
    $pageCb = $null
    if ($Progress) {
        # Script-scope functions are not resolvable by name inside GetNewClosure()
        # scriptblocks (see note in Invoke-TabLoad); capture a reference instead.
        $fnElapsed = ${function:Format-ElapsedTime}
        $pageCb = { param($n) & $Progress $n ('{0} {1} received - {2} elapsed' -f $n, $phase.Text, (& $fnElapsed $sw)) }.GetNewClosure()
    }
    $sel = 'id,displayName,mail,mailEnabled,securityEnabled,groupTypes,onPremisesSyncEnabled,onPremisesSecurityIdentifier'
    $base = 'https://graph.microsoft.com/v1.0/groups'

    # 1) Groups still synced from AD (SOA = on-prem)
    $synced = Invoke-GraphGetAll -Uri ($base + '?$filter=onPremisesSyncEnabled eq true&$select=' + $sel + '&$top=999&$count=true') -Advanced -OnPage $pageCb

    # 2) Formerly synced groups (SOA converted, or rolled back awaiting sync)
    $phase.Text = 'converted groups'
    $former = @()
    try {
        $former = Invoke-GraphGetAll -Uri ($base + '?$filter=onPremisesSyncEnabled ne true and onPremisesSecurityIdentifier ne null&$select=' + $sel + '&$top=999&$count=true') -Advanced -OnPage $pageCb
    } catch [System.OperationCanceledException] {
        throw
    } catch {
        Write-SoaLog -Message ("Advanced group filter failed ({0}); enumerating all groups instead." -f $_.Exception.Message) -Level WARN
        $phase.Text = 'groups'
        $all = Invoke-GraphGetAll -Uri ($base + '?$select=' + $sel + '&$top=999') -OnPage $pageCb
        $former = @($all | Where-Object {
            (-not [bool](Get-PropSafe $_ 'onPremisesSyncEnabled')) -and ($null -ne (Get-PropSafe $_ 'onPremisesSecurityIdentifier'))
        })
    }

    $behavior = @{}
    if ($former.Count -gt 0) {
        $ids = @($former | ForEach-Object { [string](Get-PropSafe $_ 'id') })
        $behavior = Get-SyncBehaviorMap -Ids $ids -Resource 'groups' -Progress $Progress
    }

    $items = New-Object System.Collections.ArrayList
    foreach ($g in @($synced)) {
        [void]$items.Add([pscustomobject]@{
            Type='Group'; Id=[string](Get-PropSafe $g 'id'); Name=[string](Get-PropSafe $g 'displayName')
            Email=[string](Get-PropSafe $g 'mail'); Detail=(Get-GroupKind $g); Soa='OnPrem'; Selected=$false; Raw=$g
        })
    }
    foreach ($g in @($former)) {
        $gid = [string](Get-PropSafe $g 'id')
        $soa = 'Pending'
        if ($behavior.ContainsKey($gid) -and $behavior[$gid]) { $soa = 'Cloud' }
        [void]$items.Add([pscustomobject]@{
            Type='Group'; Id=$gid; Name=[string](Get-PropSafe $g 'displayName')
            Email=[string](Get-PropSafe $g 'mail'); Detail=(Get-GroupKind $g); Soa=$soa; Selected=$false; Raw=$g
        })
    }
    Write-SoaLog -Message ("Retrieved {0} AD-linked groups ({1} synced, {2} converted/pending)." -f $items.Count, @($synced).Count, @($former).Count) -Level OK
    return ,($items.ToArray() | Sort-Object -Property Name)
}

function Get-ContactItems {
    # -Progress (optional) is invoked with (count, label) as results stream in.
    param([scriptblock]$Progress)
    if ($script:DemoMode) { return ,(New-DemoContacts) }
    Write-SoaLog -Message 'Retrieving organizational contacts from Microsoft Graph...'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $pageCb = $null
    if ($Progress) {
        # Script-scope functions are not resolvable by name inside GetNewClosure()
        # scriptblocks (see note in Invoke-TabLoad); capture a reference instead.
        $fnElapsed = ${function:Format-ElapsedTime}
        $pageCb = { param($n) & $Progress $n ('{0} contacts received - {1} elapsed' -f $n, (& $fnElapsed $sw)) }.GetNewClosure()
    }
    $sel = 'id,displayName,mail,onPremisesSyncEnabled'
    $all = Invoke-GraphGetAll -Uri ('https://graph.microsoft.com/v1.0/contacts?$select=' + $sel + '&$top=999') -OnPage $pageCb
    $synced  = @($all | Where-Object { [bool](Get-PropSafe $_ 'onPremisesSyncEnabled') })
    $others  = @($all | Where-Object { -not [bool](Get-PropSafe $_ 'onPremisesSyncEnabled') })
    $behavior = @{}
    if ($others.Count -gt 0) {
        $ids = @($others | ForEach-Object { [string](Get-PropSafe $_ 'id') })
        $behavior = Get-SyncBehaviorMap -Ids $ids -Resource 'contacts' -Progress $Progress
    }
    $items = New-Object System.Collections.ArrayList
    foreach ($c in $synced) {
        [void]$items.Add([pscustomobject]@{
            Type='Contact'; Id=[string](Get-PropSafe $c 'id'); Name=[string](Get-PropSafe $c 'displayName')
            Email=[string](Get-PropSafe $c 'mail'); Detail='Mail contact'; Soa='OnPrem'; Selected=$false; Raw=$c
        })
    }
    foreach ($c in $others) {
        $cid = [string](Get-PropSafe $c 'id')
        # Only show contacts that were SOA-converted; skip cloud-born contacts.
        if ($behavior.ContainsKey($cid) -and $behavior[$cid]) {
            [void]$items.Add([pscustomobject]@{
                Type='Contact'; Id=$cid; Name=[string](Get-PropSafe $c 'displayName')
                Email=[string](Get-PropSafe $c 'mail'); Detail='Mail contact'; Soa='Cloud'; Selected=$false; Raw=$c
            })
        }
    }
    Write-SoaLog -Message ("Retrieved {0} SOA-relevant contacts (of {1} org contacts)." -f $items.Count, @($all).Count) -Level OK
    return ,($items.ToArray() | Sort-Object -Property Name)
}

function Get-OrgState {
    if ($script:DemoMode) {
        $script:Org.Loaded = $true
        $script:Org.CloudDefault = $script:DemoOrgCloudDefault
        $script:Org.CheckedAt = Get-Date
        $script:Org.TenantName = 'Contoso (demo)'
        return $true
    }
    try {
        Write-SoaLog -Message 'Reading organization configuration...'
        $oc = Get-OrganizationConfig -ErrorAction Stop
        $script:Org.CloudDefault = [bool](Get-PropSafe $oc 'BlockExchangeProvisioningFromOnPremEnabled')
        $name = Get-PropSafe $oc 'DisplayName'
        if (-not $name) { $name = Get-PropSafe $oc 'Name' }
        $script:Org.TenantName = [string]$name
        $script:Org.CheckedAt = Get-Date
        $script:Org.Loaded = $true
        Write-SoaLog -Message ("Org config read. BlockExchangeProvisioningFromOnPremEnabled = {0}" -f $script:Org.CloudDefault) -Level OK
        return $true
    } catch {
        Write-SoaLog -Message ("Failed to read organization config: {0}" -f $_.Exception.Message) -Level ERROR
        Show-MsgModal -Title 'Error' -Lines @("Failed to read organization config:", $_.Exception.Message) -Kind Error
        return $false
    }
}

function Set-OrgState {
    param([bool]$CloudDefault)
    if ($script:DemoMode) {
        Start-Sleep -Milliseconds 300
        $script:DemoOrgCloudDefault = $CloudDefault
        Write-SoaLog -Message ("Demo: org default set to {0}." -f $CloudDefault) -Level OK
        return $true
    }
    try {
        if ($CloudDefault) {
            Write-SoaLog -Message 'Running Set-OrganizationConfig -ExchangeAttributesCloudManagedByDefault ...'
            Set-OrganizationConfig -ExchangeAttributesCloudManagedByDefault -ErrorAction Stop
        } else {
            Write-SoaLog -Message 'Running Set-OrganizationConfig -ExchangeAttributesServerManagedByDefault ...'
            Set-OrganizationConfig -ExchangeAttributesServerManagedByDefault -ErrorAction Stop
        }
        Write-SoaLog -Message 'Organization config updated.' -Level OK
        return $true
    } catch {
        Write-SoaLog -Message ("Failed to update organization config: {0}" -f $_.Exception.Message) -Level ERROR
        Show-MsgModal -Title 'Error' -Lines @('Failed to update organization config:', $_.Exception.Message) -Kind Error
        return $false
    }
}

#endregion

# ============================================================================
#region Backup, conversion, safety checks
# ============================================================================

$script:MailboxBackupProps = @(
    'DisplayName','UserPrincipalName','PrimarySmtpAddress','Alias','RecipientTypeDetails',
    'HiddenFromAddressListsEnabled','ForwardingAddress','ForwardingSmtpAddress','DeliverToMailboxAndForward',
    'RequireSenderAuthenticationEnabled','ModerationEnabled','MailTip','IsExchangeCloudManaged','WhenChanged'
)

function Get-MailboxBackupRecord {
    # Full attribute snapshot of one mailbox (uses Get-Mailbox for completeness).
    param($Item)
    if ($script:DemoMode) {
        return [ordered]@{ DisplayName=$Item.Name; PrimarySmtpAddress=$Item.Email; Note='demo backup'; CapturedAt=(Get-Date -Format o) }
    }
    $mb = Get-Mailbox -Identity $Item.Id -ErrorAction Stop
    $rec = [ordered]@{}
    foreach ($p in $script:MailboxBackupProps) {
        $v = Get-PropSafe $mb $p
        if ($null -ne $v) { $rec[$p] = [string]$v } else { $rec[$p] = $null }
    }
    $addr = Get-PropSafe $mb 'EmailAddresses'
    if ($null -ne $addr) { $rec['EmailAddresses'] = @(@($addr) | ForEach-Object { $_.ToString() }) }
    $sendOnBehalf = Get-PropSafe $mb 'GrantSendOnBehalfTo'
    if ($null -ne $sendOnBehalf) { $rec['GrantSendOnBehalfTo'] = @(@($sendOnBehalf) | ForEach-Object { $_.ToString() }) }
    for ($i = 1; $i -le 15; $i++) {
        $rec["CustomAttribute$i"] = [string](Get-PropSafe $mb "CustomAttribute$i")
    }
    for ($i = 1; $i -le 5; $i++) {
        $v = Get-PropSafe $mb "ExtensionCustomAttribute$i"
        if ($null -ne $v) { $rec["ExtensionCustomAttribute$i"] = @(@($v) | ForEach-Object { $_.ToString() }) }
    }
    $rec['CapturedAt'] = (Get-Date -Format o)
    return $rec
}

function Save-BackupFile {
    param([object[]]$Records, [string]$Kind)
    if (-not $Records -or $Records.Count -eq 0) { return $null }
    Confirm-SoaDirectory -Path $script:BackupDir
    $file = Join-Path $script:BackupDir ("Backup_{0}_{1}.json" -f $Kind, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $Records | ConvertTo-Json -Depth 8 | Set-Content -Path $file -Encoding UTF8
    Write-SoaLog -Message ("Backup written: {0} ({1} records)." -f $file, $Records.Count) -Level OK
    return $file
}

function Convert-MailboxSoa {
    param($Item, [bool]$ToCloud)
    if ($script:DemoMode) {
        Start-Sleep -Milliseconds (20 + (Get-Random -Maximum 50))
        if ((Get-Random -Maximum 100) -lt 4) { return @{ Ok=$false; Msg='demo: simulated transient error' } }
        return @{ Ok=$true; Msg='converted' }
    }
    try {
        Set-Mailbox -Identity $Item.Id -IsExchangeCloudManaged $ToCloud -ErrorAction Stop
        return @{ Ok=$true; Msg='converted' }
    } catch {
        # EXO REST cmdlets increasingly throw a terminating error whose
        # Exception.Message is empty (real text sits in ErrorDetails.Message)
        # even though the SOA change was applied server-side. Trust the actual
        # state over the error record: re-read the mailbox and, if it already
        # matches the target, report success. ponytail: one confirming re-read
        # on the failure path only - no extra round trip on the happy path.
        $errText = $_.ErrorDetails.Message
        if ([string]::IsNullOrWhiteSpace($errText)) { $errText = $_.Exception.Message }
        try {
            $mb = Get-Mailbox -Identity $Item.Id -ErrorAction Stop
            if ([bool](Get-PropSafe $mb 'IsExchangeCloudManaged') -eq $ToCloud) {
                return @{ Ok=$true; Msg='converted (confirmed after error)' }
            }
        } catch { }
        if ([string]::IsNullOrWhiteSpace($errText)) { $errText = 'unknown error (no message returned by Set-Mailbox)' }
        return @{ Ok=$false; Msg=$errText }
    }
}

function Convert-GraphSoa {
    param($Item, [bool]$ToCloud)
    if ($script:DemoMode) {
        Start-Sleep -Milliseconds (20 + (Get-Random -Maximum 50))
        if ((Get-Random -Maximum 100) -lt 4) { return @{ Ok=$false; Msg='demo: simulated transient error' } }
        return @{ Ok=$true; Msg='converted' }
    }
    $resource = 'groups'
    if ($Item.Type -eq 'Contact') { $resource = 'contacts' }
    try {
        $uri = ('https://graph.microsoft.com/v1.0/' + $resource + '/' + $Item.Id + '/onPremisesSyncBehavior')
        $body = @{ isCloudManaged = $ToCloud } | ConvertTo-Json
        [void](Invoke-GraphWorker -Job @{ method='PATCH'; uri=$uri; body=$body; contentType='application/json' })
        return @{ Ok=$true; Msg='converted' }
    } catch {
        return @{ Ok=$false; Msg=$_.Exception.Message }
    }
}

function Get-GroupCloudMembers {
    # Users in the group that are NOT dir-synced (cloud-only) - these block a
    # safe rollback to on-premises per Microsoft guidance.
    param([string]$GroupId)
    if ($script:DemoMode) {
        $r = Get-Random -Maximum 100
        if ($r -lt 25) {
            $n = 1 + (Get-Random -Maximum 4)
            $out = @()
            for ($i = 0; $i -lt $n; $i++) { $out += [pscustomobject]@{ displayName="Cloud User $i"; userPrincipalName="cloud.user$i@contoso.com" } }
            return ,$out
        }
        return ,@()
    }
    $sel = 'id,displayName,userPrincipalName,onPremisesSyncEnabled'
    $members = Invoke-GraphGetAll -Uri ('https://graph.microsoft.com/v1.0/groups/' + $GroupId + '/members/microsoft.graph.user?$select=' + $sel + '&$top=999')
    return ,@($members | Where-Object { -not [bool](Get-PropSafe $_ 'onPremisesSyncEnabled') })
}

function Get-AdMemberTree {
    # Recursively walk an AD group's membership from its SID. Preserves nesting
    # structure (does NOT use Get-ADGroupMember -Recursive, which flattens it).
    param([string]$RootSid)
    $visited = New-Object 'System.Collections.Generic.HashSet[string]'
    $nested  = New-Object System.Collections.ArrayList
    $leaves  = New-Object System.Collections.ArrayList
    $stack   = New-Object System.Collections.Stack

    $root = Get-ADGroup -Identity $RootSid -Properties member -ErrorAction Stop
    if ($root.SID) { [void]$visited.Add([string]$root.SID.Value) }
    foreach ($dn in @($root.member)) { $stack.Push([string]$dn) }

    while ($stack.Count -gt 0) {
        $dn = [string]$stack.Pop()
        $o = Get-ADObject -Identity $dn -Properties objectSid, objectClass, displayName, sAMAccountName -ErrorAction SilentlyContinue
        if ($null -eq $o) { continue }
        $sid = ''
        if ($o.objectSid) { $sid = [string]$o.objectSid.Value }
        if ([string]$o.objectClass -eq 'group') {
            [void]$nested.Add([pscustomobject]@{
                Name  = [string]$o.displayName
                DN    = [string]$o.DistinguishedName
                Sid   = $sid
                Class = 'group'
            })
            if ($sid -and -not $visited.Contains($sid)) {
                [void]$visited.Add($sid)
                $child = Get-ADGroup -Identity $dn -Properties member -ErrorAction SilentlyContinue
                if ($child) { foreach ($m in @($child.member)) { $stack.Push([string]$m) } }
            }
        } else {
            $nm = [string]$o.displayName
            if ([string]::IsNullOrEmpty($nm)) { $nm = [string]$o.sAMAccountName }
            [void]$leaves.Add([pscustomobject]@{
                Name  = $nm
                DN    = [string]$o.DistinguishedName
                Sid   = $sid
                Class = [string]$o.objectClass
            })
        }
    }
    return [pscustomobject]@{ Nested = $nested.ToArray(); Leaves = $leaves.ToArray() }
}

function Invoke-GroupForwardAudit {
    # Forward (to-cloud) audit for one group. Assumes prerequisites already passed
    # (Test-AuditPrerequisite) unless in demo mode.
    param($Group)

    if ($script:DemoMode) {
        Start-Sleep -Milliseconds (10 + (Get-Random -Maximum 30))
        $r = Get-Random -Maximum 100
        $nested  = New-Object System.Collections.ArrayList
        $dropped = New-Object System.Collections.ArrayList
        if ($r -lt 20) {
            [void]$nested.Add([pscustomobject]@{
                Name = ($Group.Name + '-Nested')
                DN   = ('CN=' + $Group.Name + '-Nested,OU=Groups,DC=contoso,DC=com')
                Sid  = ('S-1-5-21-demo-' + (Get-Random -Maximum 9999))
            })
        }
        if ($r -ge 15 -and $r -lt 38) {
            $n = 1 + (Get-Random -Maximum 2)
            for ($i = 0; $i -lt $n; $i++) {
                [void]$dropped.Add([pscustomobject]@{
                    Name  = "Unsynced User $i"
                    Sid   = ('S-1-5-21-demo-' + (Get-Random -Maximum 9999))
                    Class = 'user'
                })
            }
        }
        $state = 'Green'
        if ($nested.Count -gt 0 -or $dropped.Count -gt 0) { $state = 'Yellow' }
        return [pscustomobject]@{
            GroupId = [string]$Group.Id; State = $state
            NestedGroups = $nested.ToArray(); DroppedMembers = $dropped.ToArray(); Error = $null
        }
    }

    $raw = Get-PropSafe $Group 'Raw'
    $sid = [string](Get-PropSafe $raw 'onPremisesSecurityIdentifier')
    # No AD SID => cloud-born object; nothing on-prem to drop.
    if ([string]::IsNullOrEmpty($sid)) {
        return [pscustomobject]@{ GroupId = [string]$Group.Id; State = 'Green'; NestedGroups = @(); DroppedMembers = @(); Error = $null }
    }

    # 1) AD side: recursive membership tree.
    $tree = Get-AdMemberTree -RootSid $sid

    # 2) Entra side: transitive membership SIDs actually present in the cloud.
    $uri = 'https://graph.microsoft.com/v1.0/groups/' + [string]$Group.Id + '/transitiveMembers/microsoft.graph.directoryObject?$select=id,onPremisesSecurityIdentifier'
    $cloudMembers = Invoke-GraphGetAll -Uri $uri
    $cloudSids = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($m in @($cloudMembers)) {
        $s = [string](Get-PropSafe $m 'onPremisesSecurityIdentifier')
        if (-not [string]::IsNullOrEmpty($s)) { [void]$cloudSids.Add($s) }
    }

    # 3) Diff: AD leaf members absent from the cloud copy are dropped after conversion.
    $dropped = New-Object System.Collections.ArrayList
    foreach ($leaf in @($tree.Leaves)) {
        $ls = [string]$leaf.Sid
        if ([string]::IsNullOrEmpty($ls) -or -not $cloudSids.Contains($ls)) {
            [void]$dropped.Add([pscustomobject]@{ Name = [string]$leaf.Name; Sid = $ls; Class = [string]$leaf.Class })
        }
    }

    $state = 'Green'
    if (@($tree.Nested).Count -gt 0 -or $dropped.Count -gt 0) { $state = 'Yellow' }
    return [pscustomobject]@{
        GroupId = [string]$Group.Id; State = $state
        NestedGroups = @($tree.Nested); DroppedMembers = $dropped.ToArray(); Error = $null
    }
}

#endregion

# ============================================================================
#region CSV export / import
# ============================================================================

function Export-ViewCsv {
    param($Tab)
    if (@($Tab.View).Count -eq 0) {
        Show-MsgModal -Title 'Export' -Lines @('Nothing to export - the current view is empty.') -Kind Warn
        return
    }
    Confirm-SoaDirectory -Path $script:ExportDir
    $file = Join-Path $script:ExportDir ("{0}_{1}.csv" -f $Tab.Name, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    @($Tab.View) | ForEach-Object {
        [pscustomobject]@{
            Type        = $_.Type
            DisplayName = $_.Name
            Email       = $_.Email
            Identity    = $_.Id
            SOA         = $_.Soa
            Detail      = $_.Detail
            Selected    = $_.Selected
        }
    } | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Write-SoaLog -Message ("Exported {0} rows to {1}." -f @($Tab.View).Count, $file) -Level OK
    Show-MsgModal -Title 'Export complete' -Lines @(
        ("{0} rows exported to:" -f @($Tab.View).Count),
        @($script:T.CtxHi, $file)
    )
}

function Import-SelectionFile {
    param($Tab)
    $path = Show-InputModal -Title 'Bulk import' -Prompt 'Path to a CSV (Identity/UPN/Email/DisplayName column) or TXT (one identity per line). Matching entries get selected.' -Default ''
    if ($null -eq $path -or [string]::IsNullOrWhiteSpace($path)) { return }
    $path = $path.Trim('"').Trim()
    if (-not (Test-Path -LiteralPath $path)) {
        Show-MsgModal -Title 'Import' -Lines @("File not found:", $path) -Kind Error
        return
    }
    $wanted = New-Object System.Collections.ArrayList
    try {
        if ([IO.Path]::GetExtension($path).ToLower() -eq '.csv') {
            $rows = @(Import-Csv -Path $path)
            if ($rows.Count -gt 0) {
                $cols = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
                $col = $null
                foreach ($cand in @('Identity','UserPrincipalName','PrimarySmtpAddress','EmailAddress','Email','Mail','Id','ObjectId','DisplayName')) {
                    $hit = @($cols | Where-Object { $_ -ieq $cand })
                    if ($hit.Count -gt 0) { $col = $hit[0]; break }
                }
                if ($null -eq $col) { $col = $cols[0] }
                foreach ($r in $rows) {
                    $v = [string](Get-PropSafe $r $col)
                    if (-not [string]::IsNullOrWhiteSpace($v)) { [void]$wanted.Add($v.Trim()) }
                }
            }
        } else {
            foreach ($line in (Get-Content -Path $path)) {
                if (-not [string]::IsNullOrWhiteSpace($line)) { [void]$wanted.Add($line.Trim()) }
            }
        }
    } catch {
        Show-MsgModal -Title 'Import' -Lines @('Failed to read the file:', $_.Exception.Message) -Kind Error
        return
    }
    if ($wanted.Count -eq 0) {
        Show-MsgModal -Title 'Import' -Lines @('No identities found in the file.') -Kind Warn
        return
    }
    # Build lookup over the full item set
    $lookup = @{}
    foreach ($it in @($Tab.Items)) {
        foreach ($key in @($it.Id, $it.Email, $it.Name)) {
            if (-not [string]::IsNullOrEmpty($key)) { $lookup[$key.ToLower()] = $it }
        }
    }
    $matched = 0
    $unmatched = New-Object System.Collections.ArrayList
    foreach ($w in $wanted) {
        $k = $w.ToLower()
        if ($lookup.ContainsKey($k)) {
            if (-not $lookup[$k].Selected) { $lookup[$k].Selected = $true }
            $matched++
        } else {
            [void]$unmatched.Add($w)
        }
    }
    Write-SoaLog -Message ("Bulk import: {0} matched, {1} unmatched from {2}." -f $matched, $unmatched.Count, $path)
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add(@($script:T.Good, ("Matched and selected: $matched")))
    [void]$lines.Add(@($script:T.Warn, ("Unmatched: " + $unmatched.Count)))
    if ($unmatched.Count -gt 0) {
        [void]$lines.Add('')
        foreach ($u in @($unmatched | Select-Object -First 12)) { [void]$lines.Add(@($script:T.Muted, ("  " + $u))) }
        if ($unmatched.Count -gt 12) { [void]$lines.Add(@($script:T.Muted, ("  ...and " + ($unmatched.Count - 12) + " more"))) }
    }
    Show-ReportModal -Title 'Bulk import result' -Lines $lines.ToArray()
}

#endregion

# ============================================================================
#region Conversion pipeline
# ============================================================================

function Get-ConversionTargets {
    param($Tab)
    $sel = @($Tab.Items | Where-Object { $_.Selected })
    if ($sel.Count -gt 0) { return ,$sel }
    $view = @($Tab.View)
    if ($view.Count -gt 0 -and $Tab.Cursor -lt $view.Count) { return ,@($view[$Tab.Cursor]) }
    return ,@()
}

function Invoke-SoaConversion {
    param($Tab, [bool]$ToCloud)
    $targets = Get-ConversionTargets -Tab $Tab
    if ($targets.Count -eq 0) {
        Show-MsgModal -Title 'Convert' -Lines @('Nothing selected.') -Kind Warn
        return
    }

    $dirWord = 'ON-PREMISES managed'
    if ($ToCloud) { $dirWord = 'CLOUD managed' }

    # Skip no-ops
    $skipState = 'OnPrem'
    if ($ToCloud) { $skipState = 'Cloud' }
    $work = @($targets | Where-Object { $_.Soa -ne $skipState })
    $skipped = $targets.Count - $work.Count
    if ($work.Count -eq 0) {
        Show-MsgModal -Title 'Convert' -Lines @("All selected entries are already $dirWord.") -Kind Warn
        return
    }

    $t = $script:T
    $isGroupTab = ($Tab.Noun -eq 'groups')
    $isMailboxTab = ($Tab.Noun -eq 'mailboxes')

    # --- Group rollback safety check -------------------------------------
    $alreadyGated = $false
    if ($isGroupTab -and -not $ToCloud) {
        $findings = New-Object System.Collections.ArrayList
        $idx = 0
        Write-Screen
        foreach ($g in $work) {
            Write-ProgressModal -Title 'Rollback safety check' -Done $idx -Total $work.Count -Label $g.Name -Ok 0 -Failed 0
            try {
                $cloudMembers = Get-GroupCloudMembers -GroupId $g.Id
                if (@($cloudMembers).Count -gt 0) {
                    [void]$findings.Add(@{ Group=$g; Members=@($cloudMembers) })
                }
            } catch {
                Write-SoaLog -Message ("Cloud-member check failed for group '{0}': {1}" -f $g.Name, $_.Exception.Message) -Level WARN
                [void]$findings.Add(@{ Group=$g; Members=@(); CheckFailed=$_.Exception.Message })
            }
            $idx++
        }
        $script:UI.Dirty = $true
        if ($findings.Count -gt 0) {
            $lines = New-Object System.Collections.ArrayList
            [void]$lines.Add(@($t.Danger, 'Rollback hazard detected.'))
            [void]$lines.Add('')
            [void]$lines.Add('Microsoft guidance: remove cloud-only members and access-package references BEFORE rolling a group back to AD, otherwise those references are lost when sync takes over.')
            [void]$lines.Add('')
            foreach ($f in $findings) {
                $g = $f['Group']
                if ($f.ContainsKey('CheckFailed')) {
                    [void]$lines.Add(@($t.Warn, ("  " + $g.Name + " - member check FAILED: " + $f['CheckFailed'])))
                } else {
                    $m = $f['Members']
                    [void]$lines.Add(@($t.Warn, ("  " + $g.Name + " - " + $m.Count + " cloud-only member(s):")))
                    foreach ($mm in @($m | Select-Object -First 5)) {
                        $upn = [string](Get-PropSafe $mm 'userPrincipalName')
                        [void]$lines.Add(@($t.Muted, ("      " + (Get-PropSafe $mm 'displayName') + "  <" + $upn + ">")))
                    }
                    if ($m.Count -gt 5) { [void]$lines.Add(@($t.Muted, ("      ...and " + ($m.Count - 5) + " more"))) }
                }
            }
            Show-ReportModal -Title 'Group rollback check' -Lines $lines.ToArray()
            $proceed = Show-TypedConfirmModal -Title 'Proceed despite hazards?' -Lines @(
                ("{0} of {1} group(s) have cloud-only members or failed checks." -f $findings.Count, $work.Count),
                'Rolling back now can permanently drop those member references.'
            ) -Word 'ROLLBACK'
            if (-not $proceed) { Write-SoaLog -Message 'Group rollback cancelled after safety check.'; return }
            $alreadyGated = $true
        }
    }

    # --- Confirmation ------------------------------------------------------
    $nameList = New-Object System.Collections.ArrayList
    foreach ($w in @($work | Select-Object -First 12)) { [void]$nameList.Add(@($t.CtxHi, ('  ' + $w.Name))) }
    if ($work.Count -gt 12) { [void]$nameList.Add(@($t.Muted, ('  ...and ' + ($work.Count - 12) + ' more'))) }

    $note = ''
    if ($isMailboxTab -and $ToCloud) { $note = 'Exchange attributes become editable in EXO; on-prem values stop syncing for these mailboxes. An attribute backup (JSON) is written first.' }
    elseif ($isMailboxTab) { $note = 'The next sync cycle overwrites cloud-edited Exchange attributes with on-premises values. An attribute backup (JSON) is written first.' }
    elseif ($ToCloud) { $note = 'The object(s) become cloud-managed in Microsoft Entra; changes made in AD will no longer sync. A JSON backup of the directory object is written first.' }
    else { $note = 'SOA returns to AD once the sync client next runs (status shows Pending until then). A JSON backup of the directory object is written first.' }

    $confirmLines = New-Object System.Collections.ArrayList
    $nounWord = [string]$Tab['Noun']
    if ($work.Count -eq 1) {
        if ($nounWord -eq 'mailboxes') { $nounWord = 'mailbox' } else { $nounWord = $nounWord.TrimEnd('s') }
    }
    [void]$confirmLines.Add(("Convert {0} {1} to {2}?" -f $work.Count, $nounWord, $dirWord))
    if ($skipped -gt 0) { [void]$confirmLines.Add(@($t.Muted, ("({0} already in target state - skipped)" -f $skipped))) }
    [void]$confirmLines.Add('')
    foreach ($nl in $nameList) { [void]$confirmLines.Add($nl) }
    [void]$confirmLines.Add('')
    [void]$confirmLines.Add(@($t.Muted, $note))

    $danger = (-not $ToCloud)
    $confirmed = $alreadyGated
    if (-not $confirmed) {
        if ($danger) {
            $confirmed = Show-ConfirmModal -Title 'Confirm conversion' -Lines $confirmLines.ToArray() -Danger
        } else {
            $confirmed = Show-ConfirmModal -Title 'Confirm conversion' -Lines $confirmLines.ToArray()
        }
    }
    if (-not $confirmed) { Write-SoaLog -Message 'Conversion cancelled by operator.'; return }

    Write-SoaLog -Message ("Starting conversion of {0} {1} to {2}..." -f $work.Count, $Tab.Noun, $dirWord)

    # --- Backup ------------------------------------------------------------
    $backupFile = $null
    if (-not $isMailboxTab) {
        # Graph objects: we already hold the directory object snapshots.
        $records = @($work | ForEach-Object {
            [ordered]@{ Item=$_.Raw; Soa=$_.Soa; Type=$_.Type; CapturedAt=(Get-Date -Format o) }
        })
        try { $backupFile = Save-BackupFile -Records $records -Kind $Tab.Name } catch {
            Write-SoaLog -Message ("Backup failed: {0}" -f $_.Exception.Message) -Level ERROR
            if (-not (Show-ConfirmModal -Title 'Backup failed' -Lines @($_.Exception.Message, '', 'Continue WITHOUT a backup?') -Danger)) { return }
        }
    }

    # --- Convert loop -------------------------------------------------------
    $okCount = 0; $failCount = 0
    $results = New-Object System.Collections.ArrayList
    $mailboxBackups = New-Object System.Collections.ArrayList
    $done = 0
    Write-Screen
    foreach ($item in $work) {
        Write-ProgressModal -Title ("Converting to " + $dirWord) -Done $done -Total $work.Count -Label $item.Name -Ok $okCount -Failed $failCount
        $res = $null
        if ($isMailboxTab) {
            $backupOk = $true
            try {
                $rec = Get-MailboxBackupRecord -Item $item
                [void]$mailboxBackups.Add($rec)
            } catch {
                $backupOk = $false
                $res = @{ Ok=$false; Msg=("backup failed, conversion skipped: " + $_.Exception.Message) }
            }
            if ($backupOk) { $res = Convert-MailboxSoa -Item $item -ToCloud $ToCloud }
        } else {
            $res = Convert-GraphSoa -Item $item -ToCloud $ToCloud
        }
        if ($res['Ok']) {
            $okCount++
            if ($ToCloud) { $item.Soa = 'Cloud' }
            elseif ($isMailboxTab) { $item.Soa = 'OnPrem' }
            else { $item.Soa = 'Pending' }
            $item.Selected = $false
            [void]$results.Add(@($t.Good, ("  OK    " + $item.Name)))
            Write-SoaLog -Message ("Converted '{0}' ({1}) to {2}." -f $item.Name, $item.Id, $dirWord) -Level OK
        } else {
            $failCount++
            [void]$results.Add(@($t.Danger, ("  FAIL  " + $item.Name + " - " + $res['Msg'])))
            Write-SoaLog -Message ("FAILED converting '{0}' ({1}): {2}" -f $item.Name, $item.Id, $res['Msg']) -Level ERROR
        }
        $done++
        Write-ProgressModal -Title ("Converting to " + $dirWord) -Done $done -Total $work.Count -Label $item.Name -Ok $okCount -Failed $failCount
    }

    if ($isMailboxTab -and $mailboxBackups.Count -gt 0) {
        try { $backupFile = Save-BackupFile -Records $mailboxBackups.ToArray() -Kind 'Mailboxes' } catch {
            Write-SoaLog -Message ("Writing mailbox backup file failed: {0}" -f $_.Exception.Message) -Level ERROR
        }
    }

    Write-SoaLog -Message ("Conversion finished. OK: {0}  Failed: {1}" -f $okCount, $failCount) -Level OK

    # --- Summary ------------------------------------------------------------
    $summary = New-Object System.Collections.ArrayList
    [void]$summary.Add(@($t.Good,  ("Successful : " + $okCount)))
    [void]$summary.Add(@($t.Danger,("Failed     : " + $failCount)))
    if ($skipped -gt 0) { [void]$summary.Add(@($t.Muted, ("Skipped    : " + $skipped + " (already in target state)"))) }
    if ($backupFile) { [void]$summary.Add(@($t.Muted, ("Backup     : " + $backupFile))) }
    if (-not $ToCloud -and -not $isMailboxTab) {
        [void]$summary.Add('')
        [void]$summary.Add(@($t.Warn, 'Rollback completes after the next Connect/Cloud Sync cycle (status: Pending).'))
    }
    [void]$summary.Add('')
    foreach ($r in $results) { [void]$summary.Add($r) }
    Show-ReportModal -Title 'Conversion summary' -Lines $summary.ToArray()
    Update-TabView -Tab $Tab
}

#endregion

# ============================================================================
#region Views
# ============================================================================

function Update-TabView {
    param($Tab)
    $items = @($Tab['Items'])
    if ($Tab['Filter'] -ne 'All') {
        $items = @($items | Where-Object { $_.Soa -eq $Tab['Filter'] })
    }
    if (-not [string]::IsNullOrEmpty($Tab['Search'])) {
        $needle = $Tab['Search']
        $items = @($items | Where-Object { ($_.Name -like "*$needle*") -or ($_.Email -like "*$needle*") })
    }
    $sortProp = $Tab['SortCol']
    if ($sortProp -eq 'Soa') { $sortProp = 'Soa' }
    if ($Tab['SortDesc']) {
        $items = @($items | Sort-Object -Property $sortProp -Descending)
    } else {
        $items = @($items | Sort-Object -Property $sortProp)
    }
    $Tab['View'] = $items
    if ($Tab['Cursor'] -ge $items.Count) { $Tab['Cursor'] = [Math]::Max(0, $items.Count - 1) }
    if ($Tab['Cursor'] -lt 0) { $Tab['Cursor'] = 0 }
    $script:UI.Dirty = $true
}

function Invoke-TabLoad {
    param($Tab, [switch]$Force)
    if ($Tab['Kind'] -eq 'Log') { return }
    if ($Tab['Kind'] -eq 'Org') {
        if (-not (Connect-ExoService)) { return }
        [void](Get-OrgState)
        $script:UI.Dirty = $true
        return
    }
    if ($Tab['Loaded'] -and -not $Force) { return }
    if ($Tab['Conn'] -eq 'Exo') {
        if (-not (Connect-ExoService)) { return }
    } else {
        if (-not (Connect-GraphService)) { return }
    }
    Write-Screen
    $title = 'Loading ' + $Tab['Name']
    # Start the spinner first so this initial paint already publishes its
    # coordinates - the whole point is animating before the first results.
    Start-LoadSpinner
    $loadLabel = 'Contacting service - waiting for first results...'
    if ($Tab['Noun'] -eq 'mailboxes') {
        # EXO buffers the whole result set before returning anything, so this
        # label is on screen for the entire fetch - warn that it can be slow.
        $loadLabel = 'Fetching mailboxes - this can take several minutes...'
    }
    Write-ProgressModal -Title $title -Done 0 -Total 0 -Label $loadLabel -Ok 0 -Failed 0
    $renderState = @{ LastTick = 0 }
    # GetNewClosure() binds the scriptblock to a dynamic module whose command
    # lookup skips this script's scope, so script-level functions (e.g.
    # Write-ProgressModal) are not resolvable by name inside the closure.
    # Capture function references as variables and invoke those instead.
    $fnProgressModal = ${function:Write-ProgressModal}
    $progressCb = {
        param($Count, $Label)
        # Esc aborts the load; other keys typed during loading are discarded
        # so they do not fire as hotkeys once the load completes.
        while ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::Escape) {
                throw (New-Object System.OperationCanceledException 'Load cancelled by user.')
            }
        }
        $now = [Environment]::TickCount
        if (($now - $renderState.LastTick) -lt 150) { return }
        $renderState.LastTick = $now
        & $fnProgressModal -Title $title -Done $Count -Total 0 -Label $Label -Ok 0 -Failed 0
    }.GetNewClosure()
    try {
        $items = @()
        try {
            switch ($Tab['Noun']) {
                'mailboxes' { $items = Get-MailboxItems -Progress $progressCb }
                'groups'    { $items = Get-GroupItems -Progress $progressCb }
                'contacts'  { $items = Get-ContactItems -Progress $progressCb }
            }
        } finally {
            # Join the spinner runspace BEFORE the catch handlers below can
            # draw modals, so no stray spinner frame lands on top of them.
            Stop-LoadSpinner
        }
        $Tab['Items'] = @($items)
        $Tab['Loaded'] = $true
        $Tab['Cursor'] = 0
        $Tab['Scroll'] = 0
        Update-TabView -Tab $Tab
    } catch [System.OperationCanceledException] {
        Write-SoaLog -Message ("Loading {0} cancelled by user (Esc). Press R or Enter to reload." -f $Tab['Name']) -Level WARN
    } catch {
        Write-SoaLog -Message ("Failed to load {0}: {1}" -f $Tab['Name'], $_.Exception.Message) -Level ERROR
        Show-MsgModal -Title 'Load failed' -Lines @(("Failed to load " + $Tab['Name'] + ":"), $_.Exception.Message) -Kind Error
    }
    $script:UI.Dirty = $true
}

function Add-TitleBar {
    param([System.Text.StringBuilder]$Sb, [int]$W)
    $t = $script:T; $g = $script:G
    $left = " Exchange SOA Manager  v$script:Version"
    $pieces = New-Object System.Collections.ArrayList
    $plainRight = 0

    $exoGlyph = [string]$g.Ring; $exoStyle = $t.TitleOff; $exoText = 'EXO'
    if ($script:Conn.Exo) { $exoGlyph = [string]$g.Dot; $exoStyle = $t.TitleOk; $exoText = 'EXO ' + $script:Conn.ExoAccount }
    [void]$pieces.Add(@($exoStyle, ($exoGlyph + ' ' + $exoText)))

    $gGlyph = [string]$g.Ring; $gStyle = $t.TitleOff; $gText = 'Graph'
    if ($script:Conn.Graph) { $gGlyph = [string]$g.Dot; $gStyle = $t.TitleOk; $gText = 'Graph ' + $script:Conn.GraphAccount }
    [void]$pieces.Add(@($gStyle, ($gGlyph + ' ' + $gText)))

    if ($script:DemoMode) { [void]$pieces.Add(@($t.TitleDemo, 'DEMO')) }

    $sep = '   '
    foreach ($p in $pieces) { $plainRight += ([string]$p[1]).Length }
    $plainRight += $sep.Length * ($pieces.Count - 1) + 1   # trailing space

    $mid = $W - $left.Length - $plainRight
    if ($mid -lt 1) { $mid = 1 }
    $line = $t.TitleApp + $left + $t.TitleBg + (' ' * $mid)
    for ($i = 0; $i -lt $pieces.Count; $i++) {
        if ($i -gt 0) { $line += $t.TitleDim + $sep }
        $line += ([string]$pieces[$i][0]) + ([string]$pieces[$i][1])
    }
    $line += $t.TitleBg + ' '
    Add-FrameLine -Sb $Sb -Row 1 -Content $line
}

function Add-TabBar {
    param([System.Text.StringBuilder]$Sb, [int]$W)
    $t = $script:T
    $line = $t.TabBg + ' '
    $plain = 1
    for ($i = 0; $i -lt $script:Tabs.Count; $i++) {
        $tab = $script:Tabs[$i]
        $label = ' ' + ($i + 1) + ' ' + $tab['Name'] + ' '
        if ($i -eq $script:UI.Tab) { $line += $t.TabOn + $label + $t.TabBg }
        else { $line += $t.TabOff + $label + $t.TabBg }
        $line += ' '
        $plain += $label.Length + 1
    }
    if ($plain -lt $W) { $line += (' ' * ($W - $plain)) }
    Add-FrameLine -Sb $Sb -Row 2 -Content $line
}

function Get-ListLayout {
    param([int]$W)
    # Columns: ' ' chk(3) ' ' name email '  ' detail(13) '  ' soa(10)
    $fixed = 1 + 3 + 1 + 2 + 13 + 2 + 10
    $flex = $W - $fixed - 1
    if ($flex -lt 20) { $flex = 20 }
    $nameW = [int]($flex * 0.42)
    $emailW = $flex - $nameW - 2
    return @{ Name=$nameW; Email=$emailW; Detail=13; Soa=10 }
}

function Add-ListView {
    param([System.Text.StringBuilder]$Sb, $Tab, [int]$W, [int]$H)
    $t = $script:T; $g = $script:G

    if (-not $Tab['Loaded']) {
        Add-FrameLine -Sb $Sb -Row 3 -Content ($t.Ctx + ' not loaded')
        for ($r = 4; $r -le ($H - 1); $r++) { Add-FrameLine -Sb $Sb -Row $r -Content '' }
        $connName = 'Exchange Online'
        if ($Tab['Conn'] -eq 'Graph') { $connName = 'Microsoft Graph' }
        $isConnected = $script:Conn[$Tab['Conn']]
        $lines = New-Object System.Collections.ArrayList
        if ($isConnected) {
            $head = $Tab['Name'] + ' are not loaded yet.'
        } else {
            $head = 'Not connected to ' + $connName + '.'
        }
        [void]$lines.Add(@(($t.CtxHi + $head), $head.Length))
        [void]$lines.Add(@('', 0))
        if ($Tab['Noun'] -eq 'mailboxes' -or $Tab['Noun'] -eq 'groups' -or $Tab['Noun'] -eq 'contacts') {
            $scopeNote = "Note: Only directory-synced (hybrid) $($Tab['Noun']) will be loaded."
            [void]$lines.Add(@(($t.Muted + $scopeNote), $scopeNote.Length))
            [void]$lines.Add(@('', 0))
        }
        $hint = 'Press Enter to connect and load.'
        if ($isConnected) { $hint = 'Press Enter to load.' }
        [void]$lines.Add(@(($t.Row + $hint), $hint.Length))
        if (-not $isConnected) {
            [void]$lines.Add(@('', 0))
            if ($Tab['Conn'] -eq 'Graph') {
                $role = 'Required role: Hybrid Identity Administrator'
                [void]$lines.Add(@(($t.Warn + $role), $role.Length))
                $note = 'Graph scopes (incl. *-OnPremisesSyncBehavior) need one-time admin consent.'
                [void]$lines.Add(@(($t.Muted + $note), $note.Length))
            } else {
                $role = 'Required role: Exchange Administrator (or Hybrid Identity / Global Admin)'
                [void]$lines.Add(@(($t.Warn + $role), $role.Length))
            }
            [void]$lines.Add(@('', 0))
            $pim1 = 'Using PIM? Activate the role BEFORE signing in.'
            [void]$lines.Add(@(($t.Muted + $pim1), $pim1.Length))
            $pim2 = 'Activated it after connecting? Press W to disconnect, then reconnect.'
            [void]$lines.Add(@(($t.Muted + $pim2), $pim2.Length))
        }
        [void](Write-CenteredPanel -Sb $Sb -Lines $lines.ToArray() -Top 6 -Bottom ($H - 4) -Width $W)
        return
    }

    $view = @($Tab['View'])
    $selCount = @($Tab['Items'] | Where-Object { $_.Selected }).Count
    $dir = [string]$g.Up
    if ($Tab['SortDesc']) { $dir = [string]$g.Down }
    $ctx = (' {0} of {1} {2}   {3} selected   filter:{4}   sort:{5}{6}' -f @($view).Count, @($Tab['Items']).Count, $Tab['Noun'], $selCount, $Tab['Filter'], $Tab['SortCol'], $dir)
    if (-not [string]::IsNullOrEmpty($Tab['Search'])) { $ctx += ('   search:"' + $Tab['Search'] + '"') }
    Add-FrameLine -Sb $Sb -Row 3 -Content ($t.Ctx + $ctx)

    $col = Get-ListLayout -W $W
    $head = ' ' + (Get-PadCell 'sel' 3) + ' ' + (Get-PadCell 'Name' $col.Name) + '  ' + (Get-PadCell 'Email' $col.Email) + '  ' + (Get-PadCell 'Type' $col.Detail) + '  ' + (Get-PadCell 'SOA' $col.Soa)
    Add-FrameLine -Sb $Sb -Row 4 -Content ($t.ColHead + $head)

    $top = 5; $bottom = $H - 1
    $cap = $bottom - $top + 1
    if ($cap -lt 1) { $cap = 1 }

    # clamp scroll around cursor
    if ($Tab['Cursor'] -lt $Tab['Scroll']) { $Tab['Scroll'] = $Tab['Cursor'] }
    if ($Tab['Cursor'] -ge ($Tab['Scroll'] + $cap)) { $Tab['Scroll'] = $Tab['Cursor'] - $cap + 1 }
    $maxScroll = [Math]::Max(0, $view.Count - $cap)
    if ($Tab['Scroll'] -gt $maxScroll) { $Tab['Scroll'] = $maxScroll }
    if ($Tab['Scroll'] -lt 0) { $Tab['Scroll'] = 0 }

    for ($i = 0; $i -lt $cap; $i++) {
        $row = $top + $i
        $idx = $Tab['Scroll'] + $i
        if ($idx -ge $view.Count) { Add-FrameLine -Sb $Sb -Row $row -Content ''; continue }
        $item = $view[$idx]
        $isCursor = ($idx -eq $Tab['Cursor'])

        $chk = [string]$g.ChkOff
        if ($item.Selected) { $chk = [string]$g.ChkOn }

        $line = ''
        if ($isCursor) { $line += $t.CursorBg + $t.CursorFg }
        else { $line += $t.Row }

        if ($item.Selected) {
            if ($isCursor) { $line += ' ' + $chk + ' ' }
            else { $line += $t.SelMark + ' ' + $chk + ' ' + $t.Row }
        } else {
            $line += ' ' + $chk + ' '
        }
        $line += (Get-PadCell $item.Name $col.Name) + '  '
        if ($isCursor) { $line += (Get-PadCell $item.Email $col.Email) }
        else { $line += $t.RowDim + (Get-PadCell $item.Email $col.Email) + $t.Row }
        $line += '  ' + (Get-PadCell $item.Detail $col.Detail) + '  '
        $line += (Get-SoaBadge -Soa $item.Soa -Width $col.Soa)
        Add-FrameLine -Sb $Sb -Row $row -Content $line
    }
}

function Add-OrgView {
    param([System.Text.StringBuilder]$Sb, [int]$W, [int]$H)
    $t = $script:T; $g = $script:G
    Add-FrameLine -Sb $Sb -Row 3 -Content ($t.Ctx + ' Tenant-wide default SOA for new dir-synced mailboxes')
    for ($r = 4; $r -le ($H - 1); $r++) { Add-FrameLine -Sb $Sb -Row $r -Content '' }

    if (-not $script:Org.Loaded) {
        $head = 'Organization configuration not loaded.'
        $hint = 'Press Enter to connect to Exchange Online and read the current state.'
        $role = 'Required role: Hybrid Identity Administrator or Global Administrator'
        $pim1 = 'Using PIM? Activate the role BEFORE signing in.'
        $pim2 = 'Activated it after connecting? Press W to disconnect, then reconnect.'
        $lines = @(
            @(($t.CtxHi + $head), $head.Length),
            @('', 0),
            @(($t.Row + $hint), $hint.Length),
            @('', 0),
            @(($t.Warn + $role), $role.Length),
            @('', 0),
            @(($t.Muted + $pim1), $pim1.Length),
            @(($t.Muted + $pim2), $pim2.Length)
        )
        [void](Write-CenteredPanel -Sb $Sb -Lines $lines -Top 6 -Bottom ($H - 4) -Width $W)
        return
    }

    $margin = 4
    $row = 5
    $pad = ' ' * $margin

    Add-FrameLine -Sb $Sb -Row $row -Content ($pad + $t.Muted + 'Tenant   : ' + $t.CtxHi + $script:Org.TenantName); $row++
    $checked = ''
    if ($script:Org.CheckedAt) { $checked = $script:Org.CheckedAt.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) }
    Add-FrameLine -Sb $Sb -Row $row -Content ($pad + $t.Muted + 'Checked  : ' + $checked); $row++
    $row++

    if ($script:Org.CloudDefault) {
        Add-FrameLine -Sb $Sb -Row $row -Content ($pad + $t.Cloud + [string]$g.Dot + ' NEW dir-synced mailboxes default to CLOUD-managed Exchange attributes'); $row++
        Add-FrameLine -Sb $Sb -Row $row -Content ($pad + $t.Muted + '  BlockExchangeProvisioningFromOnPremEnabled = True'); $row++
    } else {
        Add-FrameLine -Sb $Sb -Row $row -Content ($pad + $t.OnPrem + [string]$g.Dot + ' NEW dir-synced mailboxes default to SERVER-managed (classic hybrid)'); $row++
        Add-FrameLine -Sb $Sb -Row $row -Content ($pad + $t.Muted + '  BlockExchangeProvisioningFromOnPremEnabled = False'); $row++
    }
    $row++

    $info = @(
        'When the cloud default is ENABLED, user accounts associated with newly created',
        'mailboxes sync to Microsoft Entra ID without Exchange attributes from AD; after',
        'an Exchange Online license is assigned, the mailbox is cloud-managed by default.',
        'Existing mailboxes are NOT affected - convert them on the Mailboxes tab.'
    )
    foreach ($ln in $info) {
        if ($row -gt ($H - 3)) { break }
        Add-FrameLine -Sb $Sb -Row $row -Content ($pad + $t.Row + $ln); $row++
    }
    $row++
    $warn = @(
        'Microsoft does not recommend enabling this while you still host and manage',
        'mailboxes on an on-premises Exchange Server: newly created on-prem mailboxes',
        'will not appear in the Exchange Online GAL while the feature is on.'
    )
    foreach ($ln in $warn) {
        if ($row -gt ($H - 2)) { break }
        Add-FrameLine -Sb $Sb -Row $row -Content ($pad + $t.Warn + $ln); $row++
    }
    $row++
    if ($row -le ($H - 1)) {
        $cmdLine = 'Set-OrganizationConfig -ExchangeAttributes{Cloud|Server}ManagedByDefault'
        Add-FrameLine -Sb $Sb -Row $row -Content ($pad + $t.Muted + 'Commands: ' + (Get-PadCell $cmdLine ($W - $margin - 11)))
    }
}

function Add-LogView {
    param([System.Text.StringBuilder]$Sb, [int]$W, [int]$H)
    $t = $script:T
    Add-FrameLine -Sb $Sb -Row 3 -Content ($t.Ctx + ' ' + $script:LogFile)
    $top = 4; $bottom = $H - 1
    $cap = $bottom - $top + 1
    $total = $script:LogBuffer.Count
    $maxScroll = [Math]::Max(0, $total - $cap)
    if ($script:UI.LogScroll -gt $maxScroll) { $script:UI.LogScroll = $maxScroll }
    $start = [Math]::Max(0, $total - $cap - $script:UI.LogScroll)
    for ($i = 0; $i -lt $cap; $i++) {
        $row = $top + $i
        $idx = $start + $i
        if ($idx -ge ($total - $script:UI.LogScroll)) { Add-FrameLine -Sb $Sb -Row $row -Content ''; continue }
        if ($idx -lt 0 -or $idx -ge $total) { Add-FrameLine -Sb $Sb -Row $row -Content ''; continue }
        $entry = $script:LogBuffer[$idx]
        $style = $t.Row
        switch ($entry['Level']) {
            'WARN'  { $style = $t.Warn }
            'ERROR' { $style = $t.Danger }
            'OK'    { $style = $t.Good }
        }
        $text = ' ' + $entry['Stamp'].Substring(11) + '  ' + (Get-PadCell $entry['Level'] 5) + ' ' + $entry['Message']
        Add-FrameLine -Sb $Sb -Row $row -Content ($style + (Get-PadCell $text ($W - 1)))
    }
}

function Get-TabHints {
    param($Tab)
    if ($script:UI.SearchMode) {
        return @() # footer is replaced by the search field
    }
    switch ($Tab['Kind']) {
        'List' {
            return @(
                @('Spc','select'), @('A','all'), @('N','none'),
                @('/','find'), @('F','filter'), @('S','sort'),
                @('C','to cloud'), @('O','to on-prem'),
                @('E','export'), @('I','import'), @('R','reload'),
                @('?','help'), @('Q','quit')
            )
        }
        'Org' {
            return @(
                @('E','enable cloud default'), @('D','server default'),
                @('R','refresh'), @('?','help'), @('Q','quit')
            )
        }
        'Log' {
            return @(
                @('Up/Dn','scroll'), @('O','open log file'),
                @('?','help'), @('Q','quit')
            )
        }
    }
    return @()
}

function Write-Screen {
    $size = Get-ConsoleSize
    $W = $size[0]; $H = $size[1]
    $script:UI.W = $W; $script:UI.H = $H
    $sb = New-Object System.Text.StringBuilder

    if ($W -lt 80 -or $H -lt 20) {
        [void]$sb.Append("$script:ESC[2J$script:ESC[H")
        [void]$sb.Append($script:T.Warn + "Terminal too small ($W x $H). Please resize to at least 80x20." + $script:T.Reset)
        [Console]::Write($sb.ToString())
        return
    }

    Add-TitleBar -Sb $sb -W $W
    Add-TabBar -Sb $sb -W $W

    $tab = $script:Tabs[$script:UI.Tab]
    switch ($tab['Kind']) {
        'List' { Add-ListView -Sb $sb -Tab $tab -W $W -H $H }
        'Org'  { Add-OrgView -Sb $sb -W $W -H $H }
        'Log'  { Add-LogView -Sb $sb -W $W -H $H }
    }

    # footer
    if ($script:UI.SearchMode -and $tab['Kind'] -eq 'List') {
        $t = $script:T
        $search = ' /' + $tab['Search'] + '_'
        $hint = '   Enter keep   Esc clear'
        $padLen = $script:UI.W - $search.Length - $hint.Length
        if ($padLen -lt 0) { $padLen = 0 }
        $footer = $t.FootBg + $t.FootKey + $search + $t.FootTxt + $hint + (' ' * $padLen)
        Add-FrameLine -Sb $sb -Row $H -Content $footer
    } else {
        Add-FrameLine -Sb $sb -Row $H -Content (Get-FooterBar -Hints (Get-TabHints -Tab $tab) -Width $W)
    }

    [Console]::Write($sb.ToString())
}

#endregion

# ============================================================================
#region Org actions
# ============================================================================

function Invoke-OrgEnable {
    if (-not $script:Org.Loaded) { Invoke-TabLoad -Tab $script:Tabs[3]; if (-not $script:Org.Loaded) { return } }
    if ($script:Org.CloudDefault) {
        Show-MsgModal -Title 'Organization' -Lines @('The cloud-managed default is already enabled.')
        return
    }
    $ok = Show-TypedConfirmModal -Title 'Enable tenant-wide cloud default' -Lines @(
        'This runs:',
        @($script:T.CtxHi, '  Set-OrganizationConfig -ExchangeAttributesCloudManagedByDefault'),
        '',
        'New dir-synced users will sync WITHOUT Exchange attributes from AD and their mailboxes will be cloud-managed once licensed.',
        '',
        @($script:T.Warn, 'Warning: while enabled, newly created on-premises mailboxes do not appear in the Exchange Online GAL. Not recommended while you still actively provision mailboxes on-premises.')
    ) -Word 'ENABLE'
    if (-not $ok) { Write-SoaLog -Message 'Org enable cancelled.'; return }
    if (Set-OrgState -CloudDefault $true) {
        [void](Get-OrgState)
        Show-MsgModal -Title 'Organization' -Lines @('Cloud-managed default is now ENABLED.', '', 'Verify with: Get-OrganizationConfig | FL BlockExchangeProvisioningFromOnPremEnabled')
    }
}

function Invoke-OrgDisable {
    if (-not $script:Org.Loaded) { Invoke-TabLoad -Tab $script:Tabs[3]; if (-not $script:Org.Loaded) { return } }
    if (-not $script:Org.CloudDefault) {
        Show-MsgModal -Title 'Organization' -Lines @('The server-managed default is already active.')
        return
    }
    $ok = Show-TypedConfirmModal -Title 'Revert to server-managed default' -Lines @(
        'This runs:',
        @($script:T.CtxHi, '  Set-OrganizationConfig -ExchangeAttributesServerManagedByDefault'),
        '',
        'New dir-synced mailboxes will again default to on-premises (server) management. Users created while the feature was enabled are not affected.'
    ) -Word 'DISABLE'
    if (-not $ok) { Write-SoaLog -Message 'Org disable cancelled.'; return }
    if (Set-OrgState -CloudDefault $false) {
        [void](Get-OrgState)
        Show-MsgModal -Title 'Organization' -Lines @('Server-managed default restored.')
    }
}

#endregion

# ============================================================================
#region Key dispatch
# ============================================================================

function Invoke-ListKey {
    param($Tab, [System.ConsoleKeyInfo]$K)

    # search mode captures input first
    if ($script:UI.SearchMode) {
        if ($K.Key -eq 'Escape') { $script:UI.SearchMode = $false; $Tab['Search'] = ''; Update-TabView -Tab $Tab; return }
        if ($K.Key -eq 'Enter')  { $script:UI.SearchMode = $false; return }
        if ($K.Key -eq 'Backspace') {
            if ($Tab['Search'].Length -gt 0) { $Tab['Search'] = $Tab['Search'].Substring(0, $Tab['Search'].Length - 1); Update-TabView -Tab $Tab }
            return
        }
        if ($K.KeyChar -and -not [char]::IsControl($K.KeyChar)) {
            $Tab['Search'] = $Tab['Search'] + $K.KeyChar
            $Tab['Cursor'] = 0
            Update-TabView -Tab $Tab
        }
        return
    }

    if (-not $Tab['Loaded']) {
        if ($K.Key -eq 'Enter' -or [char]::ToUpper($K.KeyChar) -eq 'R') { Invoke-TabLoad -Tab $Tab -Force }
        return
    }

    $view = @($Tab['View'])
    $cap = [Math]::Max(1, $script:UI.H - 5)

    switch ($K.Key) {
        'UpArrow'   { if ($Tab['Cursor'] -gt 0) { $Tab['Cursor']-- }; return }
        'DownArrow' { if ($Tab['Cursor'] -lt ($view.Count - 1)) { $Tab['Cursor']++ }; return }
        'PageUp'    { $Tab['Cursor'] = [Math]::Max(0, $Tab['Cursor'] - $cap); return }
        'PageDown'  { $Tab['Cursor'] = [Math]::Min([Math]::Max(0, $view.Count - 1), $Tab['Cursor'] + $cap); return }
        'Home'      { $Tab['Cursor'] = 0; return }
        'End'       { $Tab['Cursor'] = [Math]::Max(0, $view.Count - 1); return }
        'Spacebar'  {
            if ($view.Count -gt 0 -and $Tab['Cursor'] -lt $view.Count) {
                $item = $view[$Tab['Cursor']]
                $item.Selected = -not $item.Selected
                if ($Tab['Cursor'] -lt ($view.Count - 1)) { $Tab['Cursor']++ }
            }
            return
        }
        'Enter'     { return }
    }

    switch ([char]::ToUpper($K.KeyChar)) {
        'A' { foreach ($it in $view) { $it.Selected = $true }; return }
        'N' { foreach ($it in @($Tab['Items'])) { $it.Selected = $false }; return }
        '/' { $script:UI.SearchMode = $true; return }
        'F' {
            $order = @('All','Cloud','OnPrem','Pending')
            $idx = [Array]::IndexOf($order, $Tab['Filter'])
            $Tab['Filter'] = $order[(($idx + 1) % $order.Count)]
            $Tab['Cursor'] = 0
            Update-TabView -Tab $Tab
            return
        }
        'S' {
            $order = @('Name','Email','Soa')
            $idx = [Array]::IndexOf($order, $Tab['SortCol'])
            $Tab['SortCol'] = $order[(($idx + 1) % $order.Count)]
            Update-TabView -Tab $Tab
            return
        }
        'D' { $Tab['SortDesc'] = -not $Tab['SortDesc']; Update-TabView -Tab $Tab; return }
        'R' { Invoke-TabLoad -Tab $Tab -Force; return }
        'E' { Export-ViewCsv -Tab $Tab; return }
        'I' { Import-SelectionFile -Tab $Tab; Update-TabView -Tab $Tab; return }
        'C' { Invoke-SoaConversion -Tab $Tab -ToCloud $true; return }
        'O' { Invoke-SoaConversion -Tab $Tab -ToCloud $false; return }
    }
}

function Invoke-OrgKey {
    param([System.ConsoleKeyInfo]$K)
    if ($K.Key -eq 'Enter' -and -not $script:Org.Loaded) { Invoke-TabLoad -Tab $script:Tabs[3]; return }
    switch ([char]::ToUpper($K.KeyChar)) {
        'E' { Invoke-OrgEnable; return }
        'D' { Invoke-OrgDisable; return }
        'R' { Invoke-TabLoad -Tab $script:Tabs[3]; return }
    }
}

function Invoke-LogKey {
    param([System.ConsoleKeyInfo]$K)
    $cap = [Math]::Max(1, $script:UI.H - 4)
    $maxScroll = [Math]::Max(0, $script:LogBuffer.Count - $cap)
    switch ($K.Key) {
        'UpArrow'   { $script:UI.LogScroll = [Math]::Min($maxScroll, $script:UI.LogScroll + 1); return }
        'DownArrow' { $script:UI.LogScroll = [Math]::Max(0, $script:UI.LogScroll - 1); return }
        'PageUp'    { $script:UI.LogScroll = [Math]::Min($maxScroll, $script:UI.LogScroll + $cap); return }
        'PageDown'  { $script:UI.LogScroll = [Math]::Max(0, $script:UI.LogScroll - $cap); return }
        'Home'      { $script:UI.LogScroll = $maxScroll; return }
        'End'       { $script:UI.LogScroll = 0; return }
    }
    if ([char]::ToUpper($K.KeyChar) -eq 'O') {
        try {
            if ($script:IsWin) { Start-Process notepad.exe -ArgumentList $script:LogFile }
            elseif ($PSVersionTable.PSVersion.Major -ge 6 -and (Get-Variable -Name IsMacOS -ErrorAction SilentlyContinue) -and $IsMacOS) { Start-Process open -ArgumentList $script:LogFile }
            else { Start-Process xdg-open -ArgumentList $script:LogFile }
            Write-SoaLog -Message 'Opened log file in external viewer.'
        } catch {
            Show-MsgModal -Title 'Log' -Lines @('Could not open the log file:', $_.Exception.Message) -Kind Error
        }
    }
}

function Invoke-KeyDispatch {
    param([System.ConsoleKeyInfo]$K)
    $tab = $script:Tabs[$script:UI.Tab]

    # Ctrl+C quits from anywhere
    if (($K.Modifiers -band [ConsoleModifiers]::Control) -and $K.Key -eq 'C') { $script:UI.Quit = $true; return }

    # search mode owns nearly all keys
    if ($script:UI.SearchMode -and $tab['Kind'] -eq 'List') {
        Invoke-ListKey -Tab $tab -K $K
        return
    }

    if ($K.Key -eq 'Tab') {
        $delta = 1
        if ($K.Modifiers -band [ConsoleModifiers]::Shift) { $delta = -1 }
        $script:UI.Tab = ($script:UI.Tab + $delta + $script:Tabs.Count) % $script:Tabs.Count
        return
    }
    if ($K.KeyChar -ge '1' -and $K.KeyChar -le [char]([int][char]'0' + $script:Tabs.Count)) {
        $script:UI.Tab = [int][string]$K.KeyChar - 1
        return
    }
    if ($K.KeyChar -eq '?') { Show-HelpModal; return }
    $upper = [char]::ToUpper($K.KeyChar)
    if ($upper -eq 'Q') { $script:UI.Quit = $true; return }
    if ($upper -eq 'W') {
        if ($script:Conn.Exo -or $script:Conn.Graph) {
            if (Show-ConfirmModal -Title 'Disconnect' -Lines @('Disconnect the Exchange Online and Microsoft Graph sessions?')) {
                Disconnect-AllServices
                foreach ($lt in $script:Tabs) {
                    if ($lt['Kind'] -eq 'List') { $lt['Loaded'] = $false; $lt['Items'] = @(); $lt['View'] = @(); $lt['Cursor'] = 0; $lt['Scroll'] = 0 }
                }
                $script:Org.Loaded = $false
            }
        }
        return
    }

    switch ($tab['Kind']) {
        'List' { Invoke-ListKey -Tab $tab -K $K }
        'Org'  { Invoke-OrgKey -K $K }
        'Log'  { Invoke-LogKey -K $K }
    }
}

#endregion

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
