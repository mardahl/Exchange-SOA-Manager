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
