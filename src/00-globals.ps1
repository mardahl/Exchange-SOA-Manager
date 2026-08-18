# ============================================================================
#region Globals & State
# ============================================================================

$script:Version = '1.5.3'
$script:ESC     = [char]27
$script:IsWin   = ($PSVersionTable.PSVersion.Major -lt 6) -or ($null -ne (Get-Variable -Name IsWindows -ErrorAction SilentlyContinue) -and $IsWindows)

if (-not $script:Root) {
    if ($PSScriptRoot) { $script:Root = $PSScriptRoot } else { $script:Root = (Get-Location).Path }
}

$script:LogFile   = Join-Path $script:Root ("SOA-Manager_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$script:DebugLog  = $false
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
        AuditOk='OK'; AuditWarn='!'; AuditError='X'
    }
} else {
    $script:G = @{
        H=([char]0x2500); V=([char]0x2502); TL=([char]0x250C); TR=([char]0x2510); BL=([char]0x2514); BR=([char]0x2518)
        Dot=([char]0x25CF); Half=([char]0x25D0); Ring=([char]0x25CB)
        BarOn=([char]0x2588); BarOff=([char]0x2591)
        Up=([char]0x2191); Down=([char]0x2193); Ell=([char]0x2026)
        ChkOn=('[' + [char]0x25A0 + ']'); ChkOff='[ ]'
        Arrow=([char]0x2192)
        AuditOk=([char]0x2713); AuditWarn=([char]0x25B2); AuditError=([char]0x2716)
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
