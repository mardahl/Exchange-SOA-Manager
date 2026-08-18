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
            $loaded = Get-Module -Name ExchangeOnlineManagement | Select-Object -First 1
            if ($loaded) { Write-SoaLog -Message ("Module loaded: ExchangeOnlineManagement v{0}" -f $loaded.Version) }
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        } catch {
            $script:LastConnectError = $_.Exception.Message
            Write-SoaErrorLog -Context 'Connect-ExoService: Connect-ExchangeOnline failed' -ErrorRecord $_
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
            Write-SoaErrorLog -Context 'Connect-GraphService: Start-GraphWorker failed' -ErrorRecord $_
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
    if ($gw.StdErrTask) {
        try { $gw.StdErrTask.Wait(500) } catch { }
        $gw.StdErrTask = $null
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
# Send non-JSON output (warnings, verbose, info) to stderr so stdout stays a
# clean JSON channel. Parent drains stderr into the debug log.
function Send-Envelope($obj) { Write-Output ('SOA::' + ($obj | ConvertTo-Json -Depth 10 -Compress)) }
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop 3>&1 4>&1 5>&1 6>&1 | ForEach-Object { [Console]::Error.WriteLine($_) }
$scopes = $ScopeList -split '\|'
Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop 3>&1 4>&1 5>&1 6>&1 | ForEach-Object { [Console]::Error.WriteLine($_) }
$ctx = Get-MgContext
$acct = ''
if ($ctx) { $acct = [string]$ctx.Account }
$gmod = Get-Module -Name Microsoft.Graph.Authentication | Select-Object -First 1
$gver = ''
if ($gmod) { $gver = [string]$gmod.Version }
Send-Envelope @{ type='ready'; account=$acct; graphModuleVersion=$gver }
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
                Send-Envelope @{ type='ok'; id=$job.id; value=$resp }
            }
            'PATCH' {
                [void](Invoke-MgGraphRequest -Method PATCH -Uri $job.uri -Body $job.body -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop)
                Send-Envelope @{ type='ok'; id=$job.id }
            }
            'POST' {
                $body = $job.body
                $ct = $job.contentType
                if (-not $ct) { $ct = 'application/json' }
                $resp = Invoke-MgGraphRequest -Method POST -Uri $job.uri -Body $body -ContentType $ct -OutputType PSObject -ErrorAction Stop
                Send-Envelope @{ type='ok'; id=$job.id; value=$resp }
            }
            default { throw "Unknown method $($job.method)" }
        }
    } catch {
        $msg = $_.Exception.Message
        $exType = $_.Exception.GetType().FullName
        if ($_.Exception -is [System.Management.Automation.MethodInvocationException] -and $_.Exception.InnerException) {
            $msg = $_.Exception.InnerException.Message
            $exType = $_.Exception.InnerException.GetType().FullName
        }
        Send-Envelope @{ type='err'; id=$job.id; message=$msg; exceptionType=$exType }
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
    # Drain worker stderr into the debug log on a background task. Keeps the
    # stderr pipe from filling and preserves worker diagnostics for debugging.
    $gw.StdErrTask = [System.Threading.Tasks.Task]::Run([Action]{
        param($sr)
        while ($null -ne ($l = $sr.ReadLine())) {
            Write-SoaLog -Message ("[GraphWorker] {0}" -f $l) -Level DEBUG
        }
    }, $proc.StandardError)
    # Read the ready envelope. Sign-in is synchronous in the child, so this
    # blocks until the user completes auth in the browser. Anything not
    # prefixed with SOA:: on stdout is noise and is logged then skipped.
    $readyObj = $null
    while ($null -ne ($line = $gw.StdOut.ReadLine())) {
        if ($line.StartsWith('SOA::')) {
            $readyObj = $line.Substring(5) | ConvertFrom-Json
            break
        }
        Write-SoaLog -Message ("[GraphWorker stdout noise] {0}" -f $line) -Level DEBUG
    }
    if ($null -eq $readyObj) {
        Stop-GraphWorker
        throw 'Graph worker exited before signalling readiness.'
    }
    if ($readyObj.type -ne 'ready') {
        $err = $readyObj.message
        if (-not $err) { $err = 'Graph worker failed during authentication.' }
        Stop-GraphWorker
        throw $err
    }
    $gw.Account = [string]$readyObj.account
    if ($readyObj.graphModuleVersion) {
        Write-SoaLog -Message ("Module loaded (Graph worker): Microsoft.Graph.Authentication v{0}" -f $readyObj.graphModuleVersion)
    }
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
    Write-SoaLog -Message ("Graph request: {0}" -f $line) -Level DEBUG
    $gw.StdIn.WriteLine($line)
    # Skip stdout noise; only SOA:: envelopes are protocol data. stderr is
    # drained by the background task and shows up in the debug log.
    $obj = $null
    while ($null -ne ($resp = $gw.StdOut.ReadLine())) {
        if ($resp.StartsWith('SOA::')) {
            Write-SoaLog -Message ("Graph response: {0}" -f $resp.Substring(5)) -Level DEBUG
            $obj = $resp.Substring(5) | ConvertFrom-Json
            break
        }
        Write-SoaLog -Message ("[GraphWorker stdout noise] {0}" -f $resp) -Level DEBUG
    }
    if ($null -eq $obj) {
        Stop-GraphWorker
        throw 'Graph worker closed the response stream.'
    }
    if ($obj.type -eq 'err') {
        $msg = [string]$obj.message
        if ($obj.exceptionType) { $msg = "{0} ({1})" -f $msg, [string]$obj.exceptionType }
        throw $msg
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
