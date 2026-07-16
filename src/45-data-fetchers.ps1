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
