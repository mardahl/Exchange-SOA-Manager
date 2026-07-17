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
    Write-SoaLog -Message ("Set-Mailbox -Identity '{0}' -IsExchangeCloudManaged {1}" -f $Item.Id, $ToCloud) -Level DEBUG
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
        Write-SoaErrorLog -Context ("Convert-MailboxSoa failed for '{0}'" -f $Item.Name) -ErrorRecord $_
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
        Write-SoaLog -Message ("PATCH {0} body={1}" -f $uri, $body) -Level DEBUG
        [void](Invoke-GraphWorker -Job @{ method='PATCH'; uri=$uri; body=$body; contentType='application/json' })
        return @{ Ok=$true; Msg='converted' }
    } catch {
        Write-SoaErrorLog -Context ("Convert-GraphSoa failed for '{0}'" -f $Item.Name) -ErrorRecord $_
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
    # Separate from $visited: $visited guards re-traversal of a group's children
    # (cycle prevention). $recordedSids dedups *reported* nested groups/leaves,
    # since the same SID can be reached via two different parent paths (diamond
    # membership) without ever revisiting an already-traversed group.
    $recordedSids = New-Object 'System.Collections.Generic.HashSet[string]'
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
            if (-not $sid -or -not $recordedSids.Contains($sid)) {
                if ($sid) { [void]$recordedSids.Add($sid) }
                [void]$nested.Add([pscustomobject]@{
                    Name  = [string]$o.displayName
                    DN    = [string]$o.DistinguishedName
                    Sid   = $sid
                    Class = 'group'
                })
            }
            if ($sid -and -not $visited.Contains($sid)) {
                [void]$visited.Add($sid)
                $child = Get-ADGroup -Identity $dn -Properties member -ErrorAction SilentlyContinue
                if ($child) { foreach ($m in @($child.member)) { $stack.Push([string]$m) } }
            }
        } else {
            if (-not $sid -or -not $recordedSids.Contains($sid)) {
                if ($sid) { [void]$recordedSids.Add($sid) }
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
        # State contract: 'Green' = verified clean, 'Yellow' = findings present.
        # ('Error' is assigned by the caller's catch block on a thrown exception -
        # never returned from here, since this function's own audit succeeded.)
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

    # State contract: 'Green' = verified clean, 'Yellow' = findings present.
    # ('Error' is assigned by the caller's catch block on a thrown exception -
    # never returned from here, since this function's own audit succeeded.)
    $state = 'Green'
    if (@($tree.Nested).Count -gt 0 -or $dropped.Count -gt 0) { $state = 'Yellow' }
    return [pscustomobject]@{
        GroupId = [string]$Group.Id; State = $state
        NestedGroups = @($tree.Nested); DroppedMembers = $dropped.ToArray(); Error = $null
    }
}

#endregion
