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

function Export-GroupAudit {
    # Flatten audit findings (one row per nested group / dropped member) to CSV.
    param([object[]]$Findings)
    if (-not $Findings -or @($Findings).Count -eq 0) { return $null }
    Confirm-SoaDirectory -Path $script:ExportDir
    $file = Join-Path $script:ExportDir ("GroupAudit_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $rows = New-Object System.Collections.ArrayList
    foreach ($f in @($Findings)) {
        $g   = $f.Group
        $rec = $f.Record
        if ([string]$rec.State -eq 'Error') {
            [void]$rows.Add([pscustomobject]@{
                GroupName   = [string]$g.Name
                GroupId     = [string]$g.Id
                FindingType = 'AuditError'
                MemberName  = ''
                MemberSid   = ''
                MemberClass = ''
                Reason      = [string]$rec.Error
            })
            continue
        }
        foreach ($ng in @($rec.NestedGroups)) {
            [void]$rows.Add([pscustomobject]@{
                GroupName   = [string]$g.Name
                GroupId     = [string]$g.Id
                FindingType = 'NestedGroup'
                MemberName  = [string]$ng.Name
                MemberSid   = [string]$ng.Sid
                MemberClass = 'group'
                Reason      = 'nested - stays on-prem, convert separately'
            })
        }
        foreach ($dm in @($rec.DroppedMembers)) {
            [void]$rows.Add([pscustomobject]@{
                GroupName   = [string]$g.Name
                GroupId     = [string]$g.Id
                FindingType = 'DroppedMember'
                MemberName  = [string]$dm.Name
                MemberSid   = [string]$dm.Sid
                MemberClass = [string]$dm.Class
                Reason      = 'not in sync scope - dropped after conversion'
            })
        }
    }
    if ($rows.Count -eq 0) { return $null }
    $rows.ToArray() | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Write-SoaLog -Message ("Wrote {0} audit finding(s) to {1}." -f $rows.Count, $file) -Level OK
    return $file
}

function Invoke-GroupAudit {
    # On-demand forward-conversion audit for selected (else visible) groups.
    param($Tab)
    if ($Tab['Noun'] -ne 'groups') {
        Show-MsgModal -Title 'Forward audit' -Lines @('The forward audit runs on the Groups tab only.') -Kind Info
        return
    }
    $targets = @($Tab['Items'] | Where-Object { $_.Selected })
    if ($targets.Count -eq 0) { $targets = @($Tab['View']) }
    if ($targets.Count -eq 0) {
        Show-MsgModal -Title 'Forward audit' -Lines @('There are no groups to audit.') -Kind Warn
        return
    }

    $pre = Test-AuditPrerequisite
    if (-not $pre.Ok) {
        Show-MsgModal -Title 'Forward audit unavailable' -Lines @($pre.Reason) -Kind Info
        return
    }

    Write-SoaLog -Message ("Forward audit starting for {0} group(s)..." -f $targets.Count)
    $findings = New-Object System.Collections.ArrayList
    $flagged = 0
    $idx = 0
    foreach ($g in $targets) {
        Write-Screen
        Write-ProgressModal -Title 'Forward audit' -Done $idx -Total $targets.Count -Label $g.Name -Ok $flagged -Failed 0
        $rec = $null
        try {
            $rec = Invoke-GroupForwardAudit -Group $g
        } catch {
            Write-SoaLog -Message ("Audit failed for group '{0}': {1}" -f $g.Name, $_.Exception.Message) -Level WARN
            $rec = [pscustomobject]@{ GroupId = [string]$g.Id; State = 'Error'; NestedGroups = @(); DroppedMembers = @(); Error = $_.Exception.Message }
        }
        $g | Add-Member -NotePropertyName Audit -NotePropertyValue $rec -Force
        if ([string]$rec.State -eq 'Yellow' -or [string]$rec.State -eq 'Error') {
            $flagged++
            [void]$findings.Add([pscustomobject]@{ Group = $g; Record = $rec })
        }
        $idx++
    }
    $script:UI.Dirty = $true

    $csv = $null
    if ($findings.Count -gt 0) { $csv = Export-GroupAudit -Findings $findings.ToArray() }

    Write-SoaLog -Message ("Forward audit complete: {0} of {1} group(s) flagged." -f $flagged, $targets.Count)
    if ($flagged -eq 0) {
        Show-MsgModal -Title 'Forward audit complete' -Lines @(
            ("Audited {0} group(s). All clear - no nested groups or dropped members." -f $targets.Count)
        )
    } else {
        $lines = New-Object System.Collections.ArrayList
        [void]$lines.Add(("{0} of {1} group(s) flagged (nested groups or members that would be dropped)." -f $flagged, $targets.Count))
        [void]$lines.Add('')
        [void]$lines.Add('Findings written to:')
        [void]$lines.Add(@($script:T.CtxHi, [string]$csv))
        Show-MsgModal -Title 'Forward audit complete' -Lines $lines.ToArray() -Kind Warn
    }
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
