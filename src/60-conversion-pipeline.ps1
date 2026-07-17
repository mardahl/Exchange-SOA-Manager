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
                Write-SoaErrorLog -Context ("Get-GroupCloudMembers failed for group '{0}'" -f $g.Name) -ErrorRecord $_
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
            Write-SoaErrorLog -Context 'Save-BackupFile failed' -ErrorRecord $_
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
                Write-SoaErrorLog -Context ("Get-MailboxBackupRecord failed for '{0}'" -f $item.Name) -ErrorRecord $_
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
            Write-SoaErrorLog -Context 'Save-BackupFile (mailboxes) failed' -ErrorRecord $_
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
