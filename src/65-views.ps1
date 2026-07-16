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
    # Columns: ' ' chk(3) ' ' name email '  ' detail(13) '  ' soa(10) '  ' audit(4)
    $fixed = 1 + 3 + 1 + 2 + 13 + 2 + 10 + 2 + 4
    $flex = $W - $fixed - 1
    if ($flex -lt 20) { $flex = 20 }
    $nameW = [int]($flex * 0.42)
    $emailW = $flex - $nameW - 2
    return @{ Name=$nameW; Email=$emailW; Detail=13; Soa=10; Chk=4 }
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
    $head = ' ' + (Get-PadCell 'sel' 3) + ' ' + (Get-PadCell 'Name' $col.Name) + '  ' + (Get-PadCell 'Email' $col.Email) + '  ' + (Get-PadCell 'Type' $col.Detail) + '  ' + (Get-PadCell 'SOA' $col.Soa) + '  ' + (Get-PadCell 'Chk' $col.Chk)
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
        $line += '  ' + (Get-AuditGlyph -Item $item -Width $col.Chk)
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
                @('C','to cloud'), @('O','to on-prem'), @('V','audit'),
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
