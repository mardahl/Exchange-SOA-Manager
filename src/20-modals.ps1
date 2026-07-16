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
        @($t.Row, '  V                    audit selected groups for cloud-conversion risk (Windows only)'),
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
