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

function Get-AuditGlyph {
    param($Item, [int]$Width = 4)
    $t = $script:T; $g = $script:G
    $rec = Get-PropSafe $Item 'Audit'
    if ($null -eq $rec) { return ($t.Muted + (Get-PadCell '' $Width)) }
    $state = [string](Get-PropSafe $rec 'State')
    switch ($state) {
        'Yellow' { return ($t.Warn + (Get-PadCell ([string]$g.AuditWarn) $Width)) }
        'Green'  { return ($t.Good + (Get-PadCell ([string]$g.AuditOk) $Width)) }
        'Error'  { return ($t.Danger + (Get-PadCell ([string]$g.AuditError) $Width)) }
        default  { return ($t.Muted + (Get-PadCell '?' $Width)) }
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
