# ============================================================================
#region Key dispatch
# ============================================================================

function Invoke-ListKey {
    param($Tab, [System.ConsoleKeyInfo]$K)

    # search mode captures input first
    if ($script:UI.SearchMode) {
        if ($K.Key -eq 'Escape') { $script:UI.SearchMode = $false; $Tab['Search'] = ''; Update-TabView -Tab $Tab; return }
        if ($K.Key -eq 'Enter')  { $script:UI.SearchMode = $false; return }
        if ($K.Key -eq 'Backspace') {
            if ($Tab['Search'].Length -gt 0) { $Tab['Search'] = $Tab['Search'].Substring(0, $Tab['Search'].Length - 1); Update-TabView -Tab $Tab }
            return
        }
        if ($K.KeyChar -and -not [char]::IsControl($K.KeyChar)) {
            $Tab['Search'] = $Tab['Search'] + $K.KeyChar
            $Tab['Cursor'] = 0
            Update-TabView -Tab $Tab
        }
        return
    }

    if (-not $Tab['Loaded']) {
        if ($K.Key -eq 'Enter' -or [char]::ToUpper($K.KeyChar) -eq 'R') { Invoke-TabLoad -Tab $Tab -Force }
        return
    }

    $view = @($Tab['View'])
    $cap = [Math]::Max(1, $script:UI.H - 5)

    switch ($K.Key) {
        'UpArrow'   { if ($Tab['Cursor'] -gt 0) { $Tab['Cursor']-- }; return }
        'DownArrow' { if ($Tab['Cursor'] -lt ($view.Count - 1)) { $Tab['Cursor']++ }; return }
        'PageUp'    { $Tab['Cursor'] = [Math]::Max(0, $Tab['Cursor'] - $cap); return }
        'PageDown'  { $Tab['Cursor'] = [Math]::Min([Math]::Max(0, $view.Count - 1), $Tab['Cursor'] + $cap); return }
        'Home'      { $Tab['Cursor'] = 0; return }
        'End'       { $Tab['Cursor'] = [Math]::Max(0, $view.Count - 1); return }
        'Spacebar'  {
            if ($view.Count -gt 0 -and $Tab['Cursor'] -lt $view.Count) {
                $item = $view[$Tab['Cursor']]
                $item.Selected = -not $item.Selected
                if ($Tab['Cursor'] -lt ($view.Count - 1)) { $Tab['Cursor']++ }
            }
            return
        }
        'Enter'     { return }
    }

    switch ([char]::ToUpper($K.KeyChar)) {
        'A' { foreach ($it in $view) { $it.Selected = $true }; return }
        'N' { foreach ($it in @($Tab['Items'])) { $it.Selected = $false }; return }
        '/' { $script:UI.SearchMode = $true; return }
        'F' {
            $order = @('All','Cloud','OnPrem','Pending')
            $idx = [Array]::IndexOf($order, $Tab['Filter'])
            $Tab['Filter'] = $order[(($idx + 1) % $order.Count)]
            $Tab['Cursor'] = 0
            Update-TabView -Tab $Tab
            return
        }
        'S' {
            $order = @('Name','Email','Soa')
            $idx = [Array]::IndexOf($order, $Tab['SortCol'])
            $Tab['SortCol'] = $order[(($idx + 1) % $order.Count)]
            Update-TabView -Tab $Tab
            return
        }
        'D' { $Tab['SortDesc'] = -not $Tab['SortDesc']; Update-TabView -Tab $Tab; return }
        'R' { Invoke-TabLoad -Tab $Tab -Force; return }
        'E' { Export-ViewCsv -Tab $Tab; return }
        'I' { Import-SelectionFile -Tab $Tab; Update-TabView -Tab $Tab; return }
        'C' { Invoke-SoaConversion -Tab $Tab -ToCloud $true; return }
        'O' { Invoke-SoaConversion -Tab $Tab -ToCloud $false; return }
        'V' { Invoke-GroupAudit -Tab $Tab; return }
    }
}

function Invoke-OrgKey {
    param([System.ConsoleKeyInfo]$K)
    if ($K.Key -eq 'Enter' -and -not $script:Org.Loaded) { Invoke-TabLoad -Tab $script:Tabs[3]; return }
    switch ([char]::ToUpper($K.KeyChar)) {
        'E' { Invoke-OrgEnable; return }
        'D' { Invoke-OrgDisable; return }
        'R' { Invoke-TabLoad -Tab $script:Tabs[3]; return }
    }
}

function Invoke-LogKey {
    param([System.ConsoleKeyInfo]$K)
    $cap = [Math]::Max(1, $script:UI.H - 4)
    $maxScroll = [Math]::Max(0, $script:LogBuffer.Count - $cap)
    switch ($K.Key) {
        'UpArrow'   { $script:UI.LogScroll = [Math]::Min($maxScroll, $script:UI.LogScroll + 1); return }
        'DownArrow' { $script:UI.LogScroll = [Math]::Max(0, $script:UI.LogScroll - 1); return }
        'PageUp'    { $script:UI.LogScroll = [Math]::Min($maxScroll, $script:UI.LogScroll + $cap); return }
        'PageDown'  { $script:UI.LogScroll = [Math]::Max(0, $script:UI.LogScroll - $cap); return }
        'Home'      { $script:UI.LogScroll = $maxScroll; return }
        'End'       { $script:UI.LogScroll = 0; return }
    }
    if ([char]::ToUpper($K.KeyChar) -eq 'O') {
        try {
            if ($script:IsWin) { Start-Process notepad.exe -ArgumentList $script:LogFile }
            elseif ($PSVersionTable.PSVersion.Major -ge 6 -and (Get-Variable -Name IsMacOS -ErrorAction SilentlyContinue) -and $IsMacOS) { Start-Process open -ArgumentList $script:LogFile }
            else { Start-Process xdg-open -ArgumentList $script:LogFile }
            Write-SoaLog -Message 'Opened log file in external viewer.'
        } catch {
            Show-MsgModal -Title 'Log' -Lines @('Could not open the log file:', $_.Exception.Message) -Kind Error
        }
    }
}

function Invoke-KeyDispatch {
    param([System.ConsoleKeyInfo]$K)
    $tab = $script:Tabs[$script:UI.Tab]

    # Ctrl+C quits from anywhere
    if (($K.Modifiers -band [ConsoleModifiers]::Control) -and $K.Key -eq 'C') { $script:UI.Quit = $true; return }

    # search mode owns nearly all keys
    if ($script:UI.SearchMode -and $tab['Kind'] -eq 'List') {
        Invoke-ListKey -Tab $tab -K $K
        return
    }

    if ($K.Key -eq 'Tab') {
        $delta = 1
        if ($K.Modifiers -band [ConsoleModifiers]::Shift) { $delta = -1 }
        $script:UI.Tab = ($script:UI.Tab + $delta + $script:Tabs.Count) % $script:Tabs.Count
        return
    }
    if ($K.KeyChar -ge '1' -and $K.KeyChar -le [char]([int][char]'0' + $script:Tabs.Count)) {
        $script:UI.Tab = [int][string]$K.KeyChar - 1
        return
    }
    if ($K.KeyChar -eq '?') { Show-HelpModal; return }
    $upper = [char]::ToUpper($K.KeyChar)
    if ($upper -eq 'Q') { $script:UI.Quit = $true; return }
    if ($upper -eq 'W') {
        if ($script:Conn.Exo -or $script:Conn.Graph) {
            if (Show-ConfirmModal -Title 'Disconnect' -Lines @('Disconnect the Exchange Online and Microsoft Graph sessions?')) {
                Disconnect-AllServices
                foreach ($lt in $script:Tabs) {
                    if ($lt['Kind'] -eq 'List') { $lt['Loaded'] = $false; $lt['Items'] = @(); $lt['View'] = @(); $lt['Cursor'] = 0; $lt['Scroll'] = 0 }
                }
                $script:Org.Loaded = $false
            }
        }
        return
    }

    switch ($tab['Kind']) {
        'List' { Invoke-ListKey -Tab $tab -K $K }
        'Org'  { Invoke-OrgKey -K $K }
        'Log'  { Invoke-LogKey -K $K }
    }
}

#endregion
