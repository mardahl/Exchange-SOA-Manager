# ============================================================================
#region Org actions
# ============================================================================

function Invoke-OrgEnable {
    if (-not $script:Org.Loaded) { Invoke-TabLoad -Tab $script:Tabs[3]; if (-not $script:Org.Loaded) { return } }
    if ($script:Org.CloudDefault) {
        Show-MsgModal -Title 'Organization' -Lines @('The cloud-managed default is already enabled.')
        return
    }
    $ok = Show-TypedConfirmModal -Title 'Enable tenant-wide cloud default' -Lines @(
        'This runs:',
        @($script:T.CtxHi, '  Set-OrganizationConfig -ExchangeAttributesCloudManagedByDefault'),
        '',
        'New dir-synced users will sync WITHOUT Exchange attributes from AD and their mailboxes will be cloud-managed once licensed.',
        '',
        @($script:T.Warn, 'Warning: while enabled, newly created on-premises mailboxes do not appear in the Exchange Online GAL. Not recommended while you still actively provision mailboxes on-premises.')
    ) -Word 'ENABLE'
    if (-not $ok) { Write-SoaLog -Message 'Org enable cancelled.'; return }
    if (Set-OrgState -CloudDefault $true) {
        [void](Get-OrgState)
        Show-MsgModal -Title 'Organization' -Lines @('Cloud-managed default is now ENABLED.', '', 'Verify with: Get-OrganizationConfig | FL BlockExchangeProvisioningFromOnPremEnabled')
    }
}

function Invoke-OrgDisable {
    if (-not $script:Org.Loaded) { Invoke-TabLoad -Tab $script:Tabs[3]; if (-not $script:Org.Loaded) { return } }
    if (-not $script:Org.CloudDefault) {
        Show-MsgModal -Title 'Organization' -Lines @('The server-managed default is already active.')
        return
    }
    $ok = Show-TypedConfirmModal -Title 'Revert to server-managed default' -Lines @(
        'This runs:',
        @($script:T.CtxHi, '  Set-OrganizationConfig -ExchangeAttributesServerManagedByDefault'),
        '',
        'New dir-synced mailboxes will again default to on-premises (server) management. Users created while the feature was enabled are not affected.'
    ) -Word 'DISABLE'
    if (-not $ok) { Write-SoaLog -Message 'Org disable cancelled.'; return }
    if (Set-OrgState -CloudDefault $false) {
        [void](Get-OrgState)
        Show-MsgModal -Title 'Organization' -Lines @('Server-managed default restored.')
    }
}

#endregion
