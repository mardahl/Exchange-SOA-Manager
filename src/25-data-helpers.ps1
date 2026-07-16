# ============================================================================
#region Data helpers
# ============================================================================

function Get-PropSafe {
    # StrictMode-safe property access for dynamic/API objects.
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Confirm-SoaDirectory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        [void](New-Item -ItemType Directory -Path $Path -Force)
    }
}

function Test-SoaModule {
    param([string]$Name)
    return ($null -ne (Get-Module -ListAvailable -Name $Name | Select-Object -First 1))
}

function Install-SoaModule {
    param([string]$Name)
    $ok = Show-ConfirmModal -Title 'Module required' -Lines @(
        "The PowerShell module '$Name' is not installed.",
        '',
        'Install it now for the current user?',
        @($script:T.Muted, "Install-Module $Name -Scope CurrentUser")
    )
    if (-not $ok) { return $false }
    Invoke-OnMainBuffer -Action {
        Write-Host ''
        Write-Host "Installing $Name (CurrentUser scope)..." -ForegroundColor Cyan
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Host "Installed $Name." -ForegroundColor Green
        } catch {
            Write-Host "Install failed: $($_.Exception.Message)" -ForegroundColor Red
            Start-Sleep -Seconds 3
        }
    }
    $success = Test-SoaModule -Name $Name
    if ($success) { Write-SoaLog -Message "Installed module $Name." -Level OK }
    else { Write-SoaLog -Message "Failed to install module $Name." -Level ERROR }
    return $success
}

function Test-RsatAd {
    return ($null -ne (Get-Module -ListAvailable -Name 'ActiveDirectory' | Select-Object -First 1))
}

function Install-RsatAd {
    # RSAT ActiveDirectory is a Windows capability/feature, not a Gallery module.
    # Client SKUs (ProductType 1) use Add-WindowsCapability; servers use the feature.
    $isClient = $true
    try { $isClient = ([int](Get-CimInstance Win32_OperatingSystem).ProductType -eq 1) } catch { $isClient = $true }
    $cmd = if ($isClient) {
        'Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'
    } else {
        'Install-WindowsFeature RSAT-AD-PowerShell'
    }
    $ok = Show-ConfirmModal -Title 'RSAT required' -Lines @(
        'The RSAT ActiveDirectory PowerShell module is required for the forward audit.',
        '',
        'Install it now? (requires an elevated / admin session)',
        @($script:T.Muted, "  $cmd")
    )
    if (-not $ok) { return $false }
    Invoke-OnMainBuffer -Action {
        Write-Host ''
        Write-Host 'Installing RSAT ActiveDirectory tools...' -ForegroundColor Cyan
        try {
            if ($isClient) {
                Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -ErrorAction Stop | Out-Null
            } else {
                Install-WindowsFeature -Name 'RSAT-AD-PowerShell' -ErrorAction Stop | Out-Null
            }
            Write-Host 'Installed RSAT ActiveDirectory tools.' -ForegroundColor Green
        } catch {
            Write-Host "Install failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host 'You may need an elevated session. Run the command shown above manually.' -ForegroundColor Yellow
            Start-Sleep -Seconds 3
        }
    }
    $success = Test-RsatAd
    if ($success) { Write-SoaLog -Message 'Installed RSAT ActiveDirectory module.' -Level OK }
    else { Write-SoaLog -Message 'Failed to install RSAT ActiveDirectory module.' -Level ERROR }
    return $success
}

function Test-AuditPrerequisite {
    # Returns Ok=$true when the forward audit can run (or in demo mode).
    if ($script:DemoMode) { return [pscustomobject]@{ Ok = $true; Reason = '' } }
    if (-not $script:IsWin) {
        return [pscustomobject]@{ Ok = $false; Reason = 'The forward audit requires a domain-joined Windows host. Active Directory is not reachable from this platform.' }
    }
    if (-not (Test-RsatAd)) {
        if (-not (Install-RsatAd)) {
            return [pscustomobject]@{ Ok = $false; Reason = 'The RSAT ActiveDirectory module is required and is not installed.' }
        }
    }
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        [void](Get-ADDomain -ErrorAction Stop)
    } catch {
        return [pscustomobject]@{ Ok = $false; Reason = ('Active Directory is not reachable: ' + $_.Exception.Message) }
    }
    return [pscustomobject]@{ Ok = $true; Reason = '' }
}

#endregion
