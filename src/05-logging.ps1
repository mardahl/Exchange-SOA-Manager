# ============================================================================
#region Logging
# ============================================================================

function Write-SoaLog {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK')][string]$Level = 'INFO'
    )
    $stamp = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
    $entry = "[$stamp] [$Level] $Message"
    [void]$script:LogBuffer.Add(@{ Stamp=$stamp; Level=$Level; Message=$Message })
    try { Add-Content -Path $script:LogFile -Value $entry -Encoding UTF8 } catch { }
}

#endregion
