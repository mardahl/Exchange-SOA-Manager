# ============================================================================
#region Logging
# ============================================================================

function Write-SoaLog {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK','DEBUG')][string]$Level = 'INFO'
    )
    if ($Level -eq 'DEBUG' -and -not $script:DebugLog) { return }
    $stamp = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
    $entry = "[$stamp] [$Level] $Message"
    [void]$script:LogBuffer.Add(@{ Stamp=$stamp; Level=$Level; Message=$Message })
    try { Add-Content -Path $script:LogFile -Value $entry -Encoding UTF8 } catch { }
}

function Write-SoaErrorLog {
    # Full-detail error logging: type, message, Graph/EXO REST body, inner
    # exception chain, stack trace - each as its own Write-SoaLog ERROR line.
    param(
        [Parameter(Mandatory=$true)][string]$Context,
        [Parameter(Mandatory=$true)]$ErrorRecord
    )
    $ex = $ErrorRecord.Exception
    Write-SoaLog -Message $Context -Level ERROR
    Write-SoaLog -Message ("  Type      : {0}" -f $ex.GetType().FullName) -Level ERROR
    Write-SoaLog -Message ("  Message   : {0}" -f $ex.Message) -Level ERROR
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        Write-SoaLog -Message ("  Details   : {0}" -f $ErrorRecord.ErrorDetails.Message) -Level ERROR
    }
    $inner = $ex.InnerException
    while ($inner) {
        Write-SoaLog -Message ("  InnerExc  : {0}: {1}" -f $inner.GetType().FullName, $inner.Message) -Level ERROR
        $inner = $inner.InnerException
    }
    if ($ErrorRecord.CategoryInfo) {
        Write-SoaLog -Message ("  Category  : {0}" -f $ErrorRecord.CategoryInfo.ToString()) -Level ERROR
    }
    if ($ErrorRecord.ScriptStackTrace) {
        foreach ($l in ($ErrorRecord.ScriptStackTrace -split "`n")) {
            Write-SoaLog -Message ("  StackTrace: {0}" -f $l.Trim()) -Level ERROR
        }
    }
}

#endregion
