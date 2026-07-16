# ============================================================================
#region Graph REST helpers
# ============================================================================

function Invoke-GraphGetAll {
    # Paged GET; returns array of PSObjects from .value.
    # -OnPage (optional) is invoked with the cumulative item count after each page.
    param([string]$Uri, [switch]$Advanced, [scriptblock]$OnPage)
    $headers = @{}
    if ($Advanced) { $headers['ConsistencyLevel'] = 'eventual' }
    $out = New-Object System.Collections.ArrayList
    $next = $Uri
    while ($next) {
        $resp = Invoke-GraphWorker -Job @{ method='GET'; uri=$next; headers=$headers }
        $val = Get-PropSafe $resp 'value'
        if ($null -ne $val) { foreach ($v in @($val)) { [void]$out.Add($v) } }
        $next = Get-PropSafe $resp '@odata.nextLink'
        if ($OnPage) { & $OnPage $out.Count }
    }
    return ,$out.ToArray()
}

function Get-SyncBehaviorMap {
    # Batched lookup of isCloudManaged for a set of object ids.
    # $Resource: 'groups' or 'contacts'. Returns hashtable id -> [bool]
    # -Progress (optional) is invoked per batch with (count, label).
    param([string[]]$Ids, [string]$Resource, [scriptblock]$Progress)
    $map = @{}
    if (-not $Ids -or $Ids.Count -eq 0) { return $map }
    $chunkSize = 20
    for ($i = 0; $i -lt $Ids.Count; $i += $chunkSize) {
        $end = [Math]::Min($i + $chunkSize - 1, $Ids.Count - 1)
        $chunk = $Ids[$i..$end]
        $requests = New-Object System.Collections.ArrayList
        $n = 1
        foreach ($id in $chunk) {
            [void]$requests.Add(@{
                id     = "$n"
                method = 'GET'
                url    = ('/' + $Resource + '/' + $id + '/onPremisesSyncBehavior?$select=isCloudManaged')
            })
            $n++
        }
        $body = @{ requests = $requests.ToArray() } | ConvertTo-Json -Depth 4
        $resp = Invoke-GraphWorker -Job @{ method='POST'; uri='https://graph.microsoft.com/v1.0/$batch'; body=$body; contentType='application/json' }
        $responses = Get-PropSafe $resp 'responses'
        if ($null -eq $responses) { continue }
        foreach ($r in @($responses)) {
            $rid = [int](Get-PropSafe $r 'id')
            $status = [int](Get-PropSafe $r 'status')
            $objId = $chunk[$rid - 1]
            if ($status -eq 200) {
                $rbody = Get-PropSafe $r 'body'
                $icm = Get-PropSafe $rbody 'isCloudManaged'
                $map[$objId] = [bool]$icm
            } else {
                Write-SoaLog -Message ("onPremisesSyncBehavior lookup failed for {0}/{1} (HTTP {2})." -f $Resource, $objId, $status) -Level WARN
            }
        }
        if ($Progress) {
            & $Progress ($end + 1) ('Checking SOA state - {0} of {1} converted {2}' -f ($end + 1), $Ids.Count, $Resource)
        }
    }
    return $map
}

#endregion
