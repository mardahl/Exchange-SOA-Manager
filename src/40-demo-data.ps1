# ============================================================================
#region Demo data
# ============================================================================

$script:DemoFirst = @('Anna','Bjorn','Carla','David','Elena','Felix','Greta','Hugo','Ines','Jonas','Klara','Lars','Mona','Nils','Olga','Per','Quinn','Rosa','Sven','Tara','Ulf','Vera','Wim','Xena','Yara','Zane','Astrid','Bruno','Celine','Dario','Edith','Frans','Gilda','Henrik','Iris','Joost','Kira','Liam','Mara','Noor')
$script:DemoLast  = @('Andersen','Bergstrom','Carlsen','Dahl','Eriksen','Fischer','Gruber','Hansen','Iversen','Jansen','Koch','Lindgren','Meyer','Nielsen','Olsen','Petersen','Qvist','Rasmussen','Sorensen','Thomsen','Ulrich','Vogel','Weber','Xander','Ylva','Zimmermann','Abel','Brandt','Clausen','Dietrich','Engel','Falk','Gerber','Holm','Ibsen','Jung','Krause','Lund','Moller','Nygaard')

function New-DemoMailboxes {
    $rng = New-Object System.Random(42)
    $items = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt 340; $i++) {
        $fn = $script:DemoFirst[$rng.Next(0, $script:DemoFirst.Count)]
        $ln = $script:DemoLast[$rng.Next(0, $script:DemoLast.Count)]
        $name = "$fn $ln"
        $upn = ('{0}.{1}{2}@contoso.com' -f $fn.ToLower(), $ln.ToLower(), $i)
        $kind = 'UserMailbox'
        if (($i % 9) -eq 0) { $kind = 'SharedMailbox'; $name = "SM $ln team $i"; $upn = "shared.$($ln.ToLower())$i@contoso.com" }
        elseif (($i % 23) -eq 0) { $kind = 'RoomMailbox'; $name = "Room $ln $i"; $upn = "room.$($ln.ToLower())$i@contoso.com" }
        $soa = 'OnPrem'
        if ($rng.Next(0, 100) -lt 30) { $soa = 'Cloud' }
        [void]$items.Add([pscustomobject]@{
            Type='Mailbox'; Id=$upn; Name=$name; Email=$upn; Detail=$kind; Soa=$soa; Selected=$false
            Raw=[pscustomobject]@{ DisplayName=$name; UserPrincipalName=$upn; PrimarySmtpAddress=$upn; RecipientTypeDetails=$kind; IsDirSynced=$true }
        })
    }
    return ,($items.ToArray() | Sort-Object -Property Name)
}

function New-DemoGroups {
    $rng = New-Object System.Random(1337)
    $kinds = @('Security','Distribution','Mail-sec','Security','Distribution')
    $words = @('Finance','HR','Sales','Engineering','Support','Legal','Marketing','Ops','Research','Field','Branch','Procurement','Payroll','Helpdesk','Admins','Auditors','Interns','Leads')
    $items = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt 72; $i++) {
        $w1 = $words[$rng.Next(0, $words.Count)]
        $w2 = $words[$rng.Next(0, $words.Count)]
        $name = "GRP $w1 $w2 $i"
        $mailNick = ('grp.{0}.{1}{2}' -f $w1.ToLower(), $w2.ToLower(), $i)
        $kind = $kinds[$i % $kinds.Count]
        $mail = ''
        if ($kind -ne 'Security') { $mail = "$mailNick@contoso.com" }
        $roll = $rng.Next(0, 100)
        $soa = 'OnPrem'
        if ($roll -lt 22) { $soa = 'Cloud' } elseif ($roll -lt 30) { $soa = 'Pending' }
        [void]$items.Add([pscustomobject]@{
            Type='Group'; Id=([guid]::NewGuid().ToString()); Name=$name; Email=$mail; Detail=$kind; Soa=$soa; Selected=$false
            Raw=[pscustomobject]@{ displayName=$name; mail=$mail; mailEnabled=($kind -ne 'Security'); securityEnabled=($kind -ne 'Distribution'); onPremisesSyncEnabled=($soa -eq 'OnPrem') }
        })
    }
    return ,($items.ToArray() | Sort-Object -Property Name)
}

function New-DemoContacts {
    $rng = New-Object System.Random(7)
    $items = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt 26; $i++) {
        $fn = $script:DemoFirst[$rng.Next(0, $script:DemoFirst.Count)]
        $ln = $script:DemoLast[$rng.Next(0, $script:DemoLast.Count)]
        $name = "$fn $ln (ext)"
        $mail = ('{0}.{1}@partner{2}.example' -f $fn.ToLower(), $ln.ToLower(), ($i % 6))
        $soa = 'OnPrem'
        if ($rng.Next(0, 100) -lt 20) { $soa = 'Cloud' }
        [void]$items.Add([pscustomobject]@{
            Type='Contact'; Id=([guid]::NewGuid().ToString()); Name=$name; Email=$mail; Detail='Mail contact'; Soa=$soa; Selected=$false
            Raw=[pscustomobject]@{ displayName=$name; mail=$mail; onPremisesSyncEnabled=($soa -eq 'OnPrem') }
        })
    }
    return ,($items.ToArray() | Sort-Object -Property Name)
}

#endregion
