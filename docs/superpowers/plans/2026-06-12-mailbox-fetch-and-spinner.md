# Mailbox Fetch Optimization + Always-Alive Spinner Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Cut mailbox-load memory by projecting `Get-Mailbox` results to slim objects in one streaming pass, and keep the `\ | / -` spinner animating from a background runspace while the main thread is blocked.

**Architecture:** All changes live in the single-file TUI `SOA-Manager.ps1`. Part 1 rewrites `Get-MailboxItems` to project each streamed mailbox immediately (no full-object buffer, no second pass, slim `Raw`). Part 2 adds `Start-LoadSpinner`/`Stop-LoadSpinner` (background runspace + synchronized hashtable), coordinate publishing in `Write-ProgressModal`, and start/stop wiring in `Invoke-TabLoad`.

**Tech Stack:** PowerShell (5.1-compatible syntax only — no ternary, `??`, `?.`), `[powershell]::Create()` runspace, ANSI/VT console rendering. Verification: AST-extraction test harness + parse check on pwsh, PSScriptAnalyzer.

**Spec:** `docs/superpowers/specs/2026-06-12-mailbox-fetch-and-spinner-design.md`

---

### Task 1: Test harness for Get-MailboxItems projection (red)

**Files:**
- Create: `/var/folders/j6/yg2mgkn91rz2r30bxy8ps6d40000gp/T/opencode/soa-mailbox-tests.ps1` (throwaway harness, NOT committed — repo has no test infrastructure; CI is PSSA + parse checks)

The harness extracts the real functions from the script via AST, fakes `Get-Mailbox`, and asserts the new behavior (slim 7-property `Raw`). It must FAIL before Task 2 (current code stores the full object in `Raw`) and PASS after.

- [ ] **Step 1: Write the harness**

```powershell
# soa-mailbox-tests.ps1
$ErrorActionPreference = 'Stop'
$scriptPath = '/Users/mum@inciro.com/Documents/opencode/SOAconverter/SOA-Manager.ps1'
$tokens = $null; $parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ("Parse errors:`n" + ($parseErrors | Out-String)) }
$wanted = @('Get-MailboxItems', 'Get-PropSafe', 'Format-ElapsedTime')
$funcs = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $wanted -contains $n.Name }, $true)
if (@($funcs).Count -ne 3) { throw "Expected 3 functions, found $(@($funcs).Count)" }
foreach ($f in $funcs) { . ([scriptblock]::Create($f.Extent.Text)) }

$script:DemoMode = $false
function Write-SoaLog { param($Message, $Level) }
function Get-Mailbox {
    param($Filter, $ResultSize, $ErrorAction)
    # Junk1 stands in for the ~200 extra properties of a real mailbox object.
    [pscustomobject]@{ DisplayName='Alice A'; UserPrincipalName='alice@x.test'; PrimarySmtpAddress='alice@x.test'; RecipientTypeDetails='UserMailbox';   IsDirSynced=$true;  IsExchangeCloudManaged=$false; ExternalDirectoryObjectId='id-alice';  Junk1=('x'*10000) }
    [pscustomobject]@{ DisplayName='Bob B';   UserPrincipalName='bob@x.test';   PrimarySmtpAddress='bob@x.test';   RecipientTypeDetails='SharedMailbox'; IsDirSynced=$true;  IsExchangeCloudManaged=$true;  ExternalDirectoryObjectId='id-bob';    Junk1=('y'*10000) }
    [pscustomobject]@{ DisplayName='Cloudy C';UserPrincipalName='cloudy@x.test';PrimarySmtpAddress='cloudy@x.test';RecipientTypeDetails='UserMailbox';   IsDirSynced=$false; IsExchangeCloudManaged=$true;  ExternalDirectoryObjectId='id-cloudy'; Junk1=('z'*10000) }
}

$calls = New-Object System.Collections.ArrayList
# Mirror app usage (Invoke-TabLoad): assign first, then wrap with @().
$result = Get-MailboxItems -Progress { param($n, $l) [void]$calls.Add($n) }
$items = @($result)

if ($items.Count -ne 2) { throw "Expected 2 dir-synced items, got $($items.Count)" }
if ($items[0].Name -ne 'Alice A' -or $items[1].Name -ne 'Bob B') { throw "Sort order wrong: $($items.Name -join ', ')" }
if ($items[0].Soa -ne 'OnPrem') { throw "Alice Soa: $($items[0].Soa)" }
if ($items[1].Soa -ne 'Cloud') { throw "Bob Soa: $($items[1].Soa)" }
if ($items[1].Id -ne 'id-bob') { throw "Bob Id: $($items[1].Id)" }
if ($items[1].Email -ne 'bob@x.test') { throw "Bob Email: $($items[1].Email)" }
if ($items[1].Detail -ne 'SharedMailbox') { throw "Bob Detail: $($items[1].Detail)" }
$rawProps = @($items[0].Raw.PSObject.Properties.Name)
if ($rawProps -contains 'Junk1') { throw 'FAIL: Raw still carries the full mailbox object (Junk1 present)' }
if ($rawProps.Count -ne 7) { throw "Raw should have exactly 7 properties, has $($rawProps.Count): $($rawProps -join ', ')" }
if (@($calls).Count -ne 3) { throw "Progress called $(@($calls).Count) times; expected 3 (one per received mailbox)" }
Write-Output 'PASS: Get-MailboxItems projection tests'
```

- [ ] **Step 2: Run it to verify it FAILS against current code**

Run: `pwsh-preview -NoProfile -File /var/folders/j6/yg2mgkn91rz2r30bxy8ps6d40000gp/T/opencode/soa-mailbox-tests.ps1`
Expected: FAIL with `Raw still carries the full mailbox object (Junk1 present)` (or the progress-count check: current code fires an extra "Filtering..." callback).

### Task 2: Rewrite Get-MailboxItems (green)

**Files:**
- Modify: `SOA-Manager.ps1:1151-1201` (function `Get-MailboxItems`, replace entire body)

- [ ] **Step 1: Replace the function with the single-pass projection version**

```powershell
function Get-MailboxItems {
    # -Progress (optional) is invoked with (count, label) as results stream in.
    param([scriptblock]$Progress)
    if ($script:DemoMode) { return ,(New-DemoMailboxes) }
    Write-SoaLog -Message 'Retrieving dir-synced mailboxes from Exchange Online...'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $items = New-Object System.Collections.ArrayList
    $stats = @{ Received = 0 }
    # Get-Mailbox has no -Properties parameter (and Get-EXOMailbox, which has
    # one, does not expose IsExchangeCloudManaged), so the full 200+ property
    # object always crosses the wire. Project each one to a slim record as
    # soon as it streams in so the heavy object can be garbage-collected
    # instead of being buffered - and kept alive in Raw - for the session.
    $build = {
        param($mb)
        $stats.Received++
        if ($Progress) { & $Progress $stats.Received ('{0} mailboxes received - {1} elapsed' -f $stats.Received, (Format-ElapsedTime $sw)) }
        if (-not [bool](Get-PropSafe $mb 'IsDirSynced')) { return }
        $cloud = [bool](Get-PropSafe $mb 'IsExchangeCloudManaged')
        $slim = [pscustomobject]@{
            DisplayName               = [string](Get-PropSafe $mb 'DisplayName')
            UserPrincipalName         = [string](Get-PropSafe $mb 'UserPrincipalName')
            PrimarySmtpAddress        = [string](Get-PropSafe $mb 'PrimarySmtpAddress')
            RecipientTypeDetails      = [string](Get-PropSafe $mb 'RecipientTypeDetails')
            IsDirSynced               = $true
            IsExchangeCloudManaged    = $cloud
            ExternalDirectoryObjectId = [string](Get-PropSafe $mb 'ExternalDirectoryObjectId')
        }
        $soa = 'OnPrem'
        if ($cloud) { $soa = 'Cloud' }
        $id = $slim.ExternalDirectoryObjectId
        if ([string]::IsNullOrEmpty($id)) { $id = $slim.UserPrincipalName }
        [void]$items.Add([pscustomobject]@{
            Type     = 'Mailbox'
            Id       = $id
            Name     = $slim.DisplayName
            Email    = $slim.PrimarySmtpAddress
            Detail   = $slim.RecipientTypeDetails
            Soa      = $soa
            Selected = $false
            Raw      = $slim
        })
    }
    try {
        # Retrieve only dir-synced mailboxes from the server side. This is
        # extremely fast and avoids fetching all cloud-only mailboxes (which
        # can cause timeouts or throttling in large tenants).
        Get-Mailbox -Filter "IsDirSynced -eq `$true" -ResultSize Unlimited -ErrorAction Stop | ForEach-Object { & $build $_ }
    } catch [System.OperationCanceledException] {
        throw
    } catch {
        # Fallback in case of filter issues; $build skips non-dir-synced items.
        Write-SoaLog -Message ("Filtered Get-Mailbox failed ({0}); trying unrestricted Get-Mailbox." -f $_.Exception.Message) -Level WARN
        $items.Clear()
        $stats.Received = 0
        Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | ForEach-Object { & $build $_ }
    }
    Write-SoaLog -Message ("Retrieved {0} mailboxes; {1} are dir-synced (shown)." -f $stats.Received, $items.Count) -Level OK
    return ,($items.ToArray() | Sort-Object -Property Name)
}
```

- [ ] **Step 2: Run the harness to verify it PASSES**

Run: `pwsh-preview -NoProfile -File /var/folders/j6/yg2mgkn91rz2r30bxy8ps6d40000gp/T/opencode/soa-mailbox-tests.ps1`
Expected: `PASS: Get-MailboxItems projection tests`

- [ ] **Step 3: Commit**

```bash
git add SOA-Manager.ps1
git commit -m "Optimize mailbox loading: single-pass slim projection in Get-MailboxItems"
```

### Task 3: Spinner state global + Start/Stop-LoadSpinner functions

**Files:**
- Modify: `SOA-Manager.ps1:100` (after `$script:LogBuffer = ...`, add spinner global)
- Modify: `SOA-Manager.ps1:658` (insert two functions immediately before `function Write-ProgressModal`)

- [ ] **Step 1: Add the global** (after line 100, `$script:LogBuffer = New-Object System.Collections.ArrayList`)

```powershell
$script:Spinner   = $null   # background-runspace spinner (see Start-LoadSpinner)
```

- [ ] **Step 2: Insert the two functions before `function Write-ProgressModal`**

```powershell
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
```

- [ ] **Step 3: Parse check**

Run: `pwsh-preview -NoProfile -Command "[void][System.Management.Automation.Language.Parser]::ParseFile('/Users/mum@inciro.com/Documents/opencode/SOAconverter/SOA-Manager.ps1', [ref]\$null, [ref]\$e); if (\$e.Count) { \$e; throw 'parse errors' } else { 'Parse OK' }"`
Expected: `Parse OK`

### Task 4: Publish spinner coordinates from Write-ProgressModal

**Files:**
- Modify: `SOA-Manager.ps1` function `Write-ProgressModal`, after the box geometry is computed (currently lines 709-713, the lines computing `$boxW`, `$x`, `$y`)

- [ ] **Step 1: Insert after the `$y = [Math]::Max(...)` line**

```powershell
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
```

(Coordinates: title row is `$y`; body rows start at `$y + 1`; the RAWBAR line is body index 2, so row `$y + 3`. Content starts at column `$x + 2` (border + space); the marquee bar is `$barW` cells, then 3 spaces, then the spinner cell. Set `Y` before `X` because `X -gt 0` is the writer's visibility gate.)

- [ ] **Step 2: Parse check** (same command as Task 3 Step 3). Expected: `Parse OK`

### Task 5: Start/stop the spinner in Invoke-TabLoad

**Files:**
- Modify: `SOA-Manager.ps1` function `Invoke-TabLoad` (currently lines 1820-1850)

- [ ] **Step 1: Start the spinner before the initial modal paint**

Replace:

```powershell
    Write-Screen
    $title = 'Loading ' + $Tab['Name']
    Write-ProgressModal -Title $title -Done 0 -Total 0 -Label 'Contacting service - waiting for first results...' -Ok 0 -Failed 0
```

with:

```powershell
    Write-Screen
    $title = 'Loading ' + $Tab['Name']
    # Start the spinner first so this initial paint already publishes its
    # coordinates - the whole point is animating before the first results.
    Start-LoadSpinner
    Write-ProgressModal -Title $title -Done 0 -Total 0 -Label 'Contacting service - waiting for first results...' -Ok 0 -Failed 0
```

- [ ] **Step 2: Wrap the fetch in try/finally so the spinner stops before any catch-handler modal renders**

Replace:

```powershell
    try {
        $items = @()
        switch ($Tab['Noun']) {
            'mailboxes' { $items = Get-MailboxItems -Progress $progressCb }
            'groups'    { $items = Get-GroupItems -Progress $progressCb }
            'contacts'  { $items = Get-ContactItems -Progress $progressCb }
        }
        $Tab['Items'] = @($items)
```

with:

```powershell
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
```

- [ ] **Step 3: Parse check** (same command as Task 3 Step 3). Expected: `Parse OK`

- [ ] **Step 4: Spinner integration harness** — extract and exercise the new functions:

Create and run `/var/folders/j6/yg2mgkn91rz2r30bxy8ps6d40000gp/T/opencode/soa-spinner-tests.ps1`:

```powershell
$ErrorActionPreference = 'Stop'
$scriptPath = '/Users/mum@inciro.com/Documents/opencode/SOAconverter/SOA-Manager.ps1'
$tokens = $null; $parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ("Parse errors:`n" + ($parseErrors | Out-String)) }
$wanted = @('Start-LoadSpinner', 'Stop-LoadSpinner')
$funcs = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $wanted -contains $n.Name }, $true)
if (@($funcs).Count -ne 2) { throw "Expected 2 spinner functions, found $(@($funcs).Count)" }
foreach ($f in $funcs) { . ([scriptblock]::Create($f.Extent.Text)) }
$script:Spinner = $null
$script:T = @{ Row = ''; Reset = '' }
function Write-SoaLog { param($Message, $Level) }

Stop-LoadSpinner   # idempotent when nothing is running
Start-LoadSpinner
if (-not $script:Spinner) { throw 'Spinner did not start' }
Start-LoadSpinner  # double-start guard must not replace the instance
$script:Spinner.State.Y = 2
$script:Spinner.State.X = 10
Start-Sleep -Milliseconds 600   # main thread blocked; spinner should write frames
Stop-LoadSpinner
if ($script:Spinner) { throw 'Spinner still set after Stop' }
Stop-LoadSpinner   # idempotent again
Write-Output ''
Write-Output 'PASS: spinner start/stop tests'
```

Run: `pwsh-preview -NoProfile -File /var/folders/j6/yg2mgkn91rz2r30bxy8ps6d40000gp/T/opencode/soa-spinner-tests.ps1`
Expected: output contains positioned `| / - \` frames followed by `PASS: spinner start/stop tests`

- [ ] **Step 5: Commit**

```bash
git add SOA-Manager.ps1
git commit -m "Keep load spinner animating from a background runspace while the pipeline blocks"
```

### Task 6: PSScriptAnalyzer + full parse verification

**Files:** none modified

- [ ] **Step 1: Run PSScriptAnalyzer with the CI settings**

Run (mirror of `.github/workflows/ci.yml`):

```bash
pwsh-preview -NoProfile -Command "
  if (-not (Get-Module -ListAvailable PSScriptAnalyzer)) { Install-Module PSScriptAnalyzer -Scope CurrentUser -Force }
  \$settings = @{ Rules = @{ PSUseCompatibleSyntax = @{ Enable = \$true; TargetVersions = @('5.1','7.0') } } }
  \$r = Invoke-ScriptAnalyzer -Path ./SOA-Manager.ps1 -Settings \$settings -Severity Warning,Error
  \$r | Format-Table -AutoSize
  if (\$r | Where-Object Severity -eq 'Error') { throw 'PSSA errors' } else { 'PSSA OK' }"
```

Expected: `PSSA OK` and no new warnings versus main.

- [ ] **Step 2: Re-run both harnesses one last time** (commands in Task 2 Step 2 and Task 5 Step 4). Expected: both PASS.

### Task 7: CHANGELOG + version bump

**Files:**
- Modify: `SOA-Manager.ps1:91` (`$script:Version = '1.2.3'` → `'1.3.0'`)
- Modify: `CHANGELOG.md` (insert release section under `## [Unreleased]`)

- [ ] **Step 1: Bump version** to `1.3.0` (new feature → minor bump per SemVer).

- [ ] **Step 2: Add CHANGELOG section** below `## [Unreleased]`:

```markdown
## [1.3.0] - 2026-06-12

### Added

- **Always-alive loading spinner.** The indeterminate progress spinner
  (`| / - \`) is now animated by a small background runspace, so it keeps
  spinning even while the main thread is blocked waiting for Exchange Online
  to return the first page of results (previously the modal froze until
  results started streaming).

### Changed

- **Mailbox loading uses far less memory.** `Get-Mailbox` has no parameter to
  select properties server-side, so each returned mailbox (200+ properties)
  is now projected down to the 7 properties the app actually uses as soon as
  it streams in, in a single pass. The full objects are no longer buffered or
  kept for the session, and the post-fetch "Filtering dir-synced mailboxes"
  pass is gone.

### Notes

- Not yet manually tested on Windows PowerShell 5.1 for this release; the
  APIs used by the spinner (runspaces, synchronized hashtables, console
  writes) are identical across editions, and CI parse/compatibility checks
  pass on both engines.
```

- [ ] **Step 3: Commit**

```bash
git add SOA-Manager.ps1 CHANGELOG.md
git commit -m "v1.3.0: always-alive load spinner and slim mailbox projection"
```

### Task 8: Push and release

- [ ] **Step 1: Push**

Run: `git push origin main`
Expected: main updated on `mardahl/Exchange-SOA-Manager`.

- [ ] **Step 2: Wait for CI green**

Run: `gh run watch --exit-status` (or `gh run list --limit 1` until completed)
Expected: PSScriptAnalyzer + 5.1/7 parse checks pass. Do NOT release on red.

- [ ] **Step 3: Create release** (gh creates the tag)

```bash
gh release create v1.3.0 --title "v1.3.0" --notes "$(cat <<'EOF'
### Added
- **Always-alive loading spinner.** The indeterminate progress spinner (`| / - \`) is now animated by a background runspace, so it keeps spinning even while the main thread is blocked waiting for Exchange Online to return the first page of results.

### Changed
- **Mailbox loading uses far less memory.** Each mailbox returned by `Get-Mailbox` (200+ properties) is projected down to the 7 properties the app uses as soon as it streams in, in a single pass. Full objects are no longer buffered or kept for the session.

### Notes
- Not yet manually tested on Windows PowerShell 5.1 for this release; the spinner uses runspace APIs that are identical across editions, and CI parse/compatibility checks pass on both engines.
EOF
)"
```

Expected: release URL printed.
