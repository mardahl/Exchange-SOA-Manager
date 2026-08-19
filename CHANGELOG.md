# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.5.8] - 2026-08-19

### Fixed
- **`PropertyNotFoundException: StdErrLines` crash on worker startup.**
  Follow-up to 1.5.6: the `StdErrLines` queue key was assigned dynamically
  but never added to the `$script:GraphWorker` state hashtable, so the first
  `Stop-GraphWorker` call threw under `Set-StrictMode -Version 2.0`.

## [1.5.7] - 2026-08-19

### Fixed
- **Graph sign-in failed with `A window handle must be configured` on
  Windows.** The worker child process was started with
  `CreateNoWindow = $true`, leaving it without a console window. MSAL's WAM
  broker (enabled by default on Windows) resolves its parent window handle
  from `GetConsoleWindow()`, which returns null without a console, so
  `Connect-MgGraph` threw `AuthenticationFailedException` before sign-in
  could even start. The worker now shares the parent's console
  (`CreateNoWindow = $false`), giving WAM a valid handle.

## [1.5.6] - 2026-08-19

### Fixed
- **Worker startup failures were invisible without `-DebugLog`.** The stderr
  drain only wrote to the debug log, so a Graph worker that died before
  signalling readiness reported the bare `Graph worker exited before
  signalling readiness.` with no cause. Drained stderr lines now also land in
  a thread-safe queue, and the thrown error includes the worker's stderr text
  and process exit code.

## [1.5.5] - 2026-08-19

### Fixed
- **Graph connection crashed with `Cannot find an overload for "Run" and the
  argument count: "2"` on Windows PowerShell 5.1.** The stderr drain used
  `Task.Run(Action, state)`, an overload that does not exist on .NET
  Framework 4.x. Scriptblocks on `Task.Run` also fail because threadpool
  threads have no default runspace. The drain now runs in a dedicated
  `[PowerShell]` runspace, which works identically on 5.1 and 7.

## [1.5.4] - 2026-08-19

### Fixed
- **Graph connection crashed with `PropertyNotFoundException: StdErrTask`.**
  Under `Set-StrictMode -Version 2.0`, reading a missing hashtable key throws.
  `$script:GraphWorker` was initialised without the `StdErrTask` key (added in
  1.5.3), so the first `Stop-GraphWorker` cleanup call inside
  `Start-GraphWorker` failed before the key was ever assigned. The key is now
  part of the initial state hashtable.

## [1.5.3] - 2026-08-18

### Fixed
- **Graph connection failed with `Invalid JSON primitive: WARNING` after
  Exchange module loaded.** After `ExchangeOnlineManagement` was imported in
  the parent process, warning text leaked onto the Graph worker's stdout
  channel, breaking the JSON envelope protocol (`ConvertFrom-Json` failed on
  `WARNING.`). The worker now prefixes every envelope with `SOA::` and sends
  warnings/verbose/info to stderr; the parent drains stderr into the debug
  log and skips non-envelope stdout lines, so worker diagnostics are visible
  and the protocol stays clean.

## [1.5.2] - 2026-07-20

### Fixed
- **Mailbox conversions reported a false `FAILED` with a blank message.**
  `ExchangeOnlineManagement` 3.10.0's `Set-Mailbox -IsExchangeCloudManaged`
  emits an object on success. That output leaked into `Convert-MailboxSoa`'s
  return value, turning the result hashtable into an `Object[]`; the pipeline
  then read `Ok`/`Msg` as empty and logged
  `FAILED converting '<name>' (<id>): ` with no message and no error detail,
  even though the SOA change was applied. The `Set-Mailbox` call is now wrapped
  in `[void](...)` so success output can no longer pollute the return value.

## [1.5.1] - 2026-07-17

### Fixed
- **`-DebugLog` switch had no effect.** In a top-level `.ps1` script the
  `-DebugLog` parameter and `$script:DebugLog` are the same variable;
  `src/00-globals.ps1`'s own default (`$script:DebugLog = $false`, dot-sourced
  after the parameter binds) silently reset it back to false regardless of
  what was passed on the command line, so DEBUG-level logging never
  activated. The switch's value is now captured before the dot-source loop
  runs and applied afterward.

## [1.5.0] - 2026-07-17

### Added
- Verbose error logging: failed conversions now log full exception detail
  (type, message, Graph/EXO response body, inner exceptions, stack trace)
  instead of a single message line.
- Startup log banner reports installed and loaded ExchangeOnlineManagement /
  Microsoft.Graph.Authentication module versions.
- New `-DebugLog` switch (and a commented-out line in
  `Launch-SOA-Manager.bat`) enables per-item conversion tracing (DEBUG log
  level).

## [1.4.0] - 2026-07-17

### Added
- Group forward-conversion audit (`V` on the Groups tab): recursively diffs on-premises AD group membership against Microsoft Entra to flag nested groups and members that would be dropped when converting SOA to cloud. Green/yellow indicator in a new `Chk` column; full findings written to `SOA-Exports/GroupAudit_<timestamp>.csv`. Windows-only (uses the RSAT ActiveDirectory module, offered for install on demand).

### Changed
- `SOA-Manager.ps1` is now a thin bootstrap that dot-sources numbered region files under `src/`; releases bundle the `src/` folder alongside the script.

## [1.3.8] - 2026-07-16

### Fixed

- **Mailbox conversions no longer report false failures.** `Set-Mailbox
  -IsExchangeCloudManaged` has started throwing a terminating error whose
  `Exception.Message` is empty even when the SOA change is applied
  server-side, causing every mailbox conversion to log
  `FAILED converting '<name>' (<id>): ` with no message while the change
  actually took effect (visible after a reload). `Convert-MailboxSoa` now
  reads `ErrorDetails.Message` as well, and on a caught error re-reads the
  mailbox to confirm the real state before reporting failure.

## [1.3.7] - 2026-07-15

### Fixed

- **Microsoft Graph is now isolated in a child PowerShell process.**
  `ExchangeOnlineManagement` and `Microsoft.Graph.Authentication` ship
  incompatible versions of `Microsoft.Identity.Client.dll`. The v1.3.5
  workaround of signing into Graph before Exchange Online started failing again
  on newer module releases with `MissingMethodException` on
  `BrokerExtension.WithBrowser`. The tool now keeps Exchange Online as the only
  MSAL consumer in the main process and runs all Graph authentication and REST
  calls in a dedicated child `pwsh`/`powershell` process that communicates over
  stdin/stdout JSON envelopes
  ([microsoftgraph/msgraph-sdk-powershell #3331](https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3331)).

## [1.3.6] - 2026-07-15

### Added

- **Packaged release zip and launcher.** Releases now ship a curated
  `Exchange-SOA-Manager-<tag>.zip` (script, `Launch-SOA-Manager.bat`, README,
  LICENSE, CHANGELOG) instead of GitHub's raw source archive.
  `Launch-SOA-Manager.bat` removes the Mark of the Web from the extracted
  files (`Unblock-File`) and starts the script, so there's no PowerShell
  execution-policy prompt after downloading and extracting the zip.

## [1.3.5] - 2026-07-15

### Fixed

- **MSAL conflict workaround rewritten - Exchange Online now signs in reliably.**
  The v1.3.3/v1.3.4 approach of hand-loading Graph's MSAL assemblies via
  `Assembly.LoadFrom` never worked across environments: the dependency
  `Microsoft.IdentityModel.Abstractions` (v8.x) still failed to resolve
  (`Could not load file or assembly 'Microsoft.IdentityModel.Abstractions,
  Version=8.14.0.0'`) on PowerShell 7, and Windows PowerShell 5.1 looped the
  interactive browser sign-in and hung. The tool now uses the documented,
  reliable workaround: it runs a full `Connect-MgGraph` **before**
  `Connect-ExchangeOnline`, so Graph's own module loader initialises MSAL with
  its complete dependency closure and Exchange Online reuses it
  ([msgraph-sdk-powershell #3394](https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3394)).

### Changed

- **Exchange Online now requires a Microsoft Graph sign-in first.** To avoid the
  MSAL assembly clash, connecting to Exchange Online (Mailboxes / Organization)
  now signs in to Microsoft Graph first. This replaces the previous
  "mailbox-first, no up-front Graph login" behaviour.

## [1.3.4] - 2026-07-15

### Fixed

- **Exchange Online sign-in looped/failed after the v1.3.3 MSAL workaround.**
  The `#3394` pre-load loaded only `Microsoft.Identity.Client.dll` (v4.82) but
  not its required dependency `Microsoft.IdentityModel.Abstractions.dll` (v8.x) -
  the assembly that owns the `WithLogging` / `IIdentityLogger` API at the heart
  of the conflict. With the dependency unresolved, the first real MSAL call
  failed: on PowerShell 7 with `Could not load file or assembly
  'Microsoft.IdentityModel.Abstractions, Version=8.14.0.0'`, and on Windows
  PowerShell 5.1 the interactive browser sign-in reloaded repeatedly and then
  hung forever. The pre-load now loads the sibling
  `Microsoft.IdentityModel.Abstractions.dll` (same edition folder) **before**
  MSAL, so the dependency resolves and both Exchange Online and a later
  `Connect-MgGraph` complete
  ([msgraph-sdk-powershell #3394](https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3394)).

## [1.3.3] - 2026-07-09

### Fixed

- **Graph sign-in still failed after Exchange Online in v1.3.2.** The MSAL
  conflict workaround only ran `Import-Module Microsoft.Graph.Authentication`
  before Exchange, but importing the module does **not** load the MSAL assembly
  (`Microsoft.Identity.Client`) - that only happens when a module actually
  authenticates. So `Connect-ExchangeOnline` was still the first call to touch
  MSAL and pinned the older version, breaking a later `Connect-MgGraph`
  ([msgraph-sdk-powershell #3394](https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3394)).
  The workaround now explicitly loads Graph's newer `Microsoft.Identity.Client.dll`
  into the process (via `Assembly.LoadFrom`) before Exchange authenticates, so
  the newer MSAL is pinned without forcing an up-front Graph sign-in - keeping
  the mailbox-first workflow intact.

## [1.3.1] - 2026-07-08

### Fixed

- **Mailbox load is responsive again.** EXO v3 REST cmdlets buffer the entire
  result set before releasing anything to the pipeline, so the progress modal
  froze (no updates, no Esc) for the whole fetch. `Get-Mailbox` now runs in a
  background runspace (reusing the process-wide EXO connection) while the main
  thread repaints the modal and polls for keys, so:
  - elapsed time ticks live during the fetch
  - Esc cancels the fetch mid-flight
  - the modal explains the fetch can take several minutes instead of showing
    a mailbox count that never updates
  If the worker runspace cannot see the EXO connection, the previous blocking
  fetch is used as a fallback.

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

## [1.2.3] - 2026-06-12

### Fixed

- **Graph sign-in failed after connecting to Exchange Online.** `Connect-MgGraph`
  threw `Method not found: '... WithLogging(Microsoft.IdentityModel.Abstractions.IIdentityLogger, Boolean)'`
  whenever Exchange Online was connected first. This is a known, unresolved
  conflict between the ExchangeOnlineManagement and Microsoft.Graph.Authentication
  modules ([msgraph-sdk-powershell #3394](https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3394)):
  both bundle different versions of the MSAL assemblies, and whichever module
  loads first wins for the whole process. The script now imports
  `Microsoft.Graph.Authentication` before `ExchangeOnlineManagement` so the
  newer MSAL assemblies are pinned first. If the conflict still occurs (e.g. the
  Graph module was installed mid-session), a dedicated error dialog now explains
  the cause and tells the user to restart the tool and/or update both modules.

## [1.2.2] - 2026-06-12

### Fixed

- **Get-EXOMailbox failed because of IsExchangeCloudManaged.** The script was
  requesting `IsExchangeCloudManaged` as a custom property from
  `Get-EXOMailbox`, which is unsupported by that modern REST cmdlet. It now
  uses a highly optimized server-side filtered `Get-Mailbox` query to load only
  directory-synced accounts directly, bypassing the error and preventing timeouts
  or throttling on larger tenants.
- **Improved UI messaging.** Connection panels on all list tabs now explicitly state
  that only directory-synced objects will be loaded.

## [1.2.1] - 2026-06-12

### Fixed

- **Loading crashed in connected sessions.** All three list tabs failed with
  `The term 'Write-ProgressModal' is not recognized` (or
  `'Format-ElapsedTime' is not recognized`) as soon as the first results
  streamed in. The progress callbacks are created with `GetNewClosure()`,
  which binds them to a dynamic module that cannot resolve script-scope
  functions by name when the script runs in its own script scope (the normal
  `./SOA-Manager.ps1` launch). Function references are now captured into
  variables and invoked through them. Demo mode never hit the callbacks,
  which is why this slipped through.
- **`Get-EXOMailbox` failed and fell back to `Get-Mailbox`.** The v3
  ExchangeOnlineManagement module removed the `-PageSize` parameter, so every
  mailbox load logged
  `Get-EXOMailbox failed (A parameter cannot be found that matches parameter
  name 'PageSize'.)` and used the slower fallback. `-PageSize` is gone;
  results still stream into the progress modal page by page (1000 per page,
  the service default).
- **Contact SOA Graph scope was wrong.** The sign-in requested the
  non-existent `OrgContact-OnPremisesSyncBehavior.ReadWrite.All` scope, which
  would fail the entire Microsoft Graph connection. Corrected to the
  documented `Contacts-OnPremisesSyncBehavior.ReadWrite.All`.

## [1.2.0] - 2026-06-12

### Added

- **Live progress while loading.** Mailboxes now stream from Exchange Online
  in pages of 250, and the loading modal shows a running count and elapsed
  time as results arrive. Graph group/contact pages and the batched SOA-state
  lookups report the same way. Previously the modal sat at a static 0% until
  the entire result set had arrived, which looked like a hang on larger
  tenants.
- **Esc cancels loading.** A load in progress can be aborted with Esc; the
  tab stays unloaded and can be reloaded with `R` or `Enter`.

### Changed

- Loads with an unknown total render an animated marquee bar with a spinner
  instead of a static 0% percent bar.

## [1.1.0] - 2026-06-12

### Added

- **Role guidance for PIM users.** The connect panels on every tab, the
  sign-in prompts, and the help overlay now state which directory role is
  required (Exchange Administrator for mailbox SOA; Hybrid Identity
  Administrator for group/contact SOA; Hybrid Identity or Global
  Administrator for the tenant-wide default) and remind PIM users to activate
  the role *before* signing in - plus how to refresh the token (`W` then
  reconnect) when a role was activated late.

## [1.0.0] - 2026-06-12

### Added

- **Terminal UI** (pure PowerShell VT/ANSI, no binary dependencies) with five
  tabs: Mailboxes, Groups, Contacts, Organization, Log.
- **Mailbox SOA conversion** - toggles `IsExchangeCloudManaged` via
  `Set-Mailbox` for dir-synced mailboxes (Exchange attribute SOA).
- **Group SOA conversion** - toggles `isCloudManaged` via the Microsoft Graph
  `onPremisesSyncBehavior` API, with batched (`$batch`) status detection of
  already-converted groups.
- **Contact SOA conversion** - same Graph mechanism for organizational
  contacts.
- **Tenant-wide default switch** - `Set-OrganizationConfig
  -ExchangeAttributesCloudManagedByDefault` / `-ExchangeAttributesServerManagedByDefault`
  with typed confirmation and Microsoft's GAL warning surfaced in-app.
- **Attribute backups** - JSON snapshot written before every conversion
  (mailbox attribute set incl. proxy addresses and custom attributes 1-15;
  full directory object for groups/contacts).
- **Group rollback safety check** - detects cloud-only members before
  converting a group back to on-premises, per Microsoft guidance; requires a
  typed `ROLLBACK` confirmation when hazards are found.
- **Pending state tracking** - rolled-back objects show as `Pending` until the
  sync client takes them over again.
- **List tooling** - live search, status filter cycling, column sorting,
  multi-select, select-all/none, batch progress with per-item results; failed
  items stay selected for retry.
- **CSV export** of the current view and **CSV/TXT bulk import** that selects
  matching objects.
- **Demo mode** (`-Demo`) with generated data - full UI without a tenant.
- **Lazy connections** - EXO and Graph connect only when needed; missing
  modules are offered for install in CurrentUser scope.
- **Logging** - timestamped log file plus an in-app scrollable log viewer.
- `-Ascii` glyph fallback and `-NoDisconnect` session persistence switches.
- CI: PSScriptAnalyzer (with 5.1/7.0 syntax compatibility rules) and parse
  checks on both Windows PowerShell 5.1 and PowerShell 7.

[Unreleased]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.5.8...HEAD
[1.5.8]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.5.7...v1.5.8
[1.5.7]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.5.6...v1.5.7
[1.5.6]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.5.5...v1.5.6
[1.5.5]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.5.4...v1.5.5
[1.5.4]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.5.3...v1.5.4
[1.3.6]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.3.5...v1.3.6
[1.3.1]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.3.0...v1.3.1
[1.3.0]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.2.3...v1.3.0
[1.2.3]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.2.2...v1.2.3
[1.2.2]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.2.1...v1.2.2
[1.2.1]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/mardahl/Exchange-SOA-Manager/releases/tag/v1.0.0
