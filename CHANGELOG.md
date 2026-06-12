# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.2.3...HEAD
[1.2.3]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.2.2...v1.2.3
[1.2.2]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.2.1...v1.2.2
[1.2.1]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/mardahl/Exchange-SOA-Manager/releases/tag/v1.0.0
