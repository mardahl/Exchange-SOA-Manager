# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/mardahl/Exchange-SOA-Manager/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/mardahl/Exchange-SOA-Manager/releases/tag/v1.0.0
