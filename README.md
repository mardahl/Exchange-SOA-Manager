# Exchange SOA Manager

[![CI](https://github.com/mardahl/Exchange-SOA-Manager/actions/workflows/ci.yml/badge.svg)](https://github.com/mardahl/Exchange-SOA-Manager/actions/workflows/ci.yml)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-5391FE?logo=powershell&logoColor=white)](#requirements)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-555)](#requirements)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![PRs welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)

A portable, single-file PowerShell **terminal UI** for managing **Source of Authority (SOA)** in Exchange Online / Microsoft Entra hybrid environments.

Switch Exchange attribute management between **on-premises** and **cloud** for mailboxes, convert the object-level SOA of **groups** and **contacts**, and control the **tenant-wide default** for new dir-synced mailboxes - all from one keyboard-driven console app. Built for the road toward decommissioning the *Last Exchange Server*.

```
 Exchange SOA Manager  v1.0.0          ● EXO admin@contoso.com   ● Graph admin@contoso.com
  1 Mailboxes   2 Groups   3 Contacts   4 Organization   5 Log
 340 of 340 mailboxes   3 selected   filter:All   sort:Name↑
 sel Name                       Email                              Type           SOA
 [■] Anna Andersen              anna.andersen@contoso.com          UserMailbox    ● Cloud
 [ ] Anna Clausen               anna.clausen@contoso.com           UserMailbox    ● On-prem
 [■] Astrid Dahl                astrid.dahl@contoso.com            UserMailbox    ● Cloud
 ...
 Spc select  A all  N none  / find  F filter  S sort  C to cloud  O to on-prem  E export
```

> **Try it in 10 seconds, no tenant required:** `.\SOA-Manager.ps1 -Demo`

---

- [Why](#why)
- [Features](#features)
- [Quick start](#quick-start)
- [Requirements](#requirements)
- [What the states mean](#what-the-states-mean)
- [Files the tool writes](#files-the-tool-writes)
- [Caveats](#caveats)
- [Contributing](#contributing) · [Security](#security) · [Changelog](CHANGELOG.md)

## Why

Microsoft now supports moving the SOA for Exchange attributes (and whole directory objects) to the cloud, which is the key to finally decommissioning the *Last Exchange Server*. The classic community tooling only covers user mailboxes. This tool covers all four surfaces:

| Surface | Mechanism | Connection |
|---|---|---|
| **Mailboxes** | `Set-Mailbox -IsExchangeCloudManaged $true/$false` | Exchange Online PowerShell |
| **Groups** | `PATCH /groups/{id}/onPremisesSyncBehavior` (`isCloudManaged`) | Microsoft Graph |
| **Contacts** | `PATCH /contacts/{id}/onPremisesSyncBehavior` (`isCloudManaged`) | Microsoft Graph |
| **Organization** | `Set-OrganizationConfig -ExchangeAttributesCloudManagedByDefault` / `-ExchangeAttributesServerManagedByDefault` | Exchange Online PowerShell |

## Features

- **Pure PowerShell TUI** - no WinForms, no WPF, no DLLs, works over SSH and in any VT-capable terminal (Windows Terminal, conhost, iTerm2, ...)
- **Windows PowerShell 5.1 and PowerShell 7+** (PS 7 includes macOS/Linux)
- **Safety first**
  - JSON **attribute backup** written before every conversion (proxy addresses, custom attributes, forwarding, moderation, and more for mailboxes; full directory object for groups/contacts)
  - **Group rollback safety check**: detects cloud-only members that would be lost when sync takes over again, per Microsoft guidance
  - Typed confirmation (`ROLLBACK` / `ENABLE` / `DISABLE`) for hazardous operations
  - Tenant-wide changes show Microsoft's GAL warning before proceeding
- **Batch operations** with live progress, per-item results, and failed items kept selected for retry
- **Search** (`/` live filter), **status filters** (Cloud / On-prem / Pending), **sorting**, **multi-select**
- **CSV export** of any view; **CSV/TXT bulk import** to select objects by identity
- **Pending state tracking** - rolled-back groups/contacts show as `◐ Pending` until the sync client takes over
- **Timestamped log file** plus an in-app log viewer
- **Demo mode** (`-Demo`) with generated data - explore the UI without touching a tenant
- **Lazy connections** - Exchange Online and Graph are only connected when a tab needs them; modules are offered for install (CurrentUser) on demand

## Quick start

```powershell
# Explore the UI with fake data (no tenant access, nothing installed)
.\SOA-Manager.ps1 -Demo

# The real thing
.\SOA-Manager.ps1
```

| Parameter | Effect |
|---|---|
| `-Demo` | Generated sample data, no connections, no changes |
| `-Ascii` | Plain ASCII glyphs for legacy consoles / raster fonts |
| `-NoDisconnect` | Keep the EXO + Graph sessions alive on exit |

### Keys

| Key | Action |
|---|---|
| `Tab` / `1-5` | Switch tab |
| `↑ ↓ PgUp PgDn Home End` | Navigate |
| `Space` | Toggle selection (`A` all, `N` none) |
| `/` | Live search (`Enter` keep, `Esc` clear) |
| `F` | Cycle status filter All → Cloud → On-prem → Pending |
| `S` / `D` | Cycle sort column / flip direction |
| `C` | Convert selection to **cloud-managed** SOA |
| `O` | Convert selection back to **on-prem-managed** SOA |
| `E` / `I` | Export view to CSV / import CSV-TXT selection |
| `R` | Reload from the service |
| `W` | Disconnect sessions |
| `?` | Help · `Q` quit |

## Requirements

### Modules (installed on demand, CurrentUser scope)

- [`ExchangeOnlineManagement`](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) - mailboxes + organization config
- [`Microsoft.Graph.Authentication`](https://www.powershellgallery.com/packages/Microsoft.Graph.Authentication) - groups + contacts (lightweight; the full Graph SDK is **not** required)

### Roles & permissions

| Task | Requirement |
|---|---|
| Mailbox SOA + org config | **Exchange Administrator** (or Global Administrator) |
| Group / contact SOA | **Hybrid Identity Administrator** + admin consent for the Graph scopes below |

Graph scopes requested: `Group.Read.All`, `GroupMember.Read.All`, `User.Read.All`, `Group-OnPremisesSyncBehavior.ReadWrite.All`, `OrgContact.Read.All`, `OrgContact-OnPremisesSyncBehavior.ReadWrite.All`.

> The `*-OnPremisesSyncBehavior` scopes require **admin consent** in the tenant. See [Configure Group SOA](https://learn.microsoft.com/entra/identity/hybrid/how-to-group-source-of-authority-configure#grant-permission-to-apps).

> **Using PIM?** Activate the relevant role **before** connecting - the tool shows the required role on each tab and in the sign-in prompt. If you activate a role *after* connecting, press `W` to disconnect, then reconnect to get a fresh token.

### Sync client versions (Microsoft prerequisites)

- Mailbox attribute SOA: Entra **Connect Sync ≥ 2.5.190.0** (or Cloud Sync)
- Group SOA: Connect Sync **≥ 2.5.76.0** or Cloud Sync **≥ 1.1.1370.0**

## What the states mean

| Badge | Meaning |
|---|---|
| `● On-prem` | Dir-synced, managed on-premises (classic hybrid) |
| `● Cloud` | SOA transferred - editable in EXO / Entra |
| `◐ Pending` | Rolled back to on-prem; waiting for the sync client to take the object over again |

Mailbox conversions only change **Exchange attribute** management - the identity itself stays synced from AD. Group/contact conversions move the **whole object's** SOA. After rolling anything back, run a sync cycle (`Start-ADSyncSyncCycle`) to complete the takeover.

## Files the tool writes

| Location | Content |
|---|---|
| `SOA-Manager_<timestamp>.log` | Session log (also viewable on the Log tab) |
| `SOA-Backups/Backup_<type>_<timestamp>.json` | Pre-conversion attribute snapshots |
| `SOA-Exports/<tab>_<timestamp>.csv` | View exports |

## Caveats

- Do **not** offboard (migrate) a mailbox to on-premises while `IsExchangeCloudManaged` is `$true` - set it back to `$false` first, otherwise sync breaks.
- Rolling back a group/contact completes only after the next Connect/Cloud Sync cycle.
- After updating mailbox attributes on-premises, Microsoft recommends waiting a full sync cycle **plus 24h** before flipping that mailbox to cloud-managed.
- Contacts that were rolled back disappear from the Contacts tab after a refresh until the sync client re-claims them (Graph cannot distinguish them from cloud-born contacts).
- Enabling the tenant-wide cloud default while still provisioning mailboxes on-premises hides those new mailboxes from the EXO GAL (Microsoft warning - shown in-app too).

## References

- [Cloud-based management of Exchange attributes (Microsoft Learn)](https://learn.microsoft.com/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management)
- [Configure Group Source of Authority](https://learn.microsoft.com/entra/identity/hybrid/how-to-group-source-of-authority-configure)
- [Configure User / Contact Source of Authority](https://learn.microsoft.com/entra/identity/hybrid/how-to-user-source-of-authority-configure)

## Contributing

Bug reports and PRs are welcome - see [CONTRIBUTING.md](CONTRIBUTING.md) for
the ground rules (single file, PS 5.1 compatible, safety UX) and a scripted
tmux recipe for testing the TUI. CI enforces PSScriptAnalyzer and parse checks
on both PowerShell engines. Release notes live in [CHANGELOG.md](CHANGELOG.md).

## Security

No telemetry, no credential handling (auth is delegated to MSAL via the
official modules). Logs/backups/exports contain directory data - treat them
accordingly. See [SECURITY.md](SECURITY.md) for the full policy and how to
report vulnerabilities privately.

## Credits

Inspired by [codeandersen's Exchange-SOA-Conversion-tool](https://github.com/codeandersen/Exchange-SOA-Conversion-tool) (WinForms, user mailboxes only). This project is an independent rewrite that adds groups, contacts, the tenant-wide default switch, safety checks, and a portable terminal UI.

## License

MIT - see [LICENSE](LICENSE).

Provided as-is, without warranty. Test in a non-production tenant first.
