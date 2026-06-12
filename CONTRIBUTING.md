# Contributing

Thanks for helping improve Exchange SOA Manager. The project is intentionally a
**single, dependency-free PowerShell script** - please keep it that way.

## Ground rules

1. **One file.** All runtime code lives in `SOA-Manager.ps1`. No companion
   modules, no DLLs, no embedded binaries. Portability is the core feature.
2. **Windows PowerShell 5.1 compatible.** The script must parse and run on
   5.1 *and* 7+. That means **no**:
   - ternary operator (`$x ? $a : $b`)
   - null-coalescing / null-conditional (`??`, `??=`, `?.`, `?[]`)
   - pipeline chain operators (`&&`, `||`)
   - `clean {}` blocks, `` `u{...} `` escapes, `ForEach-Object -Parallel`
   - `$IsWindows` / `$IsMacOS` without a `Get-Variable` guard
3. **StrictMode-safe.** The script runs under `Set-StrictMode -Version 2.0`.
   Use `Get-PropSafe` for properties on API-shaped objects and `$hash['key']`
   indexing for optional hashtable keys.
4. **No telemetry, no phoning home.** The only network calls are to Exchange
   Online and Microsoft Graph, triggered explicitly by the operator.
5. **Safety UX is not optional.** Destructive operations need a confirmation
   modal; tenant-wide or rollback operations need a *typed* confirmation.
   Conversions must write a backup first.

## Dev setup

```powershell
# No tenant needed - demo mode generates data and stubs all writes
.\SOA-Manager.ps1 -Demo
```

### Lint (must be clean before a PR)

```powershell
Install-Module PSScriptAnalyzer -Scope CurrentUser
$settings = @{ Rules = @{ PSUseCompatibleSyntax = @{ Enable = $true; TargetVersions = @('5.1','7.0') } } }
Invoke-ScriptAnalyzer -Path .\SOA-Manager.ps1 -Settings $settings -Severity Error,Warning
```

CI runs the same analyzer plus parse checks on both engines (Windows
PowerShell 5.1 on a Windows runner, PowerShell 7 on Linux).

### Testing the TUI without clicking around

`tmux` makes scripted UI testing easy (macOS/Linux, or Windows via WSL):

```bash
tmux new-session -d -s soa -x 120 -y 32 "pwsh -NoProfile -File ./SOA-Manager.ps1 -Demo"
tmux send-keys -t soa Enter        # connect + load
tmux send-keys -t soa Space c y    # select, convert, confirm
tmux capture-pane -pt soa          # assert on the rendered screen
tmux kill-session -t soa
```

If you change anything UI-related, paste a `capture-pane` snippet (or a
screenshot) into the PR.

### Testing against a real tenant

Use a **test tenant**. Mailbox conversions are reversible
(`IsExchangeCloudManaged $false`), but group/contact rollbacks require a sync
cycle to complete, and member references can be lost if you skip the safety
check - that is exactly the scenario the tool warns about.

## Code map

`SOA-Manager.ps1` is organized in regions - search for `#region`:

| Region | Contents |
|---|---|
| Globals & State | Tabs, theme, glyphs, connection state |
| Logging | `Write-SoaLog` (file + in-app ring buffer) |
| Console / VT engine | Alt-buffer handling, VT enable, `Invoke-OnMainBuffer` |
| Drawing primitives | `Get-PadCell`, `Add-FrameLine`, footer/badges |
| Modals | Message / confirm / typed-confirm / input / report / progress |
| Connections | EXO + Graph connect/disconnect, module install |
| Graph REST helpers | Paged GET, `$batch` sync-behavior lookups |
| Demo data | Generators used by `-Demo` |
| Data fetchers | Mailboxes / groups / contacts / org config |
| Backup, conversion, safety checks | The write paths |
| CSV export / import | View export, bulk selection |
| Conversion pipeline | Targets -> safety check -> confirm -> backup -> convert -> report |
| Views | Per-tab renderers |
| Key dispatch | Global + per-tab key handling |
| Main | Init, event loop, cleanup |

### Conventions that bite if ignored

- **Style strings are fg-only** (`38;5;x`) so the cursor-row background
  carries through cells. Don't embed `ESC[0m` inside row content -
  `Add-FrameLine` appends the reset.
- **Tuple lines** for modals/footers are `@($style, $text)` pairs. Inside a
  multi-element `@( ... )` literal, write `@($style,$text)` **without** a
  leading comma - `,@(...)` wraps the tuple in another array and breaks
  rendering. (Leading commas are only for single-value returns: `return ,$arr`.)
- Widths are computed from plain text *before* styling - pad with
  `Get-PadCell`, then wrap in style codes.

## Pull requests

1. Fork / branch from `main`.
2. Keep PRs focused; one feature or fix per PR.
3. Update `CHANGELOG.md` under **Unreleased**.
4. Fill in the PR template checklist honestly - "not tested" is acceptable
   information, broken `main` is not.
