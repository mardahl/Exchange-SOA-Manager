# Contributing

Thanks for helping improve Exchange SOA Manager. The project is a
**dependency-free PowerShell TUI**: a thin `SOA-Manager.ps1` bootstrap plus
one file per region under `src/` - please keep that shape.

## Ground rules

1. **Thin bootstrap, one file per region.** `SOA-Manager.ps1` is a small
   bootstrap (help header, `param()`, StrictMode) that dot-sources every
   `src/*.ps1` at script scope, then runs the Main event loop. All other
   runtime code lives in `src/`, one file per `#region`, numbered so the
   file-name sort order is the load order. No companion modules on the load
   path other than these dot-sourced files, no DLLs, no embedded binaries, no
   external module dependencies. Portability is still the core feature — the
   launcher `.bat` unblocks the whole folder recursively.
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
Invoke-ScriptAnalyzer -Path .\SOA-Manager.ps1, .\src -Recurse -Settings .\PSScriptAnalyzerSettings.psd1 -Severity Error, Warning
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

The code is split into a bootstrap plus one file per region under `src/`:

| File | Contents |
|---|---|
| `SOA-Manager.ps1` | Bootstrap: help header, `param()`, StrictMode, dot-source loop, Main event loop |
| `src/00-globals.ps1` | Tabs, theme, glyphs, connection state |
| `src/05-logging.ps1` | `Write-SoaLog` (file + in-app ring buffer) |
| `src/10-console-vt.ps1` | Alt-buffer handling, VT enable, `Invoke-OnMainBuffer` |
| `src/15-drawing.ps1` | `Get-PadCell`, `Add-FrameLine`, footer/badges |
| `src/20-modals.ps1` | Message / confirm / typed-confirm / input / report / progress |
| `src/25-data-helpers.ps1` | Shared data-shaping helpers (`Get-PropSafe`, etc.) |
| `src/30-connections.ps1` | EXO + Graph connect/disconnect, module install |
| `src/35-graph-rest.ps1` | Paged GET, `$batch` sync-behavior lookups |
| `src/40-demo-data.ps1` | Generators used by `-Demo` |
| `src/45-data-fetchers.ps1` | Mailboxes / groups / contacts / org config |
| `src/50-backup-conversion.ps1` | Backup + conversion + safety-check write paths |
| `src/55-csv.ps1` | CSV export of the current view, CSV/TXT bulk import |
| `src/60-conversion-pipeline.ps1` | Targets -> safety check -> confirm -> backup -> convert -> report |
| `src/65-views.ps1` | Per-tab renderers |
| `src/70-org-actions.ps1` | Organization-wide default switch |
| `src/75-key-dispatch.ps1` | Global + per-tab key handling |

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
