# Mailbox fetch optimization and always-alive load spinner

Date: 2026-06-12
Status: approved (design), pending implementation
Target: `SOA-Manager.ps1` (single-file TUI, Windows PowerShell 5.1 and PowerShell 7+)

## Problem

1. `Get-Mailbox` does not support server-side property selection (no
   `-Properties` parameter, verified against Microsoft Learn). `Get-MailboxItems`
   currently buffers every full mailbox object (200+ properties) in an
   ArrayList, iterates the buffer a second time to filter and project, and
   keeps the full object alive for the whole session via `Raw`. In large
   tenants this wastes a lot of memory and adds a post-fetch pass.
2. The indeterminate progress modal already shows a `\ | / -` spinner, but it
   only advances when the streaming progress callback fires. While the main
   thread is blocked inside the `Get-Mailbox` pipeline (initial wait, gaps
   between pages) the whole modal freezes.

## Verified facts the design rests on

- `Get-Mailbox` has no `-Properties`/`-PropertySets` parameter in any
  parameter set (Microsoft Learn, Get-Mailbox reference). Client-side
  projection is the only available data reduction.
- Mailbox `Raw` is dead weight: conversion uses only `$Item.Id`
  (`Convert-MailboxSoa`), and backups re-fetch fresh objects via
  `Get-MailboxBackupRecord` / `Get-Mailbox -Identity`. Only Graph objects
  (groups, contacts) use `Raw` for backups.
- .NET console streams are synchronized: concurrent `[Console]::Write` calls
  from multiple threads do not tear (System.Console docs). Both writers emit
  complete, absolutely positioned ANSI sequences, so no extra lock is needed.
- A background PowerShell runspace can paint the spinner while the main
  thread is blocked. Proven with a local proof-of-concept (PS 7): spinner
  animated during a blocked main thread, full repaints interleaved cleanly,
  clean stop via `EndInvoke`. The runspace never touches Exchange, so the
  "EXO connections are runspace-bound" concern does not apply.

## Part 1: single-pass slim projection in Get-MailboxItems

Replace the buffer-then-filter flow with one streaming pass:

- As each mailbox streams out of `Get-Mailbox`, immediately project it to a
  slim `pscustomobject` holding only the seven properties the app uses:
  `DisplayName`, `UserPrincipalName`, `PrimarySmtpAddress`,
  `RecipientTypeDetails`, `IsDirSynced`, `IsExchangeCloudManaged`,
  `ExternalDirectoryObjectId`. The full object is discarded and can be
  garbage-collected.
- Skip items where `IsDirSynced` is false inside the same pass (cheap no-op
  for the filtered query; required for the unfiltered fallback path).
- Build the list item (`Type/Id/Name/Email/Detail/Soa/Selected/Raw`) directly
  in the stream. `Raw` becomes the slim record - same shape demo mode already
  uses (`New-DemoMailboxes`).
- Keep both query paths: filtered `-Filter "IsDirSynced -eq $true"` first,
  unfiltered fallback on failure (behavior unchanged).
- Progress callback semantics unchanged: `(count, label)` with count =
  mailboxes received so far. Track received vs kept separately so the final
  log line ("Retrieved X mailboxes; Y are dir-synced") stays accurate.
- The intermediate "Filtering dir-synced mailboxes..." phase disappears
  because there is no second pass.

## Part 2: background-runspace spinner

New script-scope helper pair:

- `Start-LoadSpinner`: creates `[hashtable]::Synchronized(@{ X = 0; Y = 0;
  Run = $true; Style = ... })`, creates `[powershell]::Create()`, adds the
  spinner scriptblock with the state as argument, `BeginInvoke()`. Stores
  powershell instance, async handle, and state in `$script:Spinner`.
- `Stop-LoadSpinner`: sets `Run = $false`, calls `EndInvoke` (joins, so no
  stray writes after return), disposes the powershell instance, clears
  `$script:Spinner`.

Spinner runspace loop:

- Every ~120 ms: if `X -gt 0 -and Y -gt 0`, write one frame
  (`Style + char + Reset`) at `ESC[Y;XH`. Frame index =
  `[int](TickCount / 120) % 4` over `| / - \` - the same formula
  `Write-ProgressModal` already uses, so background ticks and full repaints
  stay in phase.
- `X = 0` means hidden (paused without stopping the runspace).

Integration:

- `Write-ProgressModal`, indeterminate branch (`Total -le 0`): after
  computing box geometry, publish the spinner cell coordinates to the state
  (if a spinner is active). The spinner cell is the suffix slot after the
  marquee bar: row = box top + 3, col = box left + 2 + barW + 3.
- `Write-ProgressModal`, determinate branch: publish `X = 0` (hidden) so the
  conversion progress modal is never painted over.
- `Invoke-TabLoad`: start the spinner right after the initial
  `Write-ProgressModal`, stop it in a `finally` around the fetch/`switch`
  block so cancellation (Esc) and errors also clean up.
- Esc handling is unchanged: it still only responds while results stream
  (main thread remains blocked in the pipeline between callbacks).

## Error handling

- Spinner start failure (runspace creation) must not break loading: wrap in
  try/catch, log WARN, continue without spinner.
- `Stop-LoadSpinner` must be idempotent and safe when no spinner is running.
- Console resize during load: coordinates refresh on the next progress
  repaint; a brief stale-position frame is acceptable (the modal itself has
  the same limitation today).

## Testing and verification

- CI: PSScriptAnalyzer (5.1/7.0 compatibility rules) and parse checks on both
  engines must stay green. No PS 7-only syntax (per PR template).
- Local: demo mode run (`-Demo`) to confirm the TUI still works; demo loads
  return instantly so the spinner is best verified with the runspace
  proof-of-concept pattern plus a connected-mode run.
- Manual: connected-mode load against a real tenant (user-side), confirming
  the spinner animates during the initial wait and the mailbox list is
  unchanged.
- Known gap: not yet manually tested on Windows PowerShell 5.1. All APIs used
  (`[powershell]::Create()`, synchronized hashtable, `[Console]::Write`,
  ANSI positioning already used throughout the script) behave identically on
  5.1; risk is low but nonzero. Note this in the CHANGELOG entry.

## Out of scope

- Making Esc responsive while the pipeline is blocked (would require moving
  the fetch itself off the main thread; rejected - EXO connection state is
  bound to the connecting runspace and untestable from this machine).
- Wire-level payload reduction for mailboxes (impossible with `Get-Mailbox`;
  `Get-EXOMailbox` cannot be used because it does not expose
  `IsExchangeCloudManaged`).
- Group/contact fetch paths (already use Graph `$select` projection).
