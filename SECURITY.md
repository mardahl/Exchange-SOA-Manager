# Security Policy

## Reporting a vulnerability

Please **do not** open a public issue for security problems. Use GitHub's
[private vulnerability reporting](../../security/advisories/new) on this
repository instead. Reports are looked at on a best-effort basis - this is a
community tool maintained in spare time.

## Scope & threat model

Exchange SOA Manager is an **operator tool**: it runs with the permissions of
the signed-in administrator and performs only the actions the operator
confirms. Relevant notes:

- **No credential handling.** Authentication is delegated entirely to
  `Connect-ExchangeOnline` and `Connect-MgGraph` (MSAL). The tool never sees,
  stores, or logs passwords or tokens.
- **No telemetry.** The only network traffic is to Exchange Online and
  Microsoft Graph endpoints, initiated explicitly by the operator.
- **Local artifacts may be sensitive.** Log files (`SOA-Manager_*.log`),
  backups (`SOA-Backups/*.json`) and exports (`SOA-Exports/*.csv`) contain
  directory data: display names, SMTP addresses, custom attributes. They are
  written next to the script, are `.gitignore`d, and should be treated like
  any other directory export - don't commit them, don't share them unredacted.
- **Module supply chain.** On demand, the tool offers to install
  `ExchangeOnlineManagement` and `Microsoft.Graph.Authentication` from the
  PowerShell Gallery in CurrentUser scope. If your organization requires
  pinned/internal module sources, install the modules yourself beforehand -
  the tool uses whatever is already available.

## Supported versions

Only the latest release on `main` is supported.
