@echo off
rem Removes the "downloaded from the internet" block (Mark of the Web) from
rem every file next to this launcher, then starts SOA-Manager.ps1.
setlocal
rem Remove "rem " from the line below to enable verbose debug logging:
rem set SOA_ARGS=-DebugLog
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%~dp0' -Recurse | Unblock-File"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0SOA-Manager.ps1" %SOA_ARGS% %*
