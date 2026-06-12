## Summary

<!-- What does this PR change and why? Link related issues with "Fixes #123". -->

## Type of change

- [ ] Bug fix
- [ ] New feature
- [ ] Documentation
- [ ] Refactor / internal

## Checklist

- [ ] Tested with `.\SOA-Manager.ps1 -Demo` (full UI flow for the affected area)
- [ ] Tested against a real tenant (if the change touches EXO/Graph calls)
- [ ] `Invoke-ScriptAnalyzer` is clean (see CONTRIBUTING.md for the command)
- [ ] No PowerShell 7-only syntax (ternary, `??`, `?.`, `&&`/`||` chains) - the script must parse on Windows PowerShell 5.1
- [ ] README / help overlay updated if keys or behavior changed
- [ ] CHANGELOG.md updated under "Unreleased"

## Screenshots / capture

<!-- For UI changes: paste a terminal capture (tmux capture-pane -pt <session>) or screenshot. -->
