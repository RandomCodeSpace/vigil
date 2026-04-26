# VIGIL

[![OpenSSF Best Practices](https://www.bestpractices.dev/projects/12648/badge)](https://www.bestpractices.dev/projects/12648)
[![OpenSSF Scorecard](https://api.scorecard.dev/projects/github.com/RandomCodeSpace/vigil/badge)](https://scorecard.dev/viewer/?uri=github.com/RandomCodeSpace/vigil)

Personal task command center for Windows. Single-file PowerShell + WPF app with
a chromeless Fluent window, Outlook sync, and a system tray icon.

## Features

- Chromeless WPF window with Fluent theme and Mica backdrop (Windows 11)
- Tabs for Tasks and Calendar, always-on-top toggle, compact mode
- Outlook sync: 24h calendar items, flagged emails, and open Outlook tasks
- Clickable top-bar badges (CAL / TASK / CRIT / HIGH) as filters
- DPAPI-encrypted local storage in `.vigil/` next to the script
- System tray icon with quick actions
- Desktop shortcut auto-installed on first run
- 116 cross-platform unit tests for the data layer

## Requirements

- Windows 10/11
- PowerShell 7.5+ (`pwsh`) with .NET 9
- Microsoft Outlook (only for sync)

## Run

```powershell
pwsh -ExecutionPolicy Bypass -File .\VIGIL.ps1
```

First launch creates `.vigil/` alongside `VIGIL.ps1`, installs a Desktop
shortcut, and registers an auto-start entry.

## Tests

```powershell
pwsh -NoProfile -File .\Test-Vigil.ps1
```

Tests run cross-platform (Linux / macOS pwsh) against the `-NoUI` data layer.

## License

[MIT](LICENSE)
