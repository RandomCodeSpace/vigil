# Install & run

Vigil is distributed as PowerShell source. There is no installer, no compiled binary, no auto-updater.

## Requirements

| Component | Required | Notes |
|---|---|---|
| OS | Windows 10 / 11 (full UI) | Cross-platform pwsh on Linux / macOS works for the data layer + tests, no UI |
| PowerShell | 7.5+ (`pwsh`) | Bundles .NET 9 |
| .NET runtime | 9 | Bundled with PowerShell 7.5 |
| Microsoft Outlook | Optional | Required only for the 15-minute Outlook COM sync (flagged emails / calendar / tasks) |
| Disk | < 1 MB | Source-only; `.vigil/` stores at most a few KB of task data |

PowerShell 5.1 is supported as a legacy fallback for the data layer only — the WPF Fluent theme degrades and parts of the UI are gated behind `if ($PSVersionTable.PSVersion.Major -ge 7)`.

## Get the source

```powershell
git clone https://github.com/RandomCodeSpace/vigil.git
cd vigil
```

To pin to a specific revision:

```powershell
git checkout <commit-sha>
```

`SECURITY.md` asks vulnerability reporters to include `git rev-parse HEAD` so the affected version is unambiguous; the commit SHA is the canonical version identifier until a tagged release line exists.

## Run

```powershell
pwsh -ExecutionPolicy Bypass -File .\VIGIL.ps1
```

First launch:

1. Runs `preflight.ps1` to emit the environment bitmap.
2. Creates `.vigil/` next to `VIGIL.ps1` (fallback `~/.vigil` if the script directory is not writable; legacy `%USERPROFILE%\.vigil` is migrated).
3. Installs a Desktop shortcut and registers an auto-start entry.
4. Starts the chromeless WPF window with system-tray integration.

### Flags

| Flag | Behaviour |
|---|---|
| `-NoUI` | Headless mode — no WPF window, no tray icon. Used by the test harness. |
| `-IncludeCalendar` | Include the next 24 hours of Outlook calendar items in the task list. |

Combine flags as needed: `pwsh -File .\VIGIL.ps1 -NoUI -IncludeCalendar`.

## Tests

```powershell
pwsh -NoProfile -File .\Test-Vigil.ps1
```

Test-Vigil.ps1 dot-sources VIGIL.ps1 with `-NoUI`; runs on PowerShell 5.1 and pwsh 6+ across Linux / macOS / Windows. Approximately 116 unit tests cover the data layer (task model, store path resolution + legacy migration, atomic writes, DPAPI key handling on Windows, Outlook sort-before-restrict invariant, RCW lifecycle hygiene, log rotation).

## Update

```powershell
git pull --ff-only origin main
```

Vigil does not phone home; security fixes land on `main` and `git pull` is the patch channel. See [`SECURITY.md`](../SECURITY.md) §Supported versions.

## Uninstall

```powershell
# Stop the running instance (close from tray) then:
Remove-Item -Recurse .\.vigil      # task store + logs
# Remove the Desktop shortcut and auto-start entry by hand if desired.
```

## See also

- [`docs/architecture.md`](architecture.md) — how the app is structured.
- [`docs/troubleshooting.md`](troubleshooting.md) — common environment blockers.
- [`docs/security.md`](security.md) — security model + threat model.
