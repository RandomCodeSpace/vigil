# Troubleshooting

Common environment blockers and how to diagnose them.

## Preflight bitmap

`VIGIL.ps1` runs `preflight.ps1` first and emits a compact status bitmap:

```
VIGIL:v1:<count>:<hex>:P<n>:F<n>:T<tenant>
```

| Field | Meaning |
|---|---|
| `v1` | Bitmap schema version |
| `<count>` | Total number of preflight checks executed |
| `<hex>` | Bit-mask of pass/fail across the check matrix |
| `P<n>` | Number of passing checks |
| `F<n>` | Number of failing checks |
| `T<tenant>` | Azure AD tenant signature (truncated SHA prefix) |

If `F<n>` is non-zero, dump the full preflight report:

```powershell
pwsh -File .\preflight.ps1 -Verbose
```

## Constrained Language Mode (CLM) — hard block

Symptom: `VIGIL.ps1` exits immediately during preflight phase 1 with an error mentioning `LanguageMode`.

Cause: corporate AppLocker / WDAC policy has put PowerShell into Constrained Language Mode. Vigil's inline `Add-Type` C# (P/Invoke for window reactivation + tray hotkeys) cannot run.

Fix: there is no in-script workaround — CLM is enforced at the Windows policy layer. Talk to your endpoint-management team about an exception or run on a different machine.

## AppLocker / AMSI / EDR may block `Add-Type`

Symptom: `Add-Type` calls fail with `Cannot add type. The compiler errors are: ...` even outside CLM.

Cause: AMSI hook in the EDR is intercepting the inline C# compile.

Fix: the inline C# is identifiable to your EDR vendor; an allow-list rule on the Vigil script path is the typical resolution. Vigil cannot work around an EDR block.

## PowerShell 5.1 — Fluent theme degrades

Symptom: the WPF window renders but without Mica backdrop or the Fluent accent palette.

Cause: PowerShell 5.1 ships with an older WPF surface that does not expose the `Microsoft.UI.Xaml`-derived styles. Vigil gates the Fluent theme behind `if ($PSVersionTable.PSVersion.Major -ge 7)`.

Fix: install PowerShell 7.5+ from https://github.com/PowerShell/PowerShell/releases and re-run via `pwsh -File .\VIGIL.ps1`.

## Outlook COM authentication prompt

Symptom: a Microsoft Authentication / "Allow Vigil to access Outlook" dialog appears every 15 minutes.

Cause: the Outlook security model is prompting per-session. Common with first-time use, after a tenant policy refresh, or with strict Outlook trust settings.

Fix: in Outlook, go to **File → Options → Trust Center → Programmatic Access** and review the antivirus / programmatic access settings. The COM Add-In list is the right surface; specifics are tenant-dependent.

## Outlook flagged-count mismatch

Symptom: the top-bar `TASK` badge shows fewer items than Outlook's flagged folder.

Cause: regression — Outlook COM `Sort()` was called **after** `Restrict()` instead of before. This is a known invariant: the data-layer test in `Test-Vigil.ps1` covers it.

Fix: this should never reach `main` — branch protection blocks unsigned / untested commits, and the test asserts `Sort()` precedes `Restrict()`. If you see it, file an issue with the commit SHA.

## DPAPI store cannot be opened

Symptom: `tasks.json` exists under `.vigil/` but `VIGIL.ps1` reports "Cannot decrypt task store".

Cause: typically one of:

1. The OS user profile has changed (DPAPI is per-user; copying `.vigil/` between users / machines does not migrate the master key).
2. BitLocker has been disabled and the DPAPI scope assumption no longer holds.
3. `tasks.json` was edited by hand (DPAPI integrity check fails).

Fix: rename `.vigil/tasks.json` to `tasks.json.broken`, restart Vigil, and a fresh empty store will be created. Open the broken file in a hex viewer to recover entries if needed; DPAPI cannot be brute-forced offline.

## Logs

Vigil rotates its own logs — 500 lines per file, oldest dropped first. Path:

```
.\\.vigil\\vigil.log
```

Tail with:

```powershell
Get-Content .\.vigil\vigil.log -Tail 80 -Wait
```

## Single-instance mutex collision

Symptom: launching `VIGIL.ps1` does nothing visible — no window appears.

Cause: an existing instance owns the `Global\VIGIL_TaskTracker` named mutex. The launcher reactivates the existing window via P/Invoke instead of starting a second process.

Fix: this is by design. If the existing instance is hung, kill it via Task Manager (look for `pwsh.exe` running `VIGIL.ps1`) and relaunch.

## See also

- [`docs/architecture.md`](architecture.md) — runtime invariants.
- [`SECURITY.md`](../SECURITY.md) — vulnerability disclosure (do not file security issues as public GitHub issues).
- [`preflight.ps1`](../preflight.ps1) — full environment-check source.
