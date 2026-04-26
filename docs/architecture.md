# Architecture

Vigil is a single-file PowerShell + WPF desktop app. Three scripts compose the project; there is no compile step.

## Repository shape

| File | Purpose |
|---|---|
| [`VIGIL.ps1`](../VIGIL.ps1) | Main app (≈99 KB). 5-phase startup, WPF window, system-tray integration, Outlook COM sync, persistent task store. Parameters: `-NoUI`, `-IncludeCalendar`. |
| [`preflight.ps1`](../preflight.ps1) | 60 environment checks (runtime, corp lockdown CLM / AppLocker / AMSI, Azure AD, hardware). Emits the compact bitmap `VIGIL:v1:<count>:<hex>:P<n>:F<n>:T<tenant>` consumed by the main app. |
| [`Test-Vigil.ps1`](../Test-Vigil.ps1) | Dot-sources `VIGIL.ps1` with `-NoUI`; custom asserts; runs on PowerShell 5.1 and pwsh 6+ across Linux / macOS / Windows. ≈116 cross-platform unit tests for the data layer. |
| `.vigil/` | Runtime state (DPAPI-wrapped `tasks.json`, logs). Colocated next to the script; fallback `~/.vigil`; legacy migration from the old userprofile path. |

## 5-phase startup (`VIGIL.ps1`)

```
1. preflight  → environment checks, emits status bitmap
2. quick-add  → optional fast-path popup if invoked with input
3. Outlook    → 15-minute COM sync (flagged emails / 24h calendar / open tasks)
                runs on a separate runspace; never blocks the dispatcher
4. shortcut   → install desktop / startup shortcut on first run
5. UI + tray  → chromeless WPF window with Fluent / Mica + system tray icon
```

Search input is debounced; the debounce is flushed on window close so the last text is never lost.

## Runtime invariants

These are enforced by the test harness and reviewed for at PR time:

1. **Single-instance mutex** — `Global\VIGIL_TaskTracker` named mutex + P/Invoke window reactivation. Do not introduce a second process model.
2. **DPAPI-wrapped `tasks.json`** — CurrentUser scope, only when BitLocker is off. Per-user OS-managed master key; no key material is exfiltrated.
3. **Atomic writes** — via `[System.IO.File]::Replace` exclusively, never `Move-Item -Force`.
4. **Outlook COM** — `Sort()` **before** `Restrict()` for correct flagged counts; reverse-order `ReleaseComObject` + forced GC after each session to prevent RCW leaks.
5. **Log rotation** — 500-line cap.
6. **Cross-platform core** — anything that is not WPF / Outlook COM must run on Linux / macOS pwsh; the test suite enforces this.
7. **`System.Drawing` is lazy-loaded** — Linux pwsh lacks `libgdiplus`; do not import at script top.
8. **Path canonicalisation** — task-store path is canonicalised via `[System.IO.Path]::GetFullPath` and asserted to live under `.vigil/` or `~/.vigil`.
9. **Reduce-motion** — variant strips Storyboards entirely; honour the user's preference.

## Stack

| Layer | Choice |
|---|---|
| Primary language | PowerShell 7.5+ (`pwsh`) — cross-platform core |
| Legacy fallback | Windows PowerShell 5.1 — data layer only; Fluent theme degrades |
| Runtime | .NET 9 |
| UI | WPF / Fluent / Mica (Windows-only) |
| Embedded C# | `Add-Type` inline P/Invoke for window reactivation + tray hotkeys |
| External integration | Outlook COM (15-minute sync) |
| Storage | DPAPI-wrapped `tasks.json` (CurrentUser scope) |
| Tests | Custom asserts, dot-sources VIGIL with `-NoUI` |
| Package manager | None — Vigil ships no `package-lock.json`, `pom.xml`, or `requirements.txt`. Dependencies are framework-bundled (.NET / pwsh built-ins). |

## See also

- [`CLAUDE.md`](../CLAUDE.md) §1–§9 — full conventions and quality gates.
- [`docs/install.md`](install.md) — install + run.
- [`docs/troubleshooting.md`](troubleshooting.md) — common environment issues.
