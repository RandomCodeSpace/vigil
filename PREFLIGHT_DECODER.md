# VIGIL Preflight Result String — Decoder Reference

**Current schema: v2 (60 checks).** v1 (35 checks) is frozen.

When you run `preflight.ps1` on the corp machine, the last line looks like:

```
VIGIL:v2:55:7FFFFFFFFFFFFFE:P54:F1:Tmatch
```

Paste **only that line** back. It fully encodes what passed and what failed.

## Format

```
VIGIL : v1 : <count> : <hexBitmap> : P<pass> : F<fail> : T<tenantTag>
```

| Field | Meaning |
|---|---|
| `v2` | Schema version — bump if check order changes |
| `<count>` | Total checks run (60 in current script) |
| `<hexBitmap>` | Big-endian hex. **Bit 0 (LSB) = check #1, bit N-1 = check #N.** `1` = pass, `0` = fail. |
| `P<pass>` | Passed count |
| `F<fail>` | Failed count |
| `T<tenantTag>` | `none` \| `detected` \| `match` \| `mismatch` |

## Canonical check order (schema v2)

Check N → bit (N-1) of the hex bitmap.

| # | Check | If FAIL → impact |
|---|---|---|
| 1 | WPF assemblies loadable | **Fatal** — no UI possible |
| 2 | PowerShell 5.1+ | **Fatal** |
| 3 | Outlook COM (MAPI + folders) | Phase 3 (sync) drops |
| 4 | user32.dll P/Invoke (RegisterHotKey) | Hotkey drops, Phase 2 degrades |
| 5 | Startup folder accessible | Auto-start drops to Task Scheduler fallback |
| 6 | ~/.vigil writable | **Fatal** — no persistence |
| 7 | `[IO.File]::Replace` atomic rename | Crash-safety degrades to `Move-Item` |
| 8 | UTF-8 without BOM write | Cosmetic — JSON still valid |
| 9 | `Marshal.ReleaseComObject` | Outlook sync leaks (mitigation required) |
| 10 | Calendar Sort-before-Restrict returns items | Phase 3 calendar sync broken |
| 11 | Outlook EntryID readable | Dedup key falls back to title+source |
| 12 | Global named mutex (single-instance) | Duplicates possible |
| 13 | FindWindow / SetForegroundWindow | Can't activate existing instance |
| 14 | WPF clipboard read | Quick-Add autofill drops |
| 15 | Pester test framework | Tests can't run (install Pester) |
| 16 | `Screen.WorkingArea` (clamp on-screen) | Off-screen recovery drops |
| 17 | `DispatcherTimer` | 15-min sync scheduler broken |
| 18 | `WScript.Shell` (shortcut creation) | Auto-start via .lnk drops |
| 19 | ExecutionPolicy allows run | **Fatal unless -ExecutionPolicy Bypass** |
| 20 | **FullLanguage mode (not CLM)** | **Fatal** — VIGIL cannot run on this box |
| 21 | AppLocker allows .ps1 from user profile | **Fatal** — script blocked at launch |
| 22 | Add-Type inline C# not blocked by AMSI/EDR | Hotkey + FindWindow die |
| 23 | Outlook.Application ProgID registered | Phase 3 drops completely |
| 24 | Defender/EDR does not quarantine ~/.vigil | Data directory unsafe — pick a whitelisted path |
| 25 | Cascadia Mono / Consolas font | Falls back to system monospace |
| 26 | Device join state (dsregcmd) | Always passes; detail has AAD/Domain/TenantId |
| 27 | Tenant ID matches expected | Only runs if `-TenantId` passed |
| 28 | Outlook profile has ≥1 mail store | Phase 3 drops |
| 29 | .NET Framework ≥ 4.7.2 | WPF transparency may be degraded |
| 30 | Task Scheduler COM | Auto-start fallback drops |
| 31 | BitLocker on system drive | tasks.json plaintext may be unacceptable |
| 32 | Windows Firewall profiles | Informational |
| 33 | System proxy visible | Informational |
| 34 | Microsoft.Graph module | Informational — Phase 5 feasibility |
| 35 | Running as non-admin (as intended) | If FAIL, re-run as regular user |
| 36 | Windows version + build | Informational. Detail gates Mica (build ≥ 22000) and PerMonitorV2. |
| 37 | Outlook version + install path | Informational. Tells me x86 vs x64, 2016/2019/365. |
| 38 | Outlook process currently running | Routing: GetActiveObject vs launch-Outlook path. |
| 39 | Windows app theme (light/dark) | Auto-theme VIGIL to match OS. |
| 40 | Transparency effects enabled | If FAIL, skip glass/blur entirely — flat fill. |
| 41 | System accent color | Informational only. |
| 42 | Display layout (count + primary res) | Widget position clamp + quick-add popup routing. |
| 43 | System DPI + scale factor | Widget pixel math. |
| 44 | High contrast mode OFF | If FAIL, disable translucency and honor system colors. |
| 45 | Window animations enabled | If FAIL, drop all WPF storyboards (respect user pref). |
| 46 | WPF render tier | If tier 0 (software), drop GPU animations. |
| 47 | Console (non-remote) session | If FAIL, widget strategy changes — no glass in RDP/Citrix. |
| 48 | Power state | Informational. Could defer sync on battery in Phase 5. |
| 49 | Locale + time zone | Due-date parsing culture. |
| 50 | PowerShell host bitness | COM marshalling — must match Outlook bitness. |
| 51 | Long path support | If FAIL, warn on paths near 260 chars. |
| 52 | Free space on user profile drive | If FAIL (<100MB), abort install. |
| 53 | TEMP directory writable | Fallback atomic-write location. |
| 54 | Toast notifications allowed | Informational — Phase 5 toast feasibility. |
| 55 | Prior VIGIL install state | If TRUE in detail, trigger upgrade path. |
| 56 | Flagged emails + due-date sample | Sizes Phase 3 flag sync; detail tells me if TaskDueDate fallback is needed. |
| 57 | Cached Exchange Mode | If FAIL (disabled), Phase 3 sync runs against live mailbox — slower, add backoff. |
| 58 | Primary SMTP address resolvable | If FAIL, skip per-user task attribution; tasks are unowned. |
| 59 | Mailbox store types | Detail tells me Exchange vs PST vs delegated. Delegated stores need different folder lookup. |
| 60 | Inbox size + unread count | Informational sizing — sync volume estimate. |

## Fatal checks (any one failing = VIGIL as designed cannot run)

**1, 2, 6, 19, 20, 21** — these short-circuit everything.

## Phase-gating checks

- **Phase 1 (widget + manual tasks):** needs 1, 2, 6, 7, 12, 13, 14, 16, 17, 19, 20, 21, 22
- **Phase 2 (hotkey + quick-add):** Phase 1 set + 4, 14, 22
- **Phase 3 (Outlook sync):** Phase 1 set + 3, 9, 10, 11, 23, 28
- **Phase 4 (polish + auto-start):** Phase 1 set + 5, 18 *or* 30

## Feature decisions I'll make from the string

| Signal | Decision |
|---|---|
| Any fatal check fails | Stop. Report what's blocking. |
| Phase 3 checks fail, Phase 1-2 pass | Ship VIGIL as manual-only. Hide sync UI. |
| #7 fails | Swap `[IO.File]::Replace` for safer `Move-Item` + manual backup. |
| #18 fails but #30 passes | Auto-start via Task Scheduler instead of Startup folder. |
| #24 fails | Move data dir to a whitelisted location (ask user which). |
| #31 fails | Add DPAPI `ProtectedData` wrap around tasks.json writes. |
| #27 returns `mismatch` | Stop. Wrong machine / wrong tenant. |
| #22 fails | **Hard stop.** No workaround — EDR is blocking inline C# compilation. Escalate to IT. |

## Feature decisions from v2-only checks

| Signal | Decision |
|---|---|
| #36 detail shows build < 22000 | Use flat `rgba(0,0,0,0.80)` for title bar instead of Mica |
| #37 shows Outlook x86 and #50 shows PS x64 (or vice versa) | Force 32-bit PowerShell for COM — add launcher wrapper |
| #38 = not running at launch | Phase 3 startup sync shows "connecting..." placeholder for 2-3s |
| #39 detail = light | Ship light-mode variant of the Apple skin |
| #40 FAIL or #44 FAIL or #47 FAIL (any of them) | Skip all glass/blur/transparency. Flat fill only. |
| #42 detail shows > 1 display | Quick-add popup uses `Screen.FromPoint(Cursor.Position)` |
| #43 scale != 1.0 | Scale widget dimensions by DPI ratio at startup |
| #45 FAIL | Remove all `Storyboard` animations — respect reduced motion |
| #46 tier 0 | Drop completion fade + sync spin — static UI |
| #49 non-English locale | Use `CultureInfo.CurrentCulture` for dueLabel formatting |
| #51 FAIL | Validate `~/.vigil` path length + warn if close to 260 chars |
| #52 FAIL | **Block install.** Cannot safely atomic-write with <100 MB free. |
| #55 shows prior install | Read existing tasks.json and migrate schema if needed |

## Example decode (by hand)

String: `VIGIL:v2:55:7FFFFFFFFFFFFFF:P55:F0:Tnone`

- Count = 35
- Hex `7FFFFFFFF` = 34 bits set (0x7FFFFFFFF is 35 bits, but MSB bit 34 is 0 — wait, `0x7FFFFFFFF` is 35 bits 0..34 all set, 35 checks pass)
- P35 F0 — clean run
- Tnone — no `-TenantId` passed

Example failure: `VIGIL:v1:35:7FFFFFBFF:P34:F1:Tnone`

- Bit 10 (0-indexed) = 0 → check #11 failed
- Check #11 = "Outlook EntryID readable" → dedup will fall back to title+source
