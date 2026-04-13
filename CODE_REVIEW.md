# VIGIL — Code Review (build50)

Written from running PSScriptAnalyzer on pwsh 7.5.5 against `VIGIL.ps1` and
`Test-Vigil.ps1`, plus structural greps. This is **not** a list of bugs.
Most items are *maintainability* improvements, ranked by impact.

---

## 1. Overall health snapshot

| Metric | Value | Verdict |
|---|---|---|
| `VIGIL.ps1` lines | **2155** | Single file is approaching the upper limit. Split is worth considering. |
| `Test-Vigil.ps1` lines | 547 | Healthy. 106 assertions, all green. |
| Functions in main script | **35** | Reasonable count but some are too large. |
| `$Global:` references | 77 | High. Each one is an `PSAvoidGlobalVars` warning from the analyzer. |
| `$script:` references | 82 | High. |
| `try` blocks | 58 | Lots of defensive code. |
| Empty `catch {}` blocks | 32 | About half are intentional (Marshal.ReleaseComObject), half are silent error swallows. |
| `Add-Type` calls | 7 | Three of these are inline C#, four are assembly loads. |
| `Marshal.ReleaseComObject` calls | 9 | Outlook COM lifecycle still hand-managed. |
| Tests passing | **106 / 106** | Cross-platform via `-NoUI` mode on Linux pwsh. |
| Largest functions (real, excluding embedded XAML herestrings) | `Show-QuickAdd` ~140, `Show-VigilEditPrompt` ~110, `Build-TaskCard` ~120, `Sync-VigilFromOutlook` ~175 | All four candidates for decomposition. |

### PSScriptAnalyzer warnings (Warning + Error severities)

| Count | Rule |
|---|---|
| 80 | `PSAvoidGlobalVars` |
| 30 | `PSAvoidUsingEmptyCatchBlock` |
| 21 | `PSReviewUnusedParameter` (mostly `$sender`/`$e` in WPF event handlers — false positives) |
| 11 | `PSUseSingularNouns` (e.g., `Save-VigilTasks` should be `Save-VigilTask`) |
| 7 | `PSUseApprovedVerbs` (e.g., `Toggle-Done`, `Handle-ContextAction`, `Apply-VigilFluentBackdrop` would-be) |
| 4 | `PSUseShouldProcessForStateChangingFunctions` |
| 1 | `PSUseDeclaredVarsMoreThanAssignments` |

`Test-Vigil.ps1`: 16 × `PSAvoidUsingWriteHost` only — that's the test reporter intentionally writing to host. Not a real concern.

---

## 2. Backend / logic layer

### What's good
- **Cross-platform `-NoUI` mode** lets pure logic run on Linux pwsh — every commit verified by 106 tests.
- **DPAPI wrap/unwrap** properly degrades to passthrough on non-Windows.
- **Save/Load atomic write** uses `[IO.File]::Replace` for tasks and `Delete + Move` for settings (correct after the earlier crash fix).
- **Settings merge-with-defaults** in `Load-VigilSettings` guarantees every key exists on the returned object — no more "missing field crash" failure mode.
- **Helper functions** (`Add-VigilTask`, `Update-VigilTask`, `Remove-VigilTask`, `Get-VigilTaskById`, `Search-VigilTasks`, `Filter-VigilTasks`, `Sort-VigilTasks`, `Export-VigilMarkdown`, `Get-VigilOverdueTasks`) are all individually testable and are tested.
- **Outlook COM lifecycle** uses `try / finally` blocks consistently with the canonical 2× `GC.Collect / WaitForPendingFinalizers` pattern.

### Issues / improvements

#### B1. Global-state proliferation (HIGH impact)
77 references to `$Global:` and 82 to `$script:` is a smell. Every one of these makes
the code harder to test in isolation and harder to reason about.

**Concrete fix**: collapse all global state into a single `$Global:Vigil` hashtable:

```powershell
$Global:Vigil = @{
    Tasks       = @()
    Settings    = $null
    SelectedId  = $null
    Mutex       = $null
    TrayIcon    = $null
    HotkeyId    = 9001
}
```

Then `$Global:VigilTasks` becomes `$Global:Vigil.Tasks`, and so on. Pros:
- Single source of truth for runtime state.
- Easy to inspect/dump for debugging.
- One name to look up in IDE.
- Reduces 77 + 82 = 159 references to maybe ~50 (deduplication).

PSScriptAnalyzer warnings drop from 80 + global-related to 1.

#### B2. Approved-verb naming (LOW impact, easy)
- `Toggle-Done` → `Set-VigilTaskDoneState` or just merge into `Update-VigilTask`
- `Handle-ContextAction` → `Invoke-VigilContextAction`
- `Refresh-Render` → `Update-VigilWidget` (refresh isn't approved, update is)
- `Move-VigilSelection` → `Step-VigilSelection`

These trip the analyzer but don't break anything. Quick rename, no behavior change.

#### B3. Singular-noun functions (LOW impact, easy)
PowerShell convention says `Save-VigilTask` (singular) even when saving multiple.
Eight functions need renaming: `Save-VigilTasks`, `Load-VigilTasks`,
`Sort-VigilTasks`, `Filter-VigilTasks`, `Search-VigilTasks`, `Get-VigilOverdueTasks`,
plus `Build-TaskCard`. Pure cosmetic.

#### B4. `Sync-VigilFromOutlook` is 175 lines and has 3 nearly-identical sub-blocks (MED impact)
Calendar / Flagged / Tasks each follow the same pattern: open folder → restrict →
foreach → release. Extract a helper:

```powershell
function Read-VigilOutlookFolder {
    param(
        $namespace,
        [int]$folderId,        # 9 cal, 6 inbox, 13 tasks
        [string]$restrict,     # DASL filter
        [scriptblock]$mapper,  # ($item) -> task
        [hashtable]$existing,
        [string]$source
    )
    # ...
}
```

Reduces `Sync-VigilFromOutlook` from 175 lines to maybe 60. Each folder type becomes 5 lines.

#### B5. `Format-DueLabel` is 222 reported lines (MED) — actually mismeasured
The reported "222 lines" is the gap between `Format-DueLabel` and the next function;
the gap contains the **main XAML herestring**. The function itself is fine.
**No action needed**, but worth knowing the metric is misleading.

#### B6. Empty catch blocks (MED impact) — 30 of them
Many are intentional (Marshal.ReleaseComObject failure on a null ref is fine). But
several swallow errors that should at least log:
- `try { [System.Windows.Forms.Cursor]::Position } catch {}` — silent on multi-monitor failure
- `try { $cmd = Get-Command pwsh } catch {}` — should log
- `try { $first.EntryID } catch {}` — should log on bad item
**Action**: scan with `grep` for `catch {}`, pick the ~10 that actually swallow useful errors, switch them to log via `Write-VigilLog`.

#### B7. Outlook lifecycle is hand-coded 9 times (MED impact)
Each `Sync-VigilFromOutlook` sub-block manually releases each RCW in a fixed order.
This pattern is copy-pasted. A helper:

```powershell
function Invoke-WithOutlookCom {
    param([scriptblock]$Action)
    $ol = $null; $ns = $null
    try {
        $ol = ...; $ns = ...
        & $Action $ns
    } finally {
        if ($ns) { try { [Marshal]::ReleaseComObject($ns) } catch {} }
        if ($ol) { try { [Marshal]::ReleaseComObject($ol) } catch {} }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}
```

Saves ~80 lines and centralizes COM lifecycle policy.

#### B8. `Save-VigilSettings $Global:VigilSettings` repeated 9 times (LOW impact)
Wrap in a `Sync-VigilSettings` zero-arg helper that always reads/writes the global.
9 copies → 1 call site each.

---

## 3. Frontend / XAML layer

### What's good
- **`ThemeMode="System"`** in all 4 windows → single switch for Fluent + Mica + light/dark adaptation.
- **`{{THEMEMODE}}` placeholder substitution** lets the same script run on PS 5.1 (no Fluent) and pwsh 7.5 (full Fluent) without two XAML variants.
- **Native Segoe MDL2 caption buttons** with proper red close hover via `Add_MouseEnter`.
- **Search + keyboard nav + selection highlight** are all wired through the DynamicResource brush system, so they adapt with theme changes.
- **No `StaticResource` brushes left** — every color is a Fluent theme token. Light theme works.

### Issues / improvements

#### F1. Embedded XAML is 4 separate herestrings (~330 lines total) (MED impact)
Lines 799–943, 1304–1329, 1331–1379, 1756–1784. Each is a `@'...'@`. Pros: single
file. Cons: editing XAML in a PS string is painful (no syntax highlighting, harder
to spot mismatches), and growing the file size.

**Option**: extract them into 4 sibling files (`VIGIL.main.xaml`, `VIGIL.popup-quickadd.xaml`,
`VIGIL.popup-edit.xaml`, `VIGIL.popup-welcome.xaml`) loaded via `Get-Content`. Adds files
but pays back in maintenance.

**Counter-option**: keep them embedded but extract a helper:
```powershell
function Initialize-VigilWindow([string]$xaml) {
    $themeAttr = if ($script:HasFluent) { 'ThemeMode="System"' } else { '' }
    $ready = $xaml.Replace('{{THEMEMODE}}', $themeAttr)
    [xml]$xml = $ready
    [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
}
```
Replaces 4 × 6-line blocks with 4 × 1-line calls. Smaller win but no new files.

#### F2. `Build-TaskCard` programmatically constructs WPF elements (HIGH maintenance cost)
~120 lines of `New-Object System.Windows.Controls.X` + property assignments, when this could
be a `DataTemplate` defined once in XAML. Right now every visual change to a task row
requires editing PowerShell, not XAML. Hard to preview, hard to test, hard to skin.

**Refactor**:
1. Define a `VigilTask` viewmodel with the right properties (`title`, `due`, `isOverdue`, `priority`, `isSelected`, `metaText`)
2. Define a `<DataTemplate DataType="...">` in the XAML with bindings
3. `ItemsControl.ItemsSource = $observableCollection` (an `ObservableCollection<VigilTask>`)
4. Refresh-Render rebuilds the collection instead of the visual elements

This is the single biggest maintainability win. Cost: **2-3 hours**, risk: medium (need to
make sure bindings re-fire on selection changes).

#### F3. The `Show-QuickAdd` pill creation duplicates code (MED impact)
Both priority pills and due-date pills are built with nearly identical PS code (~50 lines
each). A helper `New-VigilPillButton -label -tag -isSelected -onClick` would cut both
sections in half.

#### F4. Welcome / Edit / Quick-Add popups share boilerplate (MED impact)
Each has:
- The same `{{THEMEMODE}}` substitution
- The same `XmlNodeReader / XamlReader.Load`
- The same `prevFg = GetForegroundWindow()` save/restore
- The same `Add_Closed { SetForegroundWindow($prevFg) }`
- The same Escape-to-close keybind

Extract into:
```powershell
function Show-VigilPopup {
    param([string]$xaml, [scriptblock]$wireUp)
    $prevFg = [VigilWin32]::GetForegroundWindow()
    $win = Initialize-VigilWindow $xaml
    & $wireUp $win
    $win.Add_Closed({ ... }.GetNewClosure())
    $win.Add_KeyDown({ ... })
    $win.Show(); $win.Activate() | Out-Null
}
```

Each popup function shrinks by ~30 lines.

#### F5. Title bar "VIGIL FLUENT" label is dev-only (LOW impact)
The `FluentLabel` next to the wordmark shows "FLUENT" or "FLAT". Useful during the
Mica troubleshooting phase, no longer needed now that Fluent works. **Action**: remove.
Saves ~6 lines of XAML + 4 lines of code-behind.

#### F6. Title bar 5-column grid is rigid (LOW impact)
Each addition (info button, search, etc.) requires renumbering Grid.Column attributes.
Switch to a `DockPanel` with 3 `StackPanel` children: left (wordmark + badges), center
spacer, right (sync/sort/info/min/close). Easier to add buttons later.

---

## 4. Test coverage gaps

Currently covered (106 tests):
- Schema, sort, filter, format, JSON round-trip, settings merge, null healing, Update
  helpers, Add/Remove, overdue, export markdown, search.

**Missing coverage**:
1. **Outlook COM mocking** — the entire `Sync-VigilFromOutlook` path has 0 tests because
   Linux pwsh has no Outlook. Could be solved with a `$script:OutlookProvider` injection
   point so tests can pass a fake provider. **Medium effort**, high payoff for confidence
   in the 175-line sync function.
2. **Concurrent save** — what if VIGIL is open in two instances? (mutex prevents it, but
   what if the mutex is broken?) Verify the file lock protection.
3. **Disk full / read-only** on Save — no test simulates ENOSPC.
4. **Settings migration** — when a new field is added to defaults, old settings.json should
   merge cleanly. Have one test for this but could expand.
5. **Hotkey collision** — `RegisterHotKey` returning false. Currently logs and continues,
   but not asserted in tests.
6. **Welcome flag persistence** — set, save, reload, verify. Not tested.
7. **Search + filter combination** — search "billing" + filter "outlook" → only outlook
   tasks containing "billing". Not tested as a combo.

---

## 5. Dead code / cleanup opportunities

| Item | Action |
|---|---|
| `Get-VigilVisibleTasks` defined but only called from `Move-VigilSelection` | Inline if it's a single caller, or use it from `Refresh-Render` too (currently duplicates the logic) |
| `Set-VigilSourceRef` only called from inside `Sync-VigilFromOutlook` | OK as-is, but could be local |
| `Apply-VigilFluentBackdrop` was deleted in build44 — confirm no references | grep shows 0, clean |
| 32 empty catch blocks | Audit, log the ~10 that swallow real info |
| Fluent label `FluentLabel` shows `FLUENT`/`FLAT` | Was for debugging, can remove |
| `lib/` folder still has the WebView2 DLLs | We never used them since pivoting to .NET 9 ThemeMode. Either remove or keep as a fallback path |
| Color brush definitions in Window.Resources | All 15 deleted in build48 — clean |
| `IconButton`, `CloseButton`, `PrimaryButton`, `GhostButton` styles | All 4 deleted in build45 — clean |

---

## 6. Refactor priorities (ranked by ROI)

| # | Item | Effort | Risk | Lines saved | Maintainability gain |
|---|---|---|---|---|---|
| 1 | **`Build-TaskCard` → DataTemplate** | 3h | Med | ~120 | ★★★★★ |
| 2 | **All globals → `$Global:Vigil` hashtable** | 1h | Low | ~50 | ★★★★ |
| 3 | **Outlook COM helper extraction** | 1h | Low | ~80 | ★★★ |
| 4 | **Popup boilerplate → `Show-VigilPopup`** | 30m | Low | ~60 | ★★★ |
| 5 | **`Sync-VigilFromOutlook` folder helper** | 30m | Low | ~50 | ★★★ |
| 6 | **Test injection for Outlook provider** | 1h | Low | +30 | ★★★ |
| 7 | **Approved-verb / singular-noun rename** | 30m | Low | 0 | ★★ |
| 8 | **Audit empty catch blocks** | 30m | Low | 0 | ★★ |
| 9 | **Title bar DockPanel** | 30m | Low | ~5 | ★★ |
| 10 | **XAML extraction to sibling files** | 1h | Med | 0 (just relocated) | ★★ |
| 11 | **Remove `FluentLabel` debug pill** | 5m | None | ~10 | ★ |
| 12 | **Remove `lib/` WebView2 DLLs (or document)** | 5m | None | -3 files | ★ |

**Totals**: ~9 hours of work, ~375 lines deleted, analyzer warnings down from 154 to ~10.

---

## 7. Three things I would do tomorrow

If you only have one hour:

1. **Item #2** (globals → `$Global:Vigil` hashtable). Pure rename, low risk, single biggest
   analyzer-warning reduction. After this, the analyzer is mostly silent.
2. **Item #11** (remove FluentLabel debug pill). Small cleanup, finalizes the visual.
3. **Item #6** (Outlook test injection). Even a minimal mock unlocks testing the 175-line
   sync function on Linux. Without it, every Outlook code change is "test in production."

If you have 4 hours:

Add **item #1** (`Build-TaskCard` → DataTemplate). This is the single biggest visual
flexibility win and the biggest reduction in "code that exists only because we chose to
build the UI from PowerShell."

---

## 8. What's already been done very well

- **Cross-platform testable core** (`-NoUI`, Linux pwsh) — most PowerShell-WPF apps cannot
  do this. 106 tests prove the data layer is solid.
- **Encryption layer** with graceful Windows / non-Windows split.
- **Fluent integration** via `ThemeMode` is the cleanest possible path on .NET 9.
- **Settings merge-with-defaults** prevents a whole class of upgrade bugs.
- **Build stamp + log on every change** lets remote-triage diagnose runtime issues from
  one log line.
- **ASCII-clean source + UTF-8 BOM** survives the user's copy-paste workflow without
  encoding drift.
- **Self-healing null-entry purge** on startup repairs damage from a long-fixed bug.
- **PS-5.1 compatibility holdouts**: `New-Object` instead of `::new()`, `param($s,$e)`
  instead of `$this`, flattened nested subexpressions in strings — all fixed.
- **Preflight script** with 60 checks and bitmap result string — best-in-class for an
  unsigned PowerShell tool on a corporate machine.
- **Test runner is dependency-free** — no Pester needed.
