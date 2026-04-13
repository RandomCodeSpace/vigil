# VIGIL — Plan Eng Review

Opinionated review of the VIGIL spec (sections 1–16). Scope: architecture,
code quality risks, test coverage, performance. Every finding has a severity
and confidence score. Non-obvious decisions are called out.

---

## Step 0 — Scope Challenge

**Size check:** Spec targets ~1500 LOC in a single `.ps1`. For PowerShell +
WPF + COM + P/Invoke + animations + quick-add popup, that is *tight*. Honest
estimate: 1800–2400 LOC once error handling, XAML strings, and COM release
code land. Not a blocker — flag so nobody panics when it crosses 1500.

**What already exists (Windows built-ins you're not fully using):**

| Need | Built-in you could reuse | Currently in plan |
|---|---|---|
| Toast notifications | `Windows.UI.Notifications.ToastNotificationManager` via WinRT | Deferred (Phase 5) — fine |
| JSON parsing | `ConvertTo-Json` / `ConvertFrom-Json` | ✓ implicit |
| Due date parsing | `[datetime]::Parse` with culture handling | Not specified — **gap** |
| Scheduled work | `DispatcherTimer` | ✓ |
| Atomic rename | `[System.IO.File]::Replace` (three-arg, preserves ACLs) | Plan says "rename" — use `Replace` |

**Minimum viable cut:** Phase 1 alone (widget + JSON + manual tasks) is a
usable product on day one. Phases 2–4 are additive. Good phasing.

**Complete-option check (boil the lake):** At AI-assisted speeds, phases 1–4
are ~half a day of work, not four weeks. Recommend building all four in one
pass with tests per phase, not four separate milestones.

---

## 1. Architecture Review

### 1A. `[P1]` (confidence 9/10) — **COM object lifecycle leak**

Outlook COM objects (`$ol`, `$ns`, `$cal`, individual items, `Items`
collections) **must** be released via
`[System.Runtime.InteropServices.Marshal]::ReleaseComObject` in a `finally`
block. If you skip this, Outlook.exe refuses to exit after the user closes
it, because VIGIL still holds RCWs. This is the #1 Outlook COM footgun and
the spec does not mention it.

**Fix in plan:**

```
┌──────────────────────────────────────────┐
│ try {                                    │
│   $ol  = GetActiveObject or New-Object   │
│   $ns  = $ol.GetNamespace("MAPI")        │
│   $cal = $ns.GetDefaultFolder(9)         │
│   foreach ($item in $cal.Items) {        │
│     ...map to task...                    │
│     [Marshal]::ReleaseComObject($item)   │
│   }                                      │
│ } finally {                              │
│   [Marshal]::ReleaseComObject($cal)      │
│   [Marshal]::ReleaseComObject($ns)       │
│   [Marshal]::ReleaseComObject($ol)       │
│   [GC]::Collect(); [GC]::WaitForPendingFinalizers() │
│ }                                        │
└──────────────────────────────────────────┘
```

Double-`GC.Collect` is the canonical PowerShell idiom for COM cleanup.

### 1B. `[P1]` (confidence 9/10) — **Recurring meetings + Restrict filter**

Spec says "IncludeRecurrences = $true" on Calendar. This **only works if you
call `.Sort("[Start]")` first and then `.Restrict(...)` with a DASL or jet
filter**. Iterating `$cal.Items` directly without `Sort` + `Restrict` returns
master recurrences, not instances — you'll get the series "Daily Standup"
once instead of tomorrow's instance. This is the classic MAPI gotcha.

**Correct pattern:**
```
$items = $cal.Items
$items.IncludeRecurrences = $true
$items.Sort("[Start]")
$filter = "[Start] >= '$(Get-Date -Format 'g')' AND [Start] <= '$((Get-Date).AddHours(24).ToString('g'))'"
$filtered = $items.Restrict($filter)
```
Sort MUST come before Restrict. Order matters.

### 1C. `[P2]` (confidence 8/10) — **Hotkey collision with screen readers**

`Ctrl+Win+A` is unused by most apps but **Narrator** uses `Ctrl+Win+Enter`
and Windows Action Center historically used `Win+A`. On Win11 with some
accessibility tooling, `Ctrl+Win+A` can conflict. Not fatal, but worth
either:

- Making the hotkey user-configurable via `settings.json` (recommended), or
- Defaulting to `Ctrl+Alt+Space` which is cleaner and reserved nowhere.

### 1D. `[P2]` (confidence 8/10) — **Single-instance via Global mutex needs cleanup**

`Global\VIGIL_TaskTracker` is fine but you must:
1. Release mutex in a `finally` on exit (crash = orphan mutex until reboot
   only if you use `AbandonedMutexException` correctly — Mutex is actually
   OS-cleaned on process death, so this is okay, but you still need
   `ReleaseMutex()` to avoid `AbandonedMutexException` on the next run).
2. On duplicate launch, use `FindWindow` (user32) by window title to bring
   existing instance to front, not just exit silently. Spec hints at this —
   make it explicit in Phase 1.

### 1E. `[P3]` (confidence 7/10) — **Clock drift in "dueLabel"**

`dueLabel` is precomputed and stored in JSON ("Today 3:00 PM", "Tomorrow").
On the next morning, "Today 3 PM" is now *yesterday* but the label still
says Today until next sync. **Compute `dueLabel` at render time from
`dueDate`**, don't persist it. Store only the machine-readable ISO datetime.

---

## 2. Code Quality

### 2A. `[P1]` (confidence 9/10) — **JSON atomic write: use `File.Replace`, not rename**

Spec says "write to tmp then rename". On Windows, `Move-Item -Force` is
**not atomic** if the destination exists (it's delete + move). Use
`[System.IO.File]::Replace($tmp, $target, $backup)` — one atomic syscall
that even creates the backup for you. This is exactly what you want.

### 2B. `[P2]` (confidence 8/10) — **UTF-8 *without* BOM, not with**

Spec says "UTF-8 with BOM (PowerShell default)". That default bites you:
many JSON parsers and diff tools choke on BOM. Write with
`[System.IO.File]::WriteAllText($path, $json, [System.Text.UTF8Encoding]::new($false))`
— no BOM. Also consistent with how `.json` files are expected on Windows.

### 2C. `[P2]` (confidence 7/10) — **XAML as herestring = maintenance pain at 1500 LOC**

A single 1500-line `.ps1` with embedded XAML strings is cute but hard to
edit. Split the XAML into a `[xml]$xaml = @'...'@` block near the top, keep
it separate from the code-behind logic, and mark regions with `#region` /
`#endregion`. This is the only affordable structure in a single-file script.

### 2D. `[P3]` (confidence 7/10) — **ID collision: 8-char GUID prefix**

Birthday-paradox collision for 8 hex chars (~4.3B space) hits ~50% at ~65k
tasks. You'll never have 65k active tasks, but dedupe-by-id in a sync loop
can still bite for long-lived data. Either use full GUID or add a collision
check on create. Cheap fix.

### 2E. `[P3]` (confidence 6/10) — **`source: outlook-flag` dedup by `title`** is fragile — users rename email subjects, and your dedup key breaks. Prefer dedup by `EntryID` from Outlook (persistent across sessions). Store it as a hidden `sourceRef` field on the task.

---

## 3. Test Review

Spec has **zero** tests in the plan. For a single-file PowerShell app this
is fixable but non-negotiable. Pester is the framework (pre-installed on
Win10 1803+, bundled on Win11).

### Coverage diagram

```
VIGIL.ps1
│
├── Data layer (tasks.json CRUD)
│   ├── [GAP] New-VigilTask → round-trip save/load
│   ├── [GAP] Atomic write crash-safety (simulate mid-write crash)
│   ├── [GAP] Corrupt JSON recovery (bad file → load backup)
│   └── [GAP] Dedup logic (same title+source+due = skip)
│
├── Outlook sync
│   ├── [GAP] Calendar Sort-before-Restrict pattern
│   ├── [GAP] Recurring meeting expansion (1 instance per day)
│   ├── [GAP] Flagged email → task mapping
│   ├── [GAP] Outlook importance → VIGIL priority mapping
│   ├── [GAP] COM release on exception (integration test w/ mocked COM)
│   └── [GAP] Stale calendar item auto-completion
│
├── UI layer
│   ├── [GAP] [→E2E] Collapse/expand state persists to settings.json
│   ├── [GAP] [→E2E] Drag updates posX/posY
│   ├── [GAP] [→E2E] Quick-add popup Tab/Enter/Escape flow
│   └── [GAP] [→E2E] Focus restore after popup close
│
├── Hotkey
│   ├── [GAP] RegisterHotKey success path
│   ├── [GAP] Unregister on window Closing
│   └── [GAP] Clipboard empty → popup with empty title
│
└── Lifecycle
    ├── [GAP] Single-instance mutex behavior
    ├── [GAP] Off-screen position clamp on resolution change
    └── [GAP] Log rotation at 500 lines

COVERAGE: 0/21 paths tested (0%)
```

**Iron rule:** The JSON atomic-write and the Sort-before-Restrict pattern
both need regression tests before shipping. They're the two paths most
likely to break silently.

**Minimum test bar for Phase 1 ship:**
1. Task CRUD round-trip (5 tests)
2. Atomic write survives simulated crash (1 test)
3. Corrupt JSON → backup recovery (1 test)
4. Dedup logic (2 tests)
5. Position clamp on off-screen (1 test)

That's 10 Pester tests, ~30 minutes to write with AI assistance.

---

## 4. Performance Review

### 4A. `[P2]` (confidence 8/10) — **`$cal.Items` full enumeration is O(n)**

Unfiltered `.Items` enumeration on a calendar with 10k+ items (normal for a
3-year-old corp mailbox) can take **5–15 seconds per sync**. Your 10-second
timeout will eat this. Fix: ALWAYS use `.Restrict()` with a date filter, and
NEVER iterate `.Items` directly without first sorting + restricting. This is
the same pattern as 1B — call it out in the perf section too because the
symptom is "sync randomly times out" not "sync is wrong."

### 4B. `[P3]` (confidence 7/10) — **WPF opacity animation on every task**

Overdue pulse via `DoubleAnimation` + `RepeatBehavior=Forever` on every
overdue task card allocates a Storyboard per card. With 20 overdue items
(rare but possible), you're burning ~3% CPU continuously on a laptop. Use a
single shared `Storyboard` at the window level driving an attached property,
or only pulse the most-overdue item. Not a blocker for Phase 1.

### 4C. `[P3]` (confidence 6/10) — **Log rotation on startup blocks first render**

"Trim log to 500 lines on startup" — if you read the whole log into memory,
trim, and rewrite synchronously, that's 50–200ms before the window shows.
Move to an async `Register-ObjectEvent` on window Loaded, or just append to
a rolling file and rotate on size (e.g. >256KB), not line count.

---

## Failure Modes

| Failure mode | Test? | Handled? | User sees? |
|---|---|---|---|
| Outlook not running | no | yes (status) | ⚠ offline indicator ✓ |
| COM throws mid-iteration | no | partial | silent hang if no try/finally ✗ |
| JSON corrupt | no | yes (backup) | recovery ✓ |
| Disk full on atomic write | no | no | **silent data loss** ✗ |
| RegisterHotKey returns false | no | no | hotkey silently dead ✗ |
| Off-screen window after resolution change | no | yes (clamp) | ✓ |
| Outlook EntryID changes (rare) | no | no | duplicate tasks ✗ |

**Critical gaps:** disk full on save, RegisterHotKey failure, COM mid-iteration throw. Add handling + tests for each before shipping Phase 1.

---

## NOT in Scope (explicitly deferred)

- **EWS / Graph API fallback** — corp IT approval required, breaks "zero network" rule. Deferred to Phase 5.
- **Encryption at rest** — `~/.vigil/tasks.json` is plaintext. On a corp machine with BitLocker + user profile isolation, this is acceptable. Flag for the user to confirm.
- **Multi-monitor DPI scaling** — WPF handles most of it but per-monitor DPI v2 needs `app.manifest`. Single-file `.ps1` can't ship a manifest. Accept 100% scaling on primary monitor as the supported case.
- **Touch / pen input** — desktop widget only.
- **Localization** — English only; `dueLabel` computed in current culture.
- **Accessibility** — screen reader support via WPF `AutomationProperties`. Not specified. **Flag:** if this ships inside a corp that requires Section 508, it's a blocker.

---

## TODOS (for later, not this PR)

1. **System tray icon** — reduces taskbar clutter when collapsed. Phase 5.
2. **Export tasks as markdown** — 10-line feature, nice for standups.
3. **Hotkey rebinding UI** — follows from 1C.
4. **Prompt-to-task via LLM** — "remind me to review the PR when I get in" → parse date. Out of scope for the "zero network" constitution unless you ship a local model.

---

## Parallelization Strategy

Phases 1 and 2 share the same WPF window code — **sequential**, not
parallel. Phase 3 (Outlook) is independent of the UI and can be built in a
second worktree after Phase 1 data layer lands. Phase 4 (polish) runs last.

```
Lane A: Phase 1 (widget + data) → Phase 2 (hotkey + popup) → Phase 4 (polish)
Lane B: Phase 3 (Outlook sync)  ──────┐ merges into Lane A after Phase 1
```

Lane B depends on the JSON data layer from Phase 1 existing. One parallel lane.

---

## Completion Summary

| Section | Result |
|---|---|
| Step 0 — Scope | Accepted. LOC estimate revised 1500 → 2000. |
| Architecture | 5 issues (2 P1, 2 P2, 1 P3) |
| Code Quality | 5 issues (1 P1, 2 P2, 2 P3) |
| Test Review | 0/21 paths covered — **21 gaps** |
| Performance | 3 issues (1 P2, 2 P3) |
| Failure modes | 3 critical gaps (disk full, hotkey fail, COM throw) |
| Parallelization | 1 main lane, 1 parallel Outlook lane |

**Verdict:** Plan is strong on product vision and constraints. Needs work on
COM lifecycle, atomic write primitive, Sort-before-Restrict pattern, and
test coverage before Phase 1 ships. All of these are 15–30 minute fixes
each — don't defer any of them.

The three P1s (1A, 1B, 2A) are the ones that will bite silently in
production. Fix those first. Everything else is polish.
