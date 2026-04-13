# VIGIL — Apple-Inspired UI Reskin

Replaces the "electric cyan glassmorphism" look from spec §5 with an
Apple-style visual language: near-black + light gray binary, SF Pro
typography, single-accent Apple Blue, and a translucent dark glass title
bar. Same widget, same interactions, different skin.

Sourced from `~/.claude/design/apple.md`.

---

## 1. Color Tokens

Drop the cyan system. Replace `5.1` table with:

| Role | Token | Value |
|---|---|---|
| Widget surface (dark mode) | `surface.base` | `#000000` |
| Title bar glass | `surface.glass` | `rgba(0,0,0,0.80)` + `backdrop-filter: saturate(180%) blur(20px)` |
| Task card background | `surface.card` | `#1d1d1f` |
| Elevated card (hover) | `surface.card.hi` | `#272729` |
| Card separator | `surface.divider` | `rgba(255,255,255,0.08)` |
| Primary text | `text.primary` | `#ffffff` |
| Secondary text | `text.secondary` | `rgba(255,255,255,0.56)` |
| Tertiary text / timestamps | `text.tertiary` | `rgba(255,255,255,0.32)` |
| **Interactive accent (the ONLY color)** | `accent` | `#2997ff` (Bright Blue — on-dark variant) |
| Focus ring | `focus` | `#0071e3` |
| Success (completion flash) | `success` | `#30d158` |
| Destructive (delete) | `destructive` | `#ff453a` |

### Priority mapping — the hard call

Apple's design language refuses multi-color systems. Four priority dots in
red/amber/white/blue would violate the binding rule: *"single accent color
in a sea of neutrals."* Two ways to resolve:

- **A (recommended, Apple-pure):** Express priority through **typography
  weight and position**, not color. Critical = `weight 700`, high =
  `weight 600`, normal = `weight 400`, low = `weight 400` + `rgba(255,255,255,0.56)`.
  Order in list = priority. Zero extra colors introduced.
- **B (pragmatic):** Keep one non-accent priority signal — a thin 2px left
  border on cards using semantic system colors, restricted to critical and
  overdue only (red `#ff453a`). Everything else has no left border.

Recommend **A** for the default skin; users who want loud priority can
toggle to **B** in settings.

---

## 2. Typography

Replace the Cascadia Mono / Consolas stack — monospace is wrong for an
Apple-styled widget. Use:

```
font-family: 'SF Pro Text', 'Segoe UI Variable', 'Segoe UI', system-ui, sans-serif;
```

SF Pro isn't installed on Windows by default. `Segoe UI Variable` (Win11)
and `Segoe UI` (Win10) are the honest fallbacks — they're the closest
geometric neo-grotesques Windows ships. Do not bundle SF Pro; Apple does
not license it for redistribution.

### Scale (adapted from apple.md §3)

| Role | Size | Weight | Tracking | Use in VIGIL |
|---|---|---|---|---|
| Widget title | 17px | 600 | -0.374px | `⚡ VIGIL` wordmark |
| Task title | 15px | 500 | -0.24px | Primary card text |
| Critical task title | 15px | 700 | -0.24px | Option A priority weighting |
| Metadata / due label | 12px | 400 | -0.12px | "Today 3:00 PM", source tag |
| Badge count | 11px | 600 | 0 | "3" in title bar |
| Quick-add popup title | 21px | 400 | 0.196px | "New task" |
| Button | 13px | 500 | -0.08px | Action buttons |

WPF implementation: `TextOptions.TextFormattingMode="Ideal"` and
`UseLayoutRounding="True"` on the root — essential for crisp SF-style
text at small sizes on Windows.

---

## 3. Layout & Radii

Apple radius scale collapses to three values in a widget context:

| Element | Radius |
|---|---|
| Outer widget frame | `12px` |
| Task card | `8px` |
| Buttons (title bar controls) | `5px` |
| Quick-add pill buttons ("Today"/"Tomorrow"/"This Week") | `980px` (full pill) |
| Circular media-style controls (sync button) | `50%` |

Spacing: base unit **8px**, with a dense micro scale `{2, 4, 6}` for
typography inset. Apple cards don't have borders — remove the `#1e2a3a`
border from spec §5 entirely. Elevation is signaled by background color
delta alone.

---

## 4. Widget Mode — Rewritten

```
┌───────────────────────────────────────────────┐
│                                               │  ← 12px outer radius
│  VIGIL                        3    ⟳   —  ✕  │  ← 44px glass title bar
│                                               │     rgba(0,0,0,0.80) + blur(20px)
├───────────────────────────────────────────────┤
│                                               │
│  Review PR for payment service                │  ← 15px / weight 700 (critical)
│  Today at 3:00 PM                             │  ← 12px / secondary
│                                               │
│  ─────────────────────────────────────────── │  ← divider rgba(255,255,255,0.08)
│                                               │
│  Prepare demo for standup                     │  ← 15px / weight 600 (high)
│  Manual                                       │
│                                               │
│  ─────────────────────────────────────────── │
│                                               │
│  Update Helm chart values                     │  ← 15px / weight 400
│  Flagged email                                │
│                                               │
│  ─────────────────────────────────────────── │
│                                               │
│  Read Istio 1.22 changelog                    │  ← 15px / 400 / secondary text
│  Due Apr 18                                   │
│                                               │
├───────────────────────────────────────────────┤
│  Last sync 9:15 AM                 4 active  │  ← 11px / tertiary
└───────────────────────────────────────────────┘
     320×420px, widget surface #000000
     shadow: 0 5px 30px rgba(0,0,0,0.22)
```

Key departures from the cyan spec:

- **No emoji in task rows.** Apple replaces pictographic priority with
  typographic weight. `🔴 🟠 ⚪ 🔵` gone.
- **No left-edge color strip.** Priority encoded in weight + list order.
- **No card shadow.** Only the outer widget frame carries the `3px 5px
  30px / 0.22` diffuse shadow. Cards inside sit flat on the base color.
- **Divider, not card border.** Rows separated by 1px `rgba(255,255,255,0.08)`
  hairlines — the Mail.app / Settings.app pattern.
- **Sync icon is `⟳`, not `🔄`.** Apple never ships emoji in chrome.

### Hover / interaction state

- Card hover: background shifts `#000000 → #1d1d1f → #272729`. No glow,
  no border color. Color delta *is* the affordance.
- Checkbox: replaces the inline checkbox with an SF-style circle → filled
  circle transition on completion. 18px diameter, 1.5px stroke in
  `rgba(255,255,255,0.32)`, becomes filled `#30d158` on check with a
  150ms scale+fade.
- Completion: task row animates `opacity 1.0 → 0.32`, strikethrough via
  `TextDecorations="Strikethrough"`. No fade-out color change.

---

## 5. Title Bar (Glass)

The single place the Apple nav-glass effect lives:

```
Background:       rgba(0, 0, 0, 0.80)
backdrop-filter:  saturate(180%) blur(20px)
Height:           44px
Padding:          0 12px
Divider below:    1px rgba(255,255,255,0.08)
```

WPF caveat: `backdrop-filter` is CSS. WPF's equivalent is
`System.Windows.Media.Effects.BlurEffect` applied to a behind-layer with
`OpacityMask`, or — cleaner — use Windows 11 Mica via `DwmSetWindowAttribute`
with `DWMWA_SYSTEMBACKDROP_TYPE = DWMSBT_MAINWINDOW` (value 2). On Win10,
fall back to the flat `rgba(0,0,0,0.80)` fill — still passes visually.

Title bar content:

- **Wordmark:** `VIGIL`, 17px / 600 / letter-spacing `-0.374px`, color `#ffffff`.
  No `⚡`. The typography carries the brand.
- **Task count badge:** compact 20×20 circle, `#1d1d1f` background, 11px
  white text. Shows `0` never — hide when empty.
- **Sync:** `⟳` in `#2997ff`, 16px, spins 360° over 900ms on click.
- **Collapse:** `—` in `rgba(255,255,255,0.56)`, 16px.
- **Close:** `✕` in `rgba(255,255,255,0.56)`, 16px. Destructive hover →
  `#ff453a`. Apple's traffic-light pattern adapted to one button.

---

## 6. Quick-Add Popup (Ctrl+Win+A)

Reskin of spec §5.4. 440×280px, centered on active monitor.

```
╭──────────────────────────────────────────────╮
│                                              │
│     New task                            ✕    │   21px / weight 400
│                                              │
│   ┌────────────────────────────────────────┐ │
│   │  Review PR for payment service       ▍ │ │   15px input, 8px radius
│   └────────────────────────────────────────┘ │   bg #1d1d1f, no border
│                                              │
│   Priority                                   │   12px label, text.secondary
│   [  Low  ][ Normal ][  High  ][ Critical ] │   segmented control
│                                              │
│   Due                                        │
│   [  None ][ Today ][Tomorrow][ This Week ] │   4 pill buttons, 980px radius
│                                              │
│   ┌────────────────────────────────────────┐ │
│   │  Notes                                 │ │   collapsible, 12px
│   └────────────────────────────────────────┘ │
│                                              │
│                         ( Cancel )  [ Add ] │   Add = filled Apple Blue
│                                              │     Cancel = ghost / text-only
╰──────────────────────────────────────────────╯
        440×280, #000000 base, 12px outer radius,
        shadow 0 20px 60px rgba(0,0,0,0.50)
```

- **Segmented priority control** replaces the cyan row of emoji dots.
  Selected segment = `#2997ff` background + white text. Non-selected =
  `#1d1d1f` + `rgba(255,255,255,0.56)`. This is the iOS/macOS segmented
  control, adapted.
- **Due quick-picks as pills** — the 980px-radius shape from apple.md §5.
  Selected pill fills with `#2997ff`, unselected is outline-only in
  `rgba(255,255,255,0.24)`.
- **Primary button (`Add`):** filled `#2997ff`, white text, 8px radius,
  `8px 15px` padding, 13px weight 500. Apple's primary button verbatim.
- **Secondary (`Cancel`):** text-only, `rgba(255,255,255,0.56)`, no
  background. Apple dismisses tertiary actions by removing their chrome.

Keyboard flow unchanged from spec §5.4 — Apple values keyboard parity.

---

## 7. Collapsed Mode

```
┌──────────────────────────────────────┐
│  VIGIL           3 active       ⌃   │   32px high, 12px radius
└──────────────────────────────────────┘
      #000000 + 0.80 glass, no inner content
```

No color-coded badge (that was a cyan-spec idea). Collapsed state shows
plain count; urgency is communicated by expanding the widget, not by
tinting the bar.

---

## 8. Animations — All Restrained

Apple's motion language is **damped, never bouncy**. Easing curve:
`cubic-bezier(0.28, 0.11, 0.32, 1)` (the iOS "ease-out-expo" feel).

| Effect | Duration | Curve |
|---|---|---|
| Card hover background shift | 180ms | ease-out-expo |
| Completion fade + strikethrough | 240ms | ease-out-expo |
| Popup appear (scale 0.96→1.0 + opacity 0→1) | 220ms | ease-out-expo |
| Sync spin | 900ms | linear, single rotation |
| Task add flash | none | removed |
| Overdue pulse | **none** | removed (see below) |

**Overdue handling without pulse:** pulsing red borders are a cyberpunk
pattern, not an Apple one. Replace with: overdue tasks appear at the
top, their due label switches to `#ff453a` static text, and the parent
widget gains a 1px `rgba(255,69,58,0.32)` outline while any overdue item
exists. One sustained signal, not a heartbeat.

---

## 9. What This Changes in the Plan

Update spec §5 (`UI Design — Premium Futuristic Aesthetic`) to point at
this file. No other section is affected — data model, Outlook COM,
hotkey, lifecycle, and file layout are all skin-agnostic. The reskin is a
pure §5 swap.

Phase 1 deliverable updates:

- Remove emoji from XAML entirely (`⚡`, `🔄`, `▬`, `🔴`, `🟠`, `⚪`, `🔵`, `📅`, `📌`, `📧`).
- Drop the cyan `#00d4ff` token and its glow effect.
- Add Win11 Mica detection for the title bar glass (see §5 above).
- Typography stack changes from monospace to Segoe UI Variable / Segoe UI.

Everything else stays.

---

## 10. Open Question for the User

apple.md §7 says *"Don't introduce additional accent colors — the entire
chromatic budget is spent on blue."* VIGIL's job is to surface urgency at a
glance. A strict reading kills the 4-color priority system.

**Pick one:**

- **A** — Pure Apple. Priority via weight/position only. Cleanest, most
  restrained. Risk: critical and normal look similar in peripheral vision.
- **B** — Apple + one semantic exception. Red (`#ff453a`) allowed for
  critical/overdue only. Everything else neutral.
- **C** — Skin toggle. Ship both and let the user choose in settings.

Recommend **B** for the default. It is the same exception Apple itself
makes (destructive actions in System Settings, battery low, notification
badges). One semantic red, one interactive blue. Everything else in the
monochrome palette. This is "Apple" in spirit without being naive about
the widget's actual job.
