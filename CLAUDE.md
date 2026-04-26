# CLAUDE.md — Vigil

> Architecture + conventions SSoT for code-touching changes on Vigil. Per-issue specifics live in the Paperclip thread; runbooks live in `shared/runbooks/` (currently inherited from the codeiq sibling project — see References).
>
> The rule of last resort: **`/home/dev/.claude/rules/*.md` wins**. This file does not contradict it; it specialises it for Vigil.

---

## 1. What Vigil is

Vigil is a **Windows-first personal task command center** — a single-file PowerShell + WPF desktop app that unifies Outlook flagged emails / calendar items / open tasks with a manual task list, behind a chromeless Fluent (Mica) WPF window. Single-user, MIT-licensed, no compiled binary distribution.

The product target is one knowledge worker on their own machine. There is no SaaS surface, no shared backend, no public-internet attack surface. Distribution is the GitHub repo plus a local `pwsh -File .\VIGIL.ps1` invocation.

## 2. Stack identification (RAN-55 prereq)

| Layer | Choice |
|---|---|
| Primary language | **PowerShell 7.5+** (`pwsh`) — cross-platform core; `Test-Vigil.ps1` runs on Linux / macOS pwsh too |
| Legacy fallback | Windows PowerShell 5.1 — supported for the data layer; Fluent theme degrades |
| Runtime | **.NET 9** (PowerShell 7.5 is bundled on top of .NET 9) |
| UI | **WPF / Fluent / Mica** (Windows 11 visuals); guarded for Windows-only |
| Embedded C# | `Add-Type` inline P/Invoke for window reactivation + tray hotkeys |
| External integration | **Outlook COM** (15-minute sync); flagged emails / 24h calendar / open tasks |
| Storage | DPAPI-wrapped `tasks.json` (CurrentUser scope) when BitLocker is off |
| Tests | Pester-style asserts; `Test-Vigil.ps1` dot-sources VIGIL with `-NoUI` |
| Package manager | **None** — no `package-lock.json`, no `pom.xml`, no `requirements.txt`. Dependencies are framework-bundled (.NET / pwsh built-ins) |

What this means for tooling:

- **No language ecosystem dependency surface** to scan with OSV-Scanner / Dependabot package updaters. The supply-chain attack surface is GitHub Actions + the .NET / pwsh runtimes themselves.
- **The (B) OSS-CLI security stack is language-adapted** (see §4) — Semgrep keeps the secret-audit + GHA / YAML rules, OSV-Scanner runs as a no-op on the source tree (will turn into useful coverage if a `*.lock` is ever added), Trivy / Gitleaks / SBOM / jscpd are language-agnostic.
- **There is no compile / build step**. "CI" for Vigil is `pwsh -NoProfile -File ./Test-Vigil.ps1` plus the security workflow.

## 3. Repository shape

```
VIGIL.ps1            — main app (≈99 KB). 5-phase startup:
                       preflight → quick-add popup → Outlook COM 15-min sync →
                       auto-start shortcut → WPF UI + tray.
                       Params: -NoUI, -IncludeCalendar.
preflight.ps1        — 60 environment checks (runtime, corp lockdown CLM /
                       AppLocker / AMSI, Azure AD, hardware). Emits a compact
                       bitmap `VIGIL:v1:<count>:<hex>:P<n>:F<n>:T<tenant>`.
Test-Vigil.ps1       — dot-sources VIGIL with -NoUI; custom asserts; runs on
                       PS 5.1 and pwsh 6+. ≈116 cross-platform unit tests for
                       the data layer.
.vigil/              — runtime state (DPAPI-wrapped tasks.json, logs).
                       Colocated next to the script; fallback ~/.vigil;
                       legacy migration from the old userprofile path.
LICENSE              — MIT (Amit Kumar).
README.md            — user-facing run instructions + OpenSSF badge.
SECURITY.md          — vulnerability disclosure policy.
AGENTS.md            — agent collaborator entry-point.
.bestpractices.json  — OpenSSF Best Practices self-assessment (project 12648).
.github/             — workflows (scorecard, security), Dependabot config.
```

## 4. Quality gates

| Gate | Threshold | Where it runs | Failure action |
|---|---|---|---|
| Unit tests | All pass | `pwsh -NoProfile -File ./Test-Vigil.ps1` | Block merge |
| Signed commits | Every commit on `main` verifies | Branch protection + `gh api ... /commits/{sha}/check-runs` | Block merge |
| OpenSSF Scorecard | Best-effort; soft target ≥ 8.0/10. `Pinned-Dependencies` is the priority | `.github/workflows/scorecard.yml` (push to `main` + Mondays 06:00 UTC) | Surface in Security tab; do **not** gate merge |
| OpenSSF Best Practices | `passing` badge on https://www.bestpractices.dev/en/projects/12648 | bestpractices.dev admin UI | The hard gate per board. Tracked under [RAN-55](/RAN/issues/RAN-55) |
| OSS-CLI observability | Findings surfaced via SARIF + artifacts | `.github/workflows/security.yml` | Observation-only at bootstrap; promote to gate-blocking after a clean baseline |

The OSS-CLI stack mirrors the codeiq sibling — Semgrep (SAST), OSV-Scanner (deps; mostly no-op for the PS tree), Trivy (filesystem CVEs + IaC misconfig on the GHA YAML), Gitleaks (secret scan over full history), jscpd (PowerShell duplication via tokenization), and `anchore/sbom-action` (SPDX SBOM via Syft). Cron: Mondays 06:00 UTC, same window as the Scorecard sweep. All third-party actions pinned by commit SHA.

## 5. Code style

- **PowerShell 7.5 first**. PS 5.1 compatibility is preserved for the data layer — no `?.` operator, no ternary, no pipeline `?` / `??` outside guarded blocks. Where a 7-only construct is justified, gate it with `if ($PSVersionTable.PSVersion.Major -ge 7)`.
- **Cross-platform core**. Anything that is not WPF / Outlook COM must run on Linux / macOS pwsh — the test suite enforces this. WPF and `System.Windows.*` types are loaded behind `if ($IsWindows)`.
- **`System.Drawing` is lazy-loaded**. Linux pwsh lacks libgdiplus; do not import at script top.
- **Atomic writes** via `[System.IO.File]::Replace`, never `Move-Item -Force`.
- **Outlook COM hygiene**: `Sort()` **before** `Restrict()` for correct flagged counts; reverse-order `ReleaseComObject` + forced GC after each session to prevent RCW leaks.
- **Single-instance**: `Global\VIGIL_TaskTracker` named mutex + P/Invoke window reactivation. Do not introduce a second process model.
- **Logs**: 500-line rotation. Don't rely on file growth being unbounded.
- **Reduce-motion** variant strips Storyboards; honour the user's preference.

## 6. Branch / commit / PR rules

- Branch off `main`. Conventional-commit subjects (`feat:`, `fix:`, `chore:`, `refactor:`, `test:`, `docs:`, `perf:`).
- One logical change per commit. Squash-merge into `main`.
- Every commit on `main` must be signed (ssh or GPG) — branch protection rejects unsigned commits. **Action item for the board:** enable "Require signed commits" + "Require pull request before merging" + "Require status checks" on `main` via repo Settings → Branches.
- PR title = conventional-commit subject. Body links the Paperclip issue (`Closes RAN-XX`), states "why" in 1–2 sentences, and notes any rollback considerations.
- No force-push to `main` ever.

## 7. Security

- **Inputs** — every Outlook field that hits the UI is treated as untrusted text; manual task input is escaped before being interpolated into XAML.
- **Path traversal** — task store path is canonicalized via `[System.IO.Path]::GetFullPath` and asserted to live under `.vigil/` or `~/.vigil`.
- **Secrets** — never in code, config, or commit history. The DPAPI key is per-user and never exfiltrated. CI secrets are repo-level only.
- **CVE policy** — High/Critical → block the merge; Medium → fix if a patched version exists, else document non-exploitability with TechLead sign-off; Low → tracked in the next bump cycle.
- **Vulnerability reporting** — see [`SECURITY.md`](SECURITY.md). Private disclosure only.
- **OSS-CLI observability stack** — `.github/workflows/security.yml` runs Semgrep / OSV-Scanner / Trivy / Gitleaks / jscpd / `anchore/sbom-action` on every PR + push to `main` + weekly cron. OpenSSF Scorecard runs in `.github/workflows/scorecard.yml`.

## 8. Performance

- Outlook COM sync is bounded to 15 minutes and runs on a separate runspace; never block the dispatcher thread.
- WPF binding updates batch through the dispatcher — no per-item invocations from inside a tight loop.
- Search input is debounced; the debounce is flushed on window close so the last text is never lost.
- 60 fps target for animations on Windows 11; Storyboards are removed entirely under `prefers-reduced-motion`-equivalent (custom flag).

## 9. Build & distribution

- **There is no build step.** Distribution is `git clone` + `pwsh .\VIGIL.ps1`.
- GitHub Actions are pinned by **commit SHA** in every workflow (Scorecard `Pinned-Dependencies`).
- No public-CDN runtime fetches, no auto-update phone-home, no telemetry default-on.
- No container image yet. If one is added, it must build from a minimal base (distroless / scratch / Alpine), pinned by digest, pushed to GHCR with provenance attestations.

## 10. Supply-chain observability (OpenSSF) — RAN-55

| Surface | Status |
|---|---|
| OpenSSF Best Practices project page | https://www.bestpractices.dev/en/projects/12648 — `in_progress` → **target `passing`** (admin-UI flip is board-owned) |
| Best Practices self-assessment | [`.bestpractices.json`](.bestpractices.json) — `level: passing`, evidence pointers per category |
| Scorecard workflow | [`.github/workflows/scorecard.yml`](.github/workflows/scorecard.yml) — push to `main` + Mondays 06:00 UTC; SARIF → Security tab + artifact |
| OSS-CLI security workflow | [`.github/workflows/security.yml`](.github/workflows/security.yml) — Semgrep / OSV / Trivy / Gitleaks / jscpd / SBOM; PR + push + weekly cron |
| Dependabot | [`.github/dependabot.yml`](.github/dependabot.yml) — `github-actions` ecosystem only (no Maven / npm; Vigil has no package manager) |
| README badge | [![OpenSSF Best Practices](https://www.bestpractices.dev/projects/12648/badge)](https://www.bestpractices.dev/projects/12648) |

### Scorecard baseline + target

- **Baseline (first run after merge):** to be captured in the Security tab once the workflow lands. Comment with the snapshot here on first publish.
- **Stretch target:** **≥ 8.0 / 10** aggregate. Per the board, Scorecard is observational and does **not** gate merge — `passing` Best Practices is the only hard gate.
- **Priority checks:** `Pinned-Dependencies` (we already SHA-pin all actions), `Branch-Protection` (board to enable), `Token-Permissions` (workflows opt into the narrowest scopes), `Dangerous-Workflow` (no `pull_request_target` writes), `SAST` (Semgrep covers it).
- **Known low-scorers:** `Signed-Releases` is N/A until we ship a tagged release; `CII-Best-Practices` flips green once the bestpractices.dev project is at `passing`; `Webhooks` and `License` should already be green.

## 11. Documentation

- This file (`CLAUDE.md`) is the architecture + conventions SSoT.
- [`AGENTS.md`](AGENTS.md) is the entry-point for agent collaborators.
- [`SECURITY.md`](SECURITY.md) is the disclosure policy.
- [`.bestpractices.json`](.bestpractices.json) is the OpenSSF Best Practices self-assessment.

## 12. References

- `LICENSE` — MIT.
- `SECURITY.md` — disclosure policy.
- `.github/workflows/` — Scorecard + OSS-CLI security automations.
- `/home/dev/.claude/rules/*.md` — global engineering rules (parent SSoT).
- `RAN-50` — parent issue (OpenSSF rollout across the 5 paperclip projects).
- `RAN-55` — this issue (Vigil-specific landing).
- codeiq sibling (`/home/dev/projects/codeiq`) — `RAN-46` / `RAN-52` precedent.
