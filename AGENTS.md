# AGENTS.md — Vigil

Entry-point for agent collaborators working on this repo.

## Read these first

1. [`CLAUDE.md`](CLAUDE.md) — architecture, conventions, quality gates, supply-chain observability.
2. `/home/dev/.claude/rules/*.md` — global engineering rules (parent SSoT). They win on conflict with everything in this repo.
3. [`SECURITY.md`](SECURITY.md) — disclosure policy and scope.
4. [`.bestpractices.json`](.bestpractices.json) — OpenSSF Best Practices self-assessment.

## What Vigil is, in 30 seconds

A Windows-first single-file PowerShell + WPF task command center. No build step, no package manager, no SaaS surface. `VIGIL.ps1` is the app, `preflight.ps1` is the environment-check helper, `Test-Vigil.ps1` is the cross-platform unit-test harness. Distribution is `git clone` + `pwsh -File .\VIGIL.ps1`.

See [`CLAUDE.md`](CLAUDE.md) §1–§3 for the full architecture and stack.

## Workflow

- Branch off `main` (`feat/`, `fix/`, `chore/`, `docs/` prefix).
- Conventional-commit subjects. One logical change per commit. Squash-merge into `main`.
- Sign every commit on `main` (ssh or GPG). Branch protection rejects unsigned commits — turn that on before merging if it is not already.
- Run `pwsh -NoProfile -File ./Test-Vigil.ps1` locally before opening a PR.
- Open a PR with `Closes RAN-XX` linking the Paperclip issue.

## CI surfaces

- [`.github/workflows/scorecard.yml`](.github/workflows/scorecard.yml) — OpenSSF Scorecard (push to `main` + Mondays 06:00 UTC).
- [`.github/workflows/security.yml`](.github/workflows/security.yml) — Semgrep + OSV-Scanner + Trivy + Gitleaks + jscpd + SBOM (PR + push + weekly cron).
- [`.github/dependabot.yml`](.github/dependabot.yml) — `github-actions` bumps only (no Maven / npm; Vigil has no package manager).

All third-party actions are pinned by **commit SHA** (Scorecard `Pinned-Dependencies`). When you bump an action, update the comment above the `uses:` line with the new tag.

## What not to touch silently

- The 5-phase startup order in `VIGIL.ps1` (preflight → quick-add → Outlook sync → auto-start shortcut → WPF UI + tray).
- Outlook COM `Sort()` **before** `Restrict()` — flips break flagged-email counts.
- DPAPI wrap of `tasks.json` (CurrentUser scope, BitLocker-off branch).
- Atomic-write contract (`[System.IO.File]::Replace`, never `Move-Item -Force`).
- Single-instance `Global\VIGIL_TaskTracker` mutex + window reactivation P/Invoke.
- 500-line log rotation.
- Reduce-motion variant (strip Storyboards, do not re-introduce them).

These are tested invariants — if you have to touch one, call it out in the PR body and add a regression test.

## Reporting

- Code-touching issues / progress → Paperclip thread on the parent issue.
- Security findings → [`SECURITY.md`](SECURITY.md) — never a public GitHub issue.
