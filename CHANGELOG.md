# Changelog

All notable changes to Vigil are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html) once a tagged release line exists.

While Vigil remains pre-1.0 and distributed as PowerShell source from `git clone`, every commit on `main` is the canonical version identifier (`git rev-parse HEAD`). The "Unreleased" section below tracks what has landed on `main` since the last tag.

## [Unreleased]

### Added

- **OpenSSF Best Practices + Scorecard scaffolding** ([RAN-55], [RAN-60]).
  - `.github/workflows/scorecard.yml` — OpenSSF Scorecard supply-chain analysis (push to `main` + Mondays 06:00 UTC, SHA-pinned actions, SARIF + artifact).
  - `.github/workflows/security.yml` — consolidated OSS-CLI stack (Semgrep, OSV-Scanner via binary install, Trivy filesystem, Gitleaks, jscpd, anchore/sbom-action). PR + push + weekly cron.
  - `.github/dependabot.yml` — `github-actions` ecosystem only (Vigil ships no language lockfile).
  - `.bestpractices.json` — canonical autofill schema, `project_id: 12648`, `level: passing`, per-criterion `*_status` + `*_justification` fields.
  - `CLAUDE.md` — architecture + conventions SSoT, OpenSSF observability target, Scorecard policy.
  - `SECURITY.md` — private-disclosure policy, scope, hardening references.
  - `AGENTS.md` — agent-collaborator entry point.
  - `README.md` — OpenSSF Best Practices + Scorecard badges at the top.
- **`docs/` folder** ([RAN-55]) — documentation basics (architecture, run, troubleshooting, security model) for the OpenSSF `documentation_basics` criterion.
- **`CHANGELOG.md`** — this file ([RAN-55]).

### Fixed

- **OSV-Scanner CI** ([RAN-55]) — replaced the broken `google/osv-scanner-action@v2.3.5` (composite `action.yml` missing the top-level `runs:` block) with a `gh release download` binary install. Mirrors the codeiq fix; coverage activates automatically once a `*.lock` lands in-tree.
- **Search debounce on close** — pending search-input debounce is now flushed on window close, so the last text typed is never lost.
- **Deep-review findings** — fixes across `VIGIL.ps1`, `preflight.ps1`, `Test-Vigil.ps1` (DPAPI key path edge cases, atomic-write contract, Outlook RCW lifecycle, log rotation timing, single-instance mutex hand-off).

### Changed

- `LICENSE` — copyright attributed to `Amit Kumar` (matches the project lineage / sibling-repo precedent).

### Security

- Adopted the (B) OSS-CLI security stack as the project's continuous supply-chain observability surface. High/Critical findings are merge gates per `CLAUDE.md` §7. SARIF results land in the GitHub Security tab where supported and are uploaded as workflow artifacts regardless.
- Branch protection on `main` (signed commits, required PR review, required status checks) and repo-level secret scanning + push protection are board-owned toggles tracked alongside [RAN-55] until enabled.

[Unreleased]: https://github.com/RandomCodeSpace/vigil/commits/main
[RAN-55]: https://github.com/RandomCodeSpace/vigil/issues
[RAN-60]: https://github.com/RandomCodeSpace/vigil/issues
