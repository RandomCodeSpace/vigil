# Security Policy

## Supported versions

Vigil is distributed as PowerShell source, not a versioned binary. Security fixes land on `main` and are tagged when material — there is no formal LTS line.

| Tracking | Status |
|---|---|
| `main` | Supported (current) |
| Older tagged commits | Best-effort; rebase onto `main` for fixes |

If you are running Vigil from a clone, `git pull` on `main` is the patch channel.

## Reporting a vulnerability

Please **do not open a public GitHub issue** for security problems.

Use one of:

- **GitHub private vulnerability report** — preferred. Open `https://github.com/RandomCodeSpace/vigil/security/advisories/new` (you must be signed in to GitHub). The advisory channel is monitored by the maintainer.
- **Email** — `ak.nitrr13@gmail.com`. Put `[vigil security]` in the subject so the report is triaged ahead of normal mail.

Please include:

- The Vigil commit SHA (`git rev-parse HEAD`) or the line of the script affected.
- The PowerShell host + version (`$PSVersionTable | Out-String`).
- The shortest reproducer you can produce — the `pwsh -File` invocation plus any test input.
- Your assessment of impact (e.g., DPAPI key exposure, COM hijack, path traversal in the task store, info-disclosure via logs).
- Whether the issue is in Vigil's code or in a runtime / framework dependency (.NET, pwsh, Outlook COM), and the upstream advisory ID if known.

## What you can expect

- **Acknowledgement** within 72 hours.
- **Initial triage** within 7 days, with a severity rating (CVSS v3.1) and an indicative remediation timeline.
- **Coordinated disclosure** — we will agree on a public-disclosure date with the reporter; default is 90 days from triage, sooner for low-impact / already-public issues.
- **Credit** in the GHSA advisory and the README changelog (unless the reporter requests anonymity).

There is no paid bug bounty.

## Scope

In-scope:

- The `VIGIL.ps1`, `preflight.ps1`, and `Test-Vigil.ps1` scripts.
- The DPAPI-wrapped task store layout under `.vigil/` (key handling, atomic write contract, path-traversal hardening).
- The Outlook COM integration (RCW lifecycle, sort-before-restrict invariant, sync window bounding).
- The auto-start shortcut + tray icon flows (privilege escalation, hijack via writable shortcut paths).
- The single-instance mutex (`Global\VIGIL_TaskTracker`) and the P/Invoke window-reactivation surface.

Out of scope:

- Vulnerabilities that require pre-existing local code execution on the user's machine. Vigil is a single-user desktop tool — by definition you trust the code you run as yourself.
- Public-internet attack surface — Vigil does not bind a network socket and does not phone home. Deploying it behind hostile reverse proxies is out of scope because that deployment shape is unsupported.
- Findings in third-party services we do not control (Microsoft Outlook, Windows DPAPI, GitHub itself) — please report those upstream.
- Constrained Language Mode / AppLocker / AMSI-EDR interactions where the corporate policy is intentionally blocking script execution. `preflight.ps1` reports those as known blockers.

## Hardening references

- [`CLAUDE.md`](CLAUDE.md) §7 "Security" and §10 "Supply-chain observability" — CVE policy + OSS-CLI stack.
- [`.github/workflows/scorecard.yml`](.github/workflows/scorecard.yml) — OpenSSF Scorecard supply-chain checks.
- [`.github/workflows/security.yml`](.github/workflows/security.yml) — Semgrep / OSV-Scanner / Trivy / Gitleaks / jscpd / SBOM.
- [`.github/dependabot.yml`](.github/dependabot.yml) — automated GitHub Actions bumps.
- GitHub repo settings — secret scanning + push protection are expected to be **enabled** on the repo (board-owned action item if not yet).

## Branch protection (board action item)

Per [`CLAUDE.md`](CLAUDE.md) §6, `main` is expected to require:

- Signed commits
- Pull request before merging
- Status checks (Scorecard + security workflow + Test-Vigil) to pass
- Linear history (squash merges only)
- Force-push and deletion blocked

These are GitHub repo Settings → Branches and not enforceable from a workflow file. Tracked alongside [RAN-55](https://github.com/RandomCodeSpace/vigil/issues) until the toggles are flipped.

## Changelog

This file is versioned as part of the repo. Material changes (e.g., changing the disclosure timeline, raising the supported-versions table) are announced via a Release note when a tag is cut and a Paperclip board comment.
