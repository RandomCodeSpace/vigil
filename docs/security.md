# Security model

Vigil is a single-user desktop tool. The threat model assumes the user trusts the code they run as themselves.

## Threat model summary

**In scope:**

- Path traversal in the local task store (`.vigil/`).
- Command / XAML injection via Outlook fields or manual task input rendered by the WPF binding layer.
- DPAPI key handling (CurrentUser scope, BitLocker-off branch).
- Outlook COM RCW lifecycle (resource leak → privilege-escalation surface if a hostile add-in is loaded).
- Auto-start shortcut + tray-icon flows (privilege escalation through writable-shortcut hijack).
- Single-instance mutex hijack (`Global\VIGIL_TaskTracker`).

**Out of scope:**

- Pre-existing local code execution as the same user. Vigil is a single-user desktop tool; you trust the code you run.
- Public-internet attack surface. Vigil does not bind a network socket and does not phone home.
- Third-party services we do not control (Microsoft Outlook, Windows DPAPI, GitHub itself) — please report those upstream.
- Constrained Language Mode / AppLocker / AMSI-EDR interactions where the corporate policy is intentionally blocking script execution. `preflight.ps1` reports those as known blockers.

## Hardened invariants

Per [`CLAUDE.md`](../CLAUDE.md) §7 and enforced by `Test-Vigil.ps1`:

- **Inputs** — every Outlook field that hits the UI is treated as untrusted text; manual task input is escaped before being interpolated into XAML.
- **Path canonicalisation** — task-store path is canonicalised via `[System.IO.Path]::GetFullPath` and asserted to live under `.vigil/` or `~/.vigil`.
- **Atomic writes** — `[System.IO.File]::Replace` only; never `Move-Item -Force`.
- **No secrets in code / config / logs** — Gitleaks runs on full git history (.github/workflows/security.yml).
- **CVE policy** — High/Critical = block merge; Medium = fix-or-document with TechLead sign-off; Low = next bump cycle.

## Crypto

Vigil's only crypto dependency is **Windows DPAPI** (Data Protection API). Used to wrap `tasks.json` at CurrentUser scope.

- Algorithm: AES-256 + HMAC-SHA-512 (Windows 10+).
- Key handling: per-user OS-managed master key. Never exfiltrated, never leaves the machine.
- No proprietary crypto is implemented in Vigil. No MD5 / SHA-1 for integrity, no DES / 3DES, no ECB mode, no hardcoded IVs / keys.
- See [`.bestpractices.json`](../.bestpractices.json) `crypto_*` criteria for the full per-criterion rationale.

## Distribution integrity

- Source distributed exclusively over HTTPS via `git clone`.
- Every commit on `main` is ssh-signed (board-owned branch protection toggle); verify via `git verify-commit <sha>`.
- No GitHub Releases / signed tarballs yet — the commit SHA on `main` is the canonical version identifier.

## Reporting a vulnerability

**Do not open a public GitHub issue for security problems.** See [`SECURITY.md`](../SECURITY.md) for the private disclosure channels (GitHub Security Advisory + maintainer email) and the response SLAs.

## OpenSSF supply-chain observability

- [`.github/workflows/scorecard.yml`](../.github/workflows/scorecard.yml) — OpenSSF Scorecard (push to `main` + Mondays 06:00 UTC).
- [`.github/workflows/security.yml`](../.github/workflows/security.yml) — Semgrep, OSV-Scanner, Trivy, Gitleaks, jscpd, Syft SBOM.
- [`.github/dependabot.yml`](../.github/dependabot.yml) — `github-actions` ecosystem only.
- [`.bestpractices.json`](../.bestpractices.json) — OpenSSF Best Practices self-assessment, `project_id: 12648`.

## See also

- [`SECURITY.md`](../SECURITY.md) — disclosure policy + scope.
- [`CLAUDE.md`](../CLAUDE.md) §7 — security gates.
- [`docs/architecture.md`](architecture.md) — runtime invariants the security model relies on.
