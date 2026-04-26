# Vigil — Documentation

Documentation index for the Vigil project. The canonical "what / why / how" lives at the repo root in [`README.md`](../README.md); this folder holds the structured docs that the OpenSSF Best Practices `documentation_basics` criterion expects.

## Contents

- [Architecture](architecture.md) — 5-phase startup, repository shape, file-by-file responsibilities, runtime invariants.
- [Install & run](install.md) — system requirements, supported PowerShell hosts, first-run side effects (`.vigil/` path, auto-start shortcut, tray icon), `-NoUI` and `-IncludeCalendar` flags.
- [Troubleshooting](troubleshooting.md) — common environment blockers (Constrained Language Mode, AppLocker / AMSI / EDR, Outlook COM auth), preflight bitmap decoding, log file location.
- [Security model](security.md) — DPAPI scope, threat model, atomic-write contract, vulnerability disclosure pointer.

## See also

- [`README.md`](../README.md) — user-facing landing page (features, quickstart).
- [`CLAUDE.md`](../CLAUDE.md) — architecture + conventions SSoT for code-touching changes.
- [`AGENTS.md`](../AGENTS.md) — entry point for agent collaborators.
- [`SECURITY.md`](../SECURITY.md) — vulnerability disclosure policy.
- [`CHANGELOG.md`](../CHANGELOG.md) — release notes / `Unreleased` log.
- [`.bestpractices.json`](../.bestpractices.json) — OpenSSF Best Practices self-assessment.
