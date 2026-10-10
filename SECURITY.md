# Security

## Report privately

Do not publish exploit details, secrets, personal documents or production data in issues or PRs. Use this repository's private vulnerability reporting option if available, or email **helper@js.gripe** with the subject "Security: document-viewer".

Include the affected commit/version, platform, impact, a minimal sanitized reproduction and possible mitigation. Use dummy accounts/data, not live credentials. We can agree on a private way to exchange a sensitive fixture if needed.

## Scope and maintenance

Reports about current `main` source and reproducible regressions are welcome. There is no promised response SLA, automatic backport policy or security certification. Operators should review changes and dependency advisories before deploying. Vendor platform/provider issues belong to their vendor.

Keep this browser-only and read-only. Preserve script-disabled frames, local dependencies, consent, continuous pages and offscreen resource reclamation. Do not execute macros/formulas or add conversion servers, download controls or proprietary/watermarked engines. Explain renderer and initial-memory boundaries honestly. Optional operator samples must not be required by normal CI.

Ordinary bugs/ideas can use public issues. Coordinate disclosure after a fix and verification. Do not test someone else's live service without authorization.

The legacy DOC renderer pins DOMPurify to 3.4.15. This project's scoped npm override installs the patched 3.4.16 release. npm does not apply overrides from dependencies: integrating applications must include the same `@file-viewer/doc` override in their root manifest and refresh their lockfile until the upstream renderer updates its pin. Keep the script-disabled document frame and local sanitization enabled independently of this dependency fix.
