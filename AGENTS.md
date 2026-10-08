# Document viewer development

- Keep this a reusable browser-only library. No LibreOffice, conversion servers, accounts, uploads, analytics, CDN runtime or browser persistent storage.
- Open previews automatically by default without moving focus; callers can set autoOpen: false for explicit activation. Abort and dispose on close. Respect caller consent and CSP before any request. Do not expose original-download links or legacy PPT warning notes in the viewer UI.
- Isolate document-generated HTML in a script-disabled sandbox. Block external resources; never execute macros, formulas or document scripts.
- Bundle dependencies, workers, WASM, fonts and original licenses locally. Fingerprint the dependency graph in integrating sites.
- Legacy PPT parsing and rendering are project-owned implementations based on Microsoft's public MS-PPT/MS-ODRAW specifications. Do not add proprietary or watermarked engines or copy their source.
- Keep fixtures and tests under tests. Support keyboard, narrow screens, English and both Chinese locales.
- Controls follow arbitrary host palettes and fonts, not a specific theme or browser-native selectors. Respect explicit CSS overrides, infer computed host colors when semantic variables are absent, synchronize open sandbox colors on theme changes, and preserve document page colors. Validate keyboard tabs, sandbox Escape, forced colors and contrast in multiple palettes.
- Run browser tests only with isolated Playwright Chromium and extensions disabled. Never use the system Edge/Chrome channel; download-manager integration can launch host dialogs during fixture requests.
- Keep this directory as the canonical checkout. Before editing or pushing, fetch origin, inspect upstream changes and preserve local edits in a backup outside the checkout before resolving conflicts. Never force-push or create duplicate task checkouts.
- Before pushing, run npm ci and npm test. Commit source, tests, lockfile and notices, not node_modules or dist. EdgePress consumes a published full GitHub commit; update its dependency and lockfile only after that commit is available remotely.
