# Contributing to document-viewer

Curiosity is a good starting point. Small fixes, experiments, documentation, accessibility checks and reproducible bug reports are welcome. Vibe Coding and AI-assisted contributions are explicitly welcome: use any tools that help you explore. We review the result, not the tool.

歡迎 Vibe Coding 與 AI 輔助貢獻，不限制工具。請理解並檢查自己的修改，提供可重現的驗證；不用提交私人提示詞或聊天紀錄。探索尚未解決的問題也有價值。

## Start here

1. Read [README](README.md), [AGENTS.md](AGENTS.md) and nearby tests. Search existing issues before filing another one.
2. Fork this public repository and clone your fork. Create a branch for one focused change. No access to private website repositories or production credentials is needed.
3. Use Node.js 24, npm and the committed lockfile:

```sh
npm ci
npx playwright install chromium
npm run build
npm test
git diff --check
```

On Linux use `npx playwright install --with-deps chromium` if browser libraries are missing. Builds and dry-runs do not publish. Never run deploy commands as routine contribution checks.

## Find your way around

`src/renderers/`: format renderers; `src/ppt/`: project-owned parser; `src/sheets/`: workbook Worker; `tools/`: local asset graph; `tests/`: parser/browser fixtures.

Keep this browser-only and read-only. Preserve script-disabled frames, local dependencies, consent, continuous pages and offscreen resource reclamation. Do not execute macros or document scripts, compile formulas into JavaScript, or add conversion servers, download controls or proprietary/watermarked engines. Formula preview changes need bounded interpreter tests against documented Calc semantics; prefer saved results and discard calculation caches with each window. Explain renderer and initial-memory boundaries honestly. Optional operator samples must not be required by normal CI.

To check your own non-private Office/PDF examples locally, set `DOCUMENT_VIEWER_SAMPLES` to a sample folder and optionally `DOCUMENT_VIEWER_PDF` to a PDF path before running `npm test`. `DOCUMENT_VIEWER_SCREENSHOTS` optionally writes preview evidence to a local folder. These files are never uploaded by tests and must not be committed automatically.

Generated output, dependencies and temporary evidence are not source edits. Preserve APIs, content, languages, licenses and attribution. Add regression tests beside existing ones; do not weaken an assertion to conceal a failure.

## From idea to PR

- Report bugs with a minimal sanitized reproduction, expected/actual behavior and environment. Never attach private documents or secrets.
- Discuss larger directions with a proposal: the question, alternatives and an observable success condition. Small fixes need no proposal first. Uncertain experiments can be draft PRs.
- Keep diffs reviewable. Explain why a change helps and its compatibility/security/accessibility effects. List exact checks and results; say what remains unrun. AI-generated claims are not verification.
- Human or AI-written code has the same expectations: understand the diff, check provenance/licenses and provide reproducible evidence. Naming an AI tool or sharing prompts is optional.
- Open a PR against `main` using the checklist. CI checks untrusted PRs without production secrets or deployment. Respond to review; maintainers merge when changes and evidence are ready. No response-time guarantee is implied.

Be kind to people who are learning. Prefer specific questions and constructive evidence to gatekeeping. Vulnerabilities belong in [SECURITY.md](SECURITY.md)'s private reporting channel. Project code uses [LICENSE](LICENSE); third-party code/assets keep their own terms.
