# AGENTS.md (for Codex)

## Project
Electron + TypeScript app for exhibition estimation (積算システム).

## Do / Don't
- Do not add node_modules/ or dist/ to git.
- Keep changes small and focused.
- Prefer TypeScript. Avoid large refactors unless requested.

## Commands
- Install: npm ci
- Dev: npm run dev (if available)
- Build: npm run build (if available)
- Lint/Test: npm run lint / npm test (if available)

## Files
- Main process: src/main/
- Renderer: src/renderer/
- Prompts: prompts/

## Verification (Local Mac only)
Codex is running locally via Codex CLI, so UI launch is allowed.

MUST run:
- npm ci
- npm run build
- npm start (smoke test: app launches without crash, open window at least once)

MUST verify in UI:
- Import a v1.1.2 sample JSON and confirm category tabs render
- Edit siteCosts and confirm currency is always present in JSON (JPY by default)

MUST NOT:
- Add node_modules/ or dist/ to git

## Reporting
After verification, Codex MUST report:
- npm ci / npm run build result (success or error)
- npm start result (window opened or error)
- Any deviation from expected UI behavior