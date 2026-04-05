# @officekit/word

Lane-3 parity-first scaffold for the Word adapter.

## Current contents

- `src/manifest.js` — source lineage, capability families, deferred shared-core contracts, implementation milestones, and parity risks
- `src/index.js` — package exports for manifest/contract helpers
- `test/manifest.test.js` — Bun-backed invariants for parity-critical surface and path/preview expectations

## Source lineage

Primary source families:

- `OfficeCLI/src/officecli/Handlers/WordHandler.cs`
- `OfficeCLI/src/officecli/Handlers/Word/WordHandler.Add*.cs`
- `OfficeCLI/src/officecli/Handlers/Word/WordHandler.Query.cs`
- `OfficeCLI/src/officecli/Handlers/Word/WordHandler.Set*.cs`
- `OfficeCLI/src/officecli/Handlers/Word/WordHandler.View.cs`
- `OfficeCLI/src/officecli/Handlers/Word/WordHandler.FormFields.cs`
- `OfficeCLI/src/officecli/Handlers/Word/WordHandler.HtmlPreview*.cs`

## Parity-critical families

- structure and navigation
- styles, doc defaults, compatibility settings
- mutation flows (`add`, `set`, `remove`, `move`, `swap`, `copy-from`)
- raw XML + batch operations
- HTML preview
- forms extraction

## Deferred shared-core decisions

- selector/path grammar
- shared result envelope
- preview server orchestration
- filesystem helpers
