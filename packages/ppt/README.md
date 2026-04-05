# @officekit/ppt

Lane-3 parity-first scaffold for the PowerPoint adapter.

## Current contents

- `src/manifest.js` — source lineage, capability families, deferred shared-core contracts, implementation milestones, and parity risks
- `src/index.js` — package exports for manifest/contract helpers
- `test/manifest.test.js` — Bun-backed invariants for slide/preview/overflow parity expectations

## Source lineage

Primary source families:

- `OfficeCLI/src/officecli/Handlers/PowerPointHandler.cs`
- `OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.Add*.cs`
- `OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.Query.cs`
- `OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.Set*.cs`
- `OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.View.cs`
- `OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.ShapeProperties.cs`
- `OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.SvgPreview.cs`
- `OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.HtmlPreview*.cs`
- `OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.Theme.cs`

## Parity-critical families

- slide/layout/placeholder semantics
- shapes, tables, text, media, hyperlinks
- charts and theme inheritance
- HTML and SVG preview
- overflow checks
- mutation flows and raw XML

## Deferred shared-core decisions

- selector/path grammar
- shared result envelope
- shared preview server
- cross-format query helpers
