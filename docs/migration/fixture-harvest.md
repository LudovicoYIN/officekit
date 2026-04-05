# OfficeCLI fixture and example harvest plan (lane 1)

This file curates the OfficeCLI artifacts that should seed `officekit` parity tests, smoke tests, and docs validation.

## 1) README / product-surface flows to preserve

### Agent onboarding
Source: `OfficeCLI/README.md`
- Skill bootstrap via `curl -fsSL https://officecli.ai/SKILL.md`
- Agent-facing promise: create/read/modify Office docs immediately after install
- Audience split: AI agents, humans, developers

### Developer preview workflow
Source: `OfficeCLI/README.md`
1. install
2. `create deck.pptx`
3. `watch deck.pptx --port 26315`
4. `add deck.pptx / --type slide ...`

### Quick-start CRUD/readback flow
Source: `OfficeCLI/README.md`
- `create`
- `add`
- `view outline`
- `view html`
- `get --json`

### Resident + batch workflow
Source: `OfficeCLI/README.md`, `CommandBuilder.Batch.cs`, `CommandBuilder.cs`
- `open` / repeated mutation / `close`
- `batch` via stdin JSON
- `batch --commands '[...]'`

## 2) CI smoke flow to convert into officekit e2e tests

Source: `OfficeCLI/.github/workflows/build.yml`

Required smoke scenario:
1. create `test_smoke.docx`
2. add paragraph under `/body`
3. get `/body/p[1]`
4. assert non-error output
5. clean up artifact

This should become the minimum `officekit` Bun CI e2e gate even before full parity coverage exists.

## 3) Curated example scripts to harvest as first-wave fixtures

### Word
- `OfficeCLI/examples/word/gen-formulas.sh`
  - parity families: formula insertion, formatted content, Word math support
- `OfficeCLI/examples/word/gen-complex-tables.sh`
  - parity families: tables, borders, shading, alignment, row/column structure
- `OfficeCLI/examples/word/gen-complex-textbox.sh`
  - parity families: positioned text boxes, font styling, paragraph formatting

### Excel
- `OfficeCLI/examples/excel/gen-beautiful-charts.sh`
  - parity families: chart creation/styling, legend/axis formatting, positioning
- `OfficeCLI/examples/excel/gen-charts-demo.sh`
  - parity families: multi-chart-type coverage, 3D/stacked/clustered variants

### PowerPoint
- `OfficeCLI/examples/ppt/gen-beautiful-pptx.sh`
  - parity families: slide creation, shape positioning, layout patterns, morph transitions
- `OfficeCLI/examples/ppt/gen-animations-pptx.sh`
  - parity families: entrance/emphasis/exit animations, timing/sequencing
- `OfficeCLI/examples/ppt/gen-video-pptx.py`
  - parity families: embedded media/video positioning, cross-language invocation expectations

## 4) Pre-generated binary artifacts to keep as differential references

### Root examples
- `OfficeCLI/examples/Alien_Guide.pptx`
- `OfficeCLI/examples/budget_review_v2.pptx`
- `OfficeCLI/examples/Cat-Secret-Life.pptx`
- `OfficeCLI/examples/product_launch_morph.pptx`

### Generated-output families referenced by example READMEs
- Word outputs under `OfficeCLI/examples/word/outputs/`
- Excel outputs under `OfficeCLI/examples/excel/outputs/`
- PowerPoint outputs under `OfficeCLI/examples/ppt/outputs/`

Use these as comparison seeds for:
- structured `get` / `query` / `view outline` output
- preview HTML/SVG snapshots where deterministic
- document package structure diffs where stable

## 5) Documentation/fixture metadata worth preserving

Source files:
- `OfficeCLI/examples/README.md`
- `OfficeCLI/examples/word/README.md`
- `OfficeCLI/examples/excel/README.md`
- `OfficeCLI/examples/ppt/README.md`

Harvest these docs because they encode:
- expected command sequences
- canonical DOM/path examples (`/body/p[1]`, `/Sheet1/A1`, `/slide[1]/shape[1]`)
- property names that can populate the parity matrix
- user-facing descriptions of supported chart types / preview flows / layout concepts

## 6) Recommended target ingest layout (for later implementation lanes)

```text
packages/parity-tests/
  fixtures/
    officecli/
      readme-scenarios/
      word/
      excel/
      ppt/
      generated-binaries/
  tests/
    smoke/
    differential/
    preview/
```

## 7) Current lane-1 deliverables produced from this harvest
- `docs/migration/capability-matrix.md`
- `docs/migration/source-to-target-ledger.md`
- `docs/migration/fixture-harvest.md`
- `docs/migration/officecli-source-inventory.json`

## 8) Known harvest gaps
- Example script bodies have not yet been copied into `officekit`; this lane captures the manifest and target placement first.
- README showcase GIF/assets are useful for docs lineage but not yet selected as deterministic test fixtures.
- Wiki-derived help content is referenced by CI but not yet mirrored; lane 4 should decide whether to vendor, generate, or replace it.
