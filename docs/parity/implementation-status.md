# Current implementation status

This report is the lane-4 parity snapshot for the current `officekit` checkpoint. It is intentionally conservative: it records what is **explicitly evidenced today** by package manifests, fixture inventory, and runnable tests, and separates that from work that is still pending.

## Status legend

- `scaffolded` — source lineage, package contracts, and parity expectations are encoded in code/tests/docs.
- `implemented` — executable document behavior exists for that slice.
- `verified` — runnable checks prove the implemented behavior against fixtures or differential expectations.

## Word / Excel / PowerPoint slice status

| Format | Current status | Fixture-backed evidence | Explicitly supported today | Remaining gaps |
| --- | --- | --- | --- | --- |
| Word | implemented | `word-formulas-script`, `word-tables-script`, `word-textbox-script`, `word-complex-formulas-output`, `word-complex-tables-output` | Create/add/set/get/query/remove/move/swap/copy/raw/raw-set/add-part/merge/check/view flows are runnable; metadata-free OOXML fallback, mixed paragraph/table order, table cell mutation, `html/forms/json` views, and fixture-backed Word package tests are passing | Section/style/document-settings fidelity, broader raw-part parity, long-tail form/rendering fidelity, and more fixture-backed differential checks still need work |
| Excel | implemented | `excel-beautiful-charts-script`, `excel-charts-demo-script`, `excel-sales-report-output`, `excel-charts-demo-output`, `excel-beautiful-charts-output` | Workbook/sheet/cell/range CRUD, import, raw/raw-set access, named ranges, validations, comments, tables, sparklines, charts, pivots, shapes/pictures, add-part chart creation, merge for OOXML text templates, metadata-free OOXML mutation, and broad formula/view coverage are all runnable with green CLI/package tests | Deep formula parity, chart-property depth, style-manager depth, deeper pivot semantics, and more mixed real-workbook pressure testing remain the main parity gaps |
| PowerPoint | implemented | `ppt-beautiful-script`, `ppt-animations-script`, `ppt-video-script`, `ppt-3d-script`, `ppt-beautiful-output`, `ppt-data-output`, `ppt-animations-output`, referenced `ppt-3d-model-asset` | Slide/shape/table/media/chart/theme/notes/placeholder/background/hyperlink/connector/group/animation/3D flows are runnable; `html/svg` preview, overflow checks, metadata-free OOXML fallback, add-part chart creation, merge, watch/unwatch, and resident open/close flows are passing | Layout/theme inheritance fidelity, preview rendering fidelity vs OfficeCLI, long-tail media/3D semantics, and more fixture-replay/differential checks still need work |

## What is verified by lane-4 tests right now

1. The harvested fixture manifest still contains canonical Word, Excel, and PowerPoint scenarios.
2. Each format package exposes parity-critical manifest/contract metadata that matches the fixture corpus.
3. End-to-end CLI tests prove live create/get/set/add/remove/view/raw/import/watch behavior across all three formats.
4. Metadata-free OOXML fallback paths are exercised for Word, Excel, and PowerPoint.
5. Preview/watch flows are executable today, including preview server health, live refresh behavior, and explicit `unwatch` teardown.
6. `open`/`close` resident sessions are executable today, and resident-backed mutation reads/writes are covered by CLI tests.
7. `merge` and `add-part` have live CLI coverage across more than one format, not just package-local unit tests.
8. Documentation keeps the OfficeCLI lineage explicit and separates implemented behavior from remaining parity gaps.

## Remaining gaps

- The package manifests and some migration docs still understate how much executable behavior now exists; they should be read as historical planning artifacts unless they are backed by current tests.
- OfficeCLI command-surface parity is materially closer now, but still not complete: `update` is still missing, resident mode is functional but not yet deeply optimized across the whole command surface, and not every OfficeCLI long-tail flag/parameter variant has been replayed.
- Excel remains the deepest functional parity gap area, especially formula coverage, chart property depth, style semantics, pivot behavior, and complex mixed-workbook compatibility.
- Word and PowerPoint are runnable, but still have fidelity gaps in long-tail rendering, theme/layout/style behavior, and raw OOXML escape-hatch breadth.
- Differential document-output comparisons against OfficeCLI fixtures still need to expand before claiming near-complete parity.
