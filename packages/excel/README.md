# @officekit/excel

Lane-3 parity-first scaffold for the Excel adapter.

## Current contents

- `src/manifest.js` — source lineage, capability families, deferred shared-core contracts, implementation milestones, and parity risks
- `src/index.js` — package exports for manifest/contract helpers
- `test/manifest.test.js` — Bun-backed invariants for workbook/chart/import parity expectations

## Source lineage

Primary source families:

- `OfficeCLI/src/officecli/Handlers/ExcelHandler.cs`
- `OfficeCLI/src/officecli/Handlers/Excel/ExcelHandler.Add.cs`
- `OfficeCLI/src/officecli/Handlers/Excel/ExcelHandler.Import.cs`
- `OfficeCLI/src/officecli/Handlers/Excel/ExcelHandler.Query.cs`
- `OfficeCLI/src/officecli/Handlers/Excel/ExcelHandler.Set*.cs`
- `OfficeCLI/src/officecli/Handlers/Excel/ExcelHandler.View.cs`
- `OfficeCLI/src/officecli/Core/ExcelStyleManager.cs`
- `OfficeCLI/src/officecli/Core/FormulaEvaluator*.cs`
- `OfficeCLI/src/officecli/Core/PivotTableHelper.cs`

## Parity-critical families

- workbook/sheet/range semantics
- formulas and style handling
- charts and pivots
- CSV/TSV import
- filtered raw sheet views
- HTML preview

## Deferred shared-core decisions

- selector/path grammar
- shared result envelope
- shared color/unit helpers
- preview server orchestration
