# officekit source-to-target ledger (lane 1)

This ledger maps the current OfficeCLI source surface into the recommended officekit Node/Bun workspace from the approved PRD. All items below are currently **inventoried** and intentionally scoped so later implementation lanes can mark them `scaffolded`, `implemented`, and `verified` without re-doing discovery work.

## Ledger

| Source cluster | Source files | Target package / candidate module | Lane owner | Verification layer | Status | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| Root bootstrap + special command dispatch | `src/officecli/Program.cs` | `packages/cli/src/program.ts` | lane 2 (core + CLI) | unit + e2e | inventoried | Rename to `officekit`; keep UTF-8 output and non-blocking update check behavior. |
| Root command registration | `src/officecli/CommandBuilder.cs` | `packages/cli/src/commands/root.ts` | lane 2 | unit + e2e | inventoried | Registers open/close/resident, watch, CRUD, raw, validation, batch, import, create, merge. |
| Semantic view surface | `src/officecli/CommandBuilder.View.cs` | `packages/cli/src/commands/view.ts` | lane 2 + per-format lanes | integration + parity | inventoried | Includes HTML preview mode handoff into preview package. |
| Query surface | `src/officecli/CommandBuilder.GetQuery.cs` | `packages/cli/src/commands/query.ts` | lane 2 + per-format lanes | integration + parity | inventoried | Shared selector / path behavior belongs in `packages/core`. |
| Mutation surface | `src/officecli/CommandBuilder.Set.cs`, `CommandBuilder.Add.cs` | `packages/cli/src/commands/mutate.ts` | lane 2 + per-format lanes | integration + parity | inventoried | Includes add/remove/move/swap and unsupported-prop reporting. |
| Raw OpenXML surface | `src/officecli/CommandBuilder.Raw.cs`, `Core/RawXmlHelper.cs`, `Core/GenericXmlQuery.cs` | `packages/core/src/raw/*`, per-format raw modules | lane 2 + per-format lanes | unit + integration + parity | inventoried | Universal fallback for unsupported high-detail mutations. |
| Validation + issue scan | `src/officecli/CommandBuilder.Check.cs`, `Core/DocumentIssue.cs` | `packages/core/src/validation/*`, per-format validators | lane 2 + per-format lanes | integration + parity | inventoried | Includes OpenXML validation and layout/content issue reporting. |
| Batch execution | `src/officecli/CommandBuilder.Batch.cs`, `Core/BatchTypes.cs` | `packages/core/src/batch/*`, `packages/cli/src/commands/batch.ts` | lane 2 | unit + e2e | inventoried | Must support stdin and inline commands JSON. |
| Create / import / merge flows | `src/officecli/CommandBuilder.Import.cs`, `BlankDocCreator.cs`, `Core/TemplateMerger.cs` | `packages/cli/src/commands/create-import-merge.ts`, `packages/core/src/template/*` | lane 2 + lanes 3/4 | integration + parity | inventoried | `create` is a thin vertical slice gate; `merge` is cross-format. |
| Resident protocol | `Core/ResidentClient.cs`, `Core/ResidentServer.cs` | `packages/core/src/resident/*` | lane 2 | integration + e2e | inventoried | Required for low-latency agent workflows. |
| Help / wiki docs | `HelpCommands.cs`, `WikiHelpLoader.cs` | `packages/docs/src/help/*`, `packages/cli/src/commands/help.ts` | lane 4 (docs/install/preview) | docs acceptance + e2e | inventoried | Preserve layered command/element/property help model. |
| Installer + update + config + logging | `Core/Installer.cs`, `Core/UpdateChecker.cs`, `Core/CliLogger.cs` | `packages/install/src/*`, `packages/cli/src/commands/install.ts` | lane 4 | unit + e2e | inventoried | MCP branches are excluded, but skill/install/config/update remain in scope. |
| Skills packaging | `Core/SkillInstaller.cs`, embedded skill resources | `packages/skills/src/*`, `packages/install/src/skills/*` | lane 4 | e2e + docs acceptance | inventoried | Keep multi-agent install targets and resource rewrite rules visible. |
| Document abstraction | `Core/IDocumentHandler.cs`, `Core/DocumentNode.cs`, `Core/DocumentHandlerFactory.cs` | `packages/core/src/document/*` | lane 2 | unit + integration | inventoried | Foundational cross-format contract. |
| Shared parse/format helpers | `Core/PathAliases.cs`, `ParseHelpers.cs`, `AttributeFilter.cs`, `OutputFormatter.cs`, `CliException.cs` | `packages/core/src/path/*`, `packages/core/src/output/*` | lane 2 | unit + integration | inventoried | Stabilize shared semantics before handler ports. |
| Shared layout / color / unit helpers | `Core/Units.cs`, `EmuConverter.cs`, `SpacingConverter.cs`, `ColorMath.cs`, `ThemeColorResolver.cs`, `DrawingEffectsHelper.cs` | `packages/core/src/formatting/*` | lane 2 | unit | inventoried | Promote only genuinely shared helpers. |
| Shared media helpers | `Core/ImageSource.cs`, `FontMetricsReader.cs` | `packages/core/src/media/*` | lane 2 + lane 4 | unit + integration | inventoried | Supports both document output and preview rendering. |
| Word handler root | `Handlers/WordHandler.cs` | `packages/word/src/index.ts` | lane 3 (Word/Excel/PPT) | integration + parity | inventoried | Raw parts include document/styles/settings/numbering/comments/header/footer/chart. |
| Word semantic/query/mutation families | `Handlers/WordHandler.cs` + help/readme/examples | `packages/word/src/nodes/{text,table,textbox,equation,document-structure}/*` | lane 3 | integration + parity | inventoried | Paragraphs, runs, tables, text boxes, equations, headers/footers, bookmarks, TOC, sections. |
| Excel handler root | `Handlers/ExcelHandler.cs` | `packages/excel/src/index.ts` | lane 3 | integration + parity | inventoried | Raw parts include workbook/styles/sharedstrings/sheets/drawings/charts. |
| Excel formulas | `Core/FormulaParser.cs`, `Core/FormulaEvaluator*.cs` | `packages/excel/src/formula/*` | lane 3 | unit + integration + parity | inventoried | Dedicated lane needed because formula parity is deep and stateful. |
| Excel chart/pivot/style families | `Core/Chart*.cs`, `Core/PivotTableHelper.cs`, `Core/ExcelStyleManager.cs`, README/examples | `packages/excel/src/{chart,pivot,style,nodes}/*` | lane 3 | integration + parity | inventoried | Includes extended chart types, combo/waterfall helpers, pivot tables, named ranges, validation, autofilter. |
| PowerPoint handler root | `Handlers/PowerPointHandler.cs` | `packages/ppt/src/index.ts` | lane 3 | integration + parity | inventoried | Raw parts include slides, masters, layouts, notes; logic includes layout/placeholder resolution. |
| PowerPoint presentation features | `Handlers/PowerPointHandler.cs`, `Core/ThemeHandler.cs`, `Core/TemplateMerger.cs`, README/examples | `packages/ppt/src/{slide,shape,chart,theme,nodes}/*` | lane 3 | integration + parity | inventoried | Slides, shapes, charts, placeholders, themes, templates, notes, media, animations, morph, zoom, 3D/video workflows. |
| Preview/watch HTML server | `Core/HtmlPreviewHelper.cs`, `Core/WatchServer.cs`, `Core/WatchNotifier.cs`, `Resources/preview.css`, `Resources/preview.js` | `packages/preview/src/{render,server,assets}/*` | lane 4 | preview e2e + snapshots | inventoried | Critical developer-facing workflow. |
| Source docs + examples corpus | `README.md`, `examples/**`, `.github/workflows/build.yml` | `packages/parity-tests/fixtures/source-officecli/*`, `docs/parity/*` | lane 1 (this lane) | inventory + parity + docs acceptance | inventoried | Harvested in this commit as the parity baseline. |
| MCP server and registration | `Core/McpServer.cs`, `Core/McpInstaller.cs`, `Program.cs` | _none_ | n/a | excluded | Explicit non-goal for officekit v1. |

## Immediate follow-on guidance for implementation lanes

1. **Lane 2** should scaffold the workspace contracts around the inventoried command families before any deep handler port.
2. **Lane 3** should treat Word, Excel, and PowerPoint as separate modules but avoid promoting helpers into `packages/core` until a second consumer proves reuse.
3. **Lane 4** should consume the harvested fixture corpus here for preview/docs/install/skills acceptance tests and ensure README lineage language stays explicit about OfficeCLI migration.
