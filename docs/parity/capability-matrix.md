# officekit capability matrix (lane 1 inventory)

Status legend: `inventoried` = source capability mapped into officekit target ownership; `excluded` = intentionally out of scope for officekit v1.

## Product / CLI surfaces

| Source family | Source evidence | officekit target package/module | Verification layer | Status | Notes |
| --- | --- | --- | --- | --- | --- |
| Root dispatch + UTF-8 bootstrap | `OfficeCLI/src/officecli/Program.cs` | `packages/cli/src/program.ts` | unit + e2e | inventoried | Preserve agent-friendly defaults; rename product to `officekit`. |
| Format-prefixed command rewrite (`docx/xlsx/pptx`) | `Program.cs`, `HelpCommands.cs` | `packages/cli/src/aliases.ts` | unit + e2e | inventoried | Compatibility can change, but helpful aliasing should be considered. |
| `open` / `close` resident mode | `CommandBuilder.cs`, `Core/ResidentClient.cs`, `Core/ResidentServer.cs` | `packages/cli`, `packages/core/src/resident/*` | integration + e2e | inventoried | Needed for low-latency multi-step edits. |
| `watch` / `unwatch` live preview | `CommandBuilder.Watch.cs`, `Core/WatchServer.cs`, `Core/WatchNotifier.cs`, `Core/HtmlPreviewHelper.cs` | `packages/preview`, `packages/cli` | integration + preview e2e | inventoried | Developer-facing visualization workflow is a v1 requirement. |
| `view` semantic modes (`text`, `annotated`, `outline`, `stats`, `issues`, `html`) | `CommandBuilder.View.cs`, handlers, `OutputFormatter.cs` | `packages/cli`, per-format packages, `packages/preview` | unit + integration + parity | inventoried | HTML mode feeds preview parity. |
| `get` / `query` structured inspection | `CommandBuilder.GetQuery.cs`, handlers, `Core/DocumentNode.cs` | `packages/core`, per-format packages, `packages/cli` | unit + integration + parity | inventoried | Shared selector/path semantics must stay centralized. |
| `set` mutation flow | `CommandBuilder.Set.cs`, handlers | `packages/core`, per-format packages, `packages/cli` | integration + parity | inventoried | Keep unsupported-prop reporting behavior equivalent. |
| `add` / `remove` / `move` / `swap` | `CommandBuilder.Add.cs`, handlers | `packages/core`, per-format packages, `packages/cli` | integration + parity | inventoried | Includes insert-position semantics (`before` / `after` / index). |
| `raw` / `raw-set` / `add-part` | `CommandBuilder.Raw.cs`, handlers, `Core/RawXmlHelper.cs` | `packages/core`, per-format packages, `packages/cli` | unit + integration + parity | inventoried | Universal OpenXML escape hatch is parity-critical. |
| `validate` / `check` | `CommandBuilder.Check.cs`, handlers, `Core/DocumentIssue.cs` | `packages/core`, per-format packages, `packages/cli` | integration + parity | inventoried | `check` covers layout/content diagnostics. |
| `batch` JSON execution | `CommandBuilder.Batch.cs`, `Core/BatchTypes.cs` | `packages/core`, `packages/cli` | unit + e2e | inventoried | Must support stdin/inline JSON command batches. |
| `import` CSV/TSV into Excel | `CommandBuilder.Import.cs` | `packages/excel`, `packages/cli` | integration + parity | inventoried | Excel-only import surface. |
| `create` blank Office docs | `CommandBuilder.Import.cs`, `BlankDocCreator.cs` | `packages/cli`, per-format packages | integration + e2e | inventoried | Needed for CI smoke and quick start flows. |
| `merge` template merge | `CommandBuilder.Import.cs`, `Core/TemplateMerger.cs` | `packages/core`, `packages/word`, `packages/excel`, `packages/ppt`, `packages/cli` | integration + parity | inventoried | Must support placeholder replacement across all three formats. |
| `install` binary + agent setup | `Program.cs`, `Core/Installer.cs` | `packages/install`, `packages/cli` | e2e + docs acceptance | inventoried | MCP fallback is excluded; skill install remains required. |
| `skills list/install` | `Program.cs`, `Core/SkillInstaller.cs`, embedded resources | `packages/skills`, `packages/install`, `packages/cli` | e2e + docs acceptance | inventoried | Agent onboarding surface is part of v1. |
| `config` / auto-update settings / logging | `Program.cs`, `Core/UpdateChecker.cs`, `Core/CliLogger.cs` | `packages/install`, `packages/cli` | unit + e2e | inventoried | Config includes `autoUpdate`, log toggles, and clear-log flow. |
| Nested help + wiki-backed docs | `HelpCommands.cs`, `WikiHelpLoader.cs` | `packages/docs`, `packages/cli` | docs acceptance + e2e | inventoried | Preserve layered help intent in Node-native docs surface. |
| Background update check | `Core/UpdateChecker.cs` | `packages/install` | unit + e2e | inventoried | Product-level behavior, not just packaging. |
| MCP server / MCP registration | `Program.cs`, `Core/McpServer.cs`, `Core/McpInstaller.cs` | _none_ | policy check | excluded | Explicit v1 exclusion per spec/PRD. |

## Shared core families

| Shared family | Source evidence | officekit target package/module | Verification layer | Status | Notes |
| --- | --- | --- | --- | --- | --- |
| Document abstraction + node model | `Core/IDocumentHandler.cs`, `Core/DocumentNode.cs`, `Core/DocumentHandlerFactory.cs` | `packages/core/src/document/*` | unit + integration | inventoried | Defines the cross-format contract boundary. |
| Path aliases + parse helpers | `Core/PathAliases.cs`, `Core/ParseHelpers.cs`, `Core/AttributeFilter.cs` | `packages/core/src/path/*` | unit | inventoried | Critical for path/query parity. |
| Output envelopes + warnings/errors | `Core/OutputFormatter.cs`, `Core/CliException.cs` | `packages/core/src/output/*` | unit + e2e | inventoried | Agent-facing structured output must remain stable. |
| Units / geometry / color helpers | `Core/Units.cs`, `Core/EmuConverter.cs`, `Core/SpacingConverter.cs`, `Core/ColorMath.cs`, `Core/ThemeColorResolver.cs` | `packages/core/src/formatting/*` | unit | inventoried | Shared across Word/Excel/PPT positioning and styling. |
| Image + media ingestion | `Core/ImageSource.cs`, `Core/DrawingEffectsHelper.cs` | `packages/core/src/media/*` | unit + integration | inventoried | Shared media pipeline for documents and preview. |
| Generic XML query / raw fallback | `Core/GenericXmlQuery.cs`, `Core/RawXmlHelper.cs` | `packages/core/src/raw/*` | unit + integration | inventoried | Needed for universal escape hatch and selectors. |
| Template merge engine | `Core/TemplateMerger.cs` | `packages/core/src/template/*` | integration + parity | inventoried | Cross-format placeholder replacement. |
| Resident protocol | `Core/ResidentClient.cs`, `Core/ResidentServer.cs` | `packages/core/src/resident/*` | integration + e2e | inventoried | Shared infra for `open`/`close` acceleration. |
| Preview/watch transport | `Core/HtmlPreviewHelper.cs`, `Core/WatchServer.cs`, `Core/WatchNotifier.cs` | `packages/preview/src/*` | integration + preview e2e | inventoried | Must support HTML patching + SSE/live updates. |
| Installer / updater / logging | `Core/Installer.cs`, `Core/UpdateChecker.cs`, `Core/CliLogger.cs` | `packages/install/src/*` | unit + e2e | inventoried | Product bootstrap and maintenance flow. |
| Skill packaging/install | `Core/SkillInstaller.cs` | `packages/skills/src/*` | e2e + docs acceptance | inventoried | Includes multi-agent target directories + embedded skill resources. |
| Chart helpers | `Core/ChartBuilder.cs`, `Core/ChartExBuilder.cs`, `Core/ChartAdvancedFeatures.cs`, `Core/ChartHelper.cs`, `Core/ChartReader.cs`, `Core/ChartSetter.cs`, `Core/ChartSetterHelpers.cs`, `Core/ChartPresets.cs`, `Core/ChartSvgRenderer.cs` | `packages/excel`, `packages/ppt`, `packages/preview` | unit + integration + parity | inventoried | Shared chart semantics span Excel, PowerPoint, and HTML preview. |
| Formula engine | `Core/FormulaParser.cs`, `Core/FormulaEvaluator.cs`, `Core/FormulaEvaluator.Helpers.cs`, `Core/FormulaEvaluator.Functions.cs` | `packages/excel` (promote shared pieces only if reused) | unit + integration + parity | inventoried | 150+ built-in Excel functions are explicit parity scope. |
| Theme / extended properties | `Core/ThemeHandler.cs`, `Core/ExtendedPropertiesHandler.cs` | `packages/ppt`, `packages/word`, `packages/core` | integration + parity | inventoried | Theme metadata needs both read and write coverage. |

## Word capability families

| Capability family | Source evidence | officekit target package/module | Verification layer | Status | Notes |
| --- | --- | --- | --- | --- | --- |
| Word package open/raw parts | `Handlers/WordHandler.cs` | `packages/word/src/raw/*` | integration | inventoried | `/document`, `/styles`, `/settings`, `/numbering`, `/header[n]`, `/footer[n]`, `/chart[n]`, `/comments`. |
| Semantic views + issues | `IDocumentHandler`, `HelpCommands.cs` | `packages/word/src/view/*` | integration + parity | inventoried | Text/annotated/outline/stats/issues modes. |
| Paragraphs, runs, styles | `README.md`, `examples/word/*`, `HelpCommands.cs` | `packages/word/src/nodes/text/*` | integration + parity | inventoried | Core authoring path. |
| Tables / cells / formatting | `README.md`, `examples/word/gen-complex-tables.sh` | `packages/word/src/nodes/table/*` | integration + parity | inventoried | Includes sizing, borders, shading, alignment. |
| Text boxes / positioned content | `examples/word/gen-complex-textbox.sh` | `packages/word/src/nodes/textbox/*` | integration + parity | inventoried | Needed for richer layout parity. |
| Equations / formulas | `README.md`, `examples/word/gen-formulas.sh` | `packages/word/src/nodes/equation/*` | integration + parity | inventoried | LaTeX/equation insertion is explicitly showcased. |
| Headers / footers / bookmarks / comments / charts | `HelpCommands.cs`, `WordHandler.cs`, README feature list | `packages/word/src/nodes/document-structure/*` | integration + parity | inventoried | Long-tail doc structure surface. |
| Footnotes / endnotes / TOC / sections / styles | `HelpCommands.cs` | `packages/word/src/nodes/document-structure/*` | integration + parity | inventoried | High-detail parity requirement means these cannot be ignored. |
| Raw/query fallback | `WordHandler.cs`, `GenericXmlQuery.cs` | `packages/word/src/raw/*` | integration + parity | inventoried | Escape hatch for unsupported high-detail operations. |

## Excel capability families

| Capability family | Source evidence | officekit target package/module | Verification layer | Status | Notes |
| --- | --- | --- | --- | --- | --- |
| Workbook / worksheet / shared strings / styles raw access | `Handlers/ExcelHandler.cs`, `ExcelStyleManager.cs` | `packages/excel/src/raw/*` | integration | inventoried | Raw parts include workbook, styles, shared strings, sheets, drawings, charts. |
| Cell / range / sheet CRUD | `README.md`, `HelpCommands.cs` | `packages/excel/src/nodes/grid/*` | integration + parity | inventoried | Includes `$Sheet:A1` style addressing. |
| Formula parsing + evaluation | `FormulaParser.cs`, `FormulaEvaluator*.cs` | `packages/excel/src/formula/*` | unit + integration + parity | inventoried | Explicit 150+ function coverage signal. |
| Chart creation + editing | `Chart*.cs`, `README.md`, `examples/excel/*` | `packages/excel/src/chart/*` | integration + parity | inventoried | Includes standard + extended chart types, combo/waterfall/reference lines. |
| Pivot tables | `PivotTableHelper.cs`, README feature list | `packages/excel/src/pivot/*` | integration + parity | inventoried | Must be represented in ledger even if late in implementation sequence. |
| Styling / number formats / conditional formatting | `ExcelStyleManager.cs`, README feature list | `packages/excel/src/style/*` | integration + parity | inventoried | Core spreadsheet usability surface. |
| Named ranges / validation / autofilter / comments / shapes / sparklines | README feature list | `packages/excel/src/nodes/advanced/*` | integration + parity | inventoried | Long-tail parity surface. |
| CSV/TSV import | `CommandBuilder.Import.cs`, README use cases | `packages/excel/src/import/*` | integration + e2e | inventoried | Direct command surface with migration value. |
| Raw/query fallback | `Handlers/ExcelHandler.cs`, `GenericXmlQuery.cs` | `packages/excel/src/raw/*` | integration + parity | inventoried | Needed for stable escape hatch. |

## PowerPoint capability families

| Capability family | Source evidence | officekit target package/module | Verification layer | Status | Notes |
| --- | --- | --- | --- | --- | --- |
| Presentation / slide / slide master / slide layout / notes raw access | `Handlers/PowerPointHandler.cs` | `packages/ppt/src/raw/*` | integration | inventoried | Raw parts include `/presentation`, `/slide[N]`, `/slideMaster[N]`, `/slideLayout[N]`, `/noteSlide[N]`. |
| Slide lifecycle + layout resolution | `PowerPointHandler.cs` | `packages/ppt/src/slide/*` | integration + parity | inventoried | Layout matching by name/type/index is explicit logic to preserve. |
| Shapes / text / positioning / presets | `README.md`, `PowerPointHandler.cs` | `packages/ppt/src/nodes/shape/*` | integration + parity | inventoried | Core authoring path for deck construction. |
| Tables / pictures / media | `README.md`, `PowerPointHandler.cs` | `packages/ppt/src/nodes/media/*` | integration + parity | inventoried | Includes cleanup of media relationships on delete. |
| Charts | `PowerPointHandler.cs`, chart helpers, README feature list | `packages/ppt/src/chart/*` | integration + parity | inventoried | Add-part chart path + slide chart editing. |
| Placeholders / slide masters / themes / templates | `PowerPointHandler.cs`, `ThemeHandler.cs`, `TemplateMerger.cs`, README feature list | `packages/ppt/src/theme/*` | integration + parity | inventoried | Required for professional template migration. |
| Animations / morph / zoom / connectors / groups / notes | README feature list, examples/ppt/* | `packages/ppt/src/nodes/presentation-effects/*` | integration + parity | inventoried | Preview must render parity-critical effects where feasible. |
| Video / audio / 3D model workflows | `PowerPointHandler.cs`, `examples/ppt/gen-video-pptx.py`, `examples/ppt/gen-3d-sun-pptx.sh`, `examples/ppt/models/sun.glb` | `packages/ppt/src/nodes/media/*`, `packages/preview/src/render/*` | integration + parity | inventoried | Keep large reference assets in fixture manifest even if not vendored immediately. |
| HTML preview / watch patching | `HtmlPreviewHelper.cs`, `WatchServer.cs`, `WatchNotifier.cs` | `packages/preview/src/*` | preview e2e + differential snapshots | inventoried | A developer-facing release gate, not a stretch goal. |
