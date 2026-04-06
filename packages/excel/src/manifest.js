const sourceFiles = Object.freeze([
  "OfficeCLI/src/officecli/Handlers/ExcelHandler.cs",
  "OfficeCLI/src/officecli/Handlers/Excel/ExcelHandler.Add.cs",
  "OfficeCLI/src/officecli/Handlers/Excel/ExcelHandler.Import.cs",
  "OfficeCLI/src/officecli/Handlers/Excel/ExcelHandler.Query.cs",
  "OfficeCLI/src/officecli/Handlers/Excel/ExcelHandler.Set.cs",
  "OfficeCLI/src/officecli/Handlers/Excel/ExcelHandler.View.cs",
  "OfficeCLI/src/officecli/Core/ExcelStyleManager.cs",
  "OfficeCLI/src/officecli/Core/FormulaEvaluator.cs",
  "OfficeCLI/src/officecli/Core/PivotTableHelper.cs"
]);

const publicSurface = Object.freeze([
  "Add",
  "AddPart",
  "Get",
  "Query",
  "Set",
  "Remove",
  "Move",
  "Swap",
  "CopyFrom",
  "Import",
  "Raw",
  "RawSet",
  "ViewAsText",
  "ViewAsAnnotated",
  "ViewAsOutline",
  "ViewAsStats",
  "ViewAsIssues",
  "ViewAsHtml"
]);

const capabilityFamilies = Object.freeze({
  workbook: Object.freeze([
    "workbook-metadata",
    "sheet-management",
    "cell-range-navigation",
    "shared-strings",
    "filtered-raw-sheet-views"
  ]),
  calculations: Object.freeze([
    "formulas",
    "formula-parser",
    "styles",
    "charts",
    "pivots"
  ]),
  flows: Object.freeze([
    "import",
    "html-preview",
    "query",
    "mutation-operations",
    "raw-xml"
  ])
});

const deferredCoreContracts = Object.freeze([
  "selector-and-path-grammar",
  "result-envelope-shape",
  "color-and-unit-helpers",
  "preview-server-orchestration"
]);

const implementationMilestones = Object.freeze([
  "Model workbook, sheet, row, cell, and chart addressability with source-compatible semantics",
  "Port add/set/remove/move/copy/import flows before broad abstraction sharing",
  "Port formulas, styles, charts, and pivot helpers with fixture-backed scenarios",
  "Recreate HTML preview and filtered raw-sheet views for parity checks"
]);

const parityRisks = Object.freeze([
  "Formula evaluation and parser helpers may hide semantic edge cases",
  "Charts and pivots rely on helper families outside the main handler partials",
  "Filtered raw output must preserve sheet, row, and column slicing behavior"
]);

export const excelAdapterManifest = Object.freeze({
  packageName: "@officekit/excel",
  exclusions: Object.freeze(["mcp"]),
  sourceFiles,
  publicSurface,
  capabilityFamilies,
  deferredCoreContracts,
  implementationMilestones,
  parityRisks
});

export function getExcelAdapterManifest() {
  return structuredClone(excelAdapterManifest);
}

export function summarizeExcelAdapter() {
  return {
    packageName: excelAdapterManifest.packageName,
    surfaceCount: excelAdapterManifest.publicSurface.length,
    riskCount: excelAdapterManifest.parityRisks.length,
    previewModes: ["html"]
  };
}

const excelAdapterContract = Object.freeze({
  mutationFamilies: Object.freeze([
    "add",
    "set",
    "remove",
    "move",
    "swap",
    "copy-from",
    "import",
    "raw-set"
  ]),
  queryFamilies: Object.freeze([
    "get",
    "query",
    "view-text",
    "view-annotated",
    "view-outline",
    "view-stats",
    "view-issues",
    "view-html"
  ]),
  canonicalPaths: Object.freeze([
    "/workbook",
    "/styles",
    "/sharedstrings",
    "/Sheet1",
    "/Sheet1/drawing",
    "/Sheet1/chart[1]",
    "/chart[1]"
  ]),
  previewModes: Object.freeze(["html"]),
  laneNotes: Object.freeze([
    "Keep sheet addressability and filtered raw-sheet views package-local until shared semantics are proven",
    "Treat formulas, charts, pivots, and style management as parity-critical even if early vertical slices are thinner"
  ])
});

export function getExcelAdapterContract() {
  return structuredClone(excelAdapterContract);
}

export function summarizeExcelAdapterContract() {
  return {
    canonicalPathCount: excelAdapterContract.canonicalPaths.length,
    previewModes: [...excelAdapterContract.previewModes]
  };
}
