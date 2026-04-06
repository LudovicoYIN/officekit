const sourceFiles = Object.freeze([
  "OfficeCLI/src/officecli/Handlers/PowerPointHandler.cs",
  "OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.Add.cs",
  "OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.Chart.cs",
  "OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.HtmlPreview.cs",
  "OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.Mutations.cs",
  "OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.Query.cs",
  "OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.Set.cs",
  "OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.ShapeProperties.cs",
  "OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.SvgPreview.cs",
  "OfficeCLI/src/officecli/Handlers/Pptx/PowerPointHandler.Theme.cs"
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
  "Raw",
  "RawSet",
  "Batch",
  "ViewAsText",
  "ViewAsAnnotated",
  "ViewAsOutline",
  "ViewAsStats",
  "ViewAsIssues",
  "ViewAsHtml",
  "ViewAsSvg",
  "CheckShapeTextOverflow"
]);

const capabilityFamilies = Object.freeze({
  slides: Object.freeze([
    "slide-management",
    "layout-inheritance",
    "placeholder-resolution",
    "notes"
  ]),
  shapesAndMedia: Object.freeze([
    "shapes",
    "tables",
    "text",
    "media",
    "hyperlinks",
    "background-fill-effects"
  ]),
  renderingAndValidation: Object.freeze([
    "charts",
    "theme",
    "html-preview",
    "svg-preview",
    "overflow-checks",
    "animations"
  ]),
  mutations: Object.freeze([
    "add",
    "set",
    "remove",
    "move",
    "swap",
    "copy-from",
    "raw-xml",
    "batch"
  ])
});

const deferredCoreContracts = Object.freeze([
  "selector-and-path-grammar",
  "result-envelope-shape",
  "shared-preview-server",
  "cross-format-query-helpers"
]);

const implementationMilestones = Object.freeze([
  "Model slide, shape, table, chart, and placeholder resolution with parity-safe path semantics",
  "Port add/set/remove/move/copy flows for slide and shape families before shared abstraction hardening",
  "Port theme, chart, notes, and overflow-check helpers with canonical fixtures",
  "Recreate HTML and SVG preview fidelity before wiring a shared watch surface"
]);

const parityRisks = Object.freeze([
  "Layout and theme inheritance can diverge subtly from OfficeCLI behavior",
  "Preview fidelity depends on HTML and SVG rendering staying aligned with mutation semantics",
  "Overflow checks mix structural inventory with rendering heuristics"
]);

export const pptAdapterManifest = Object.freeze({
  packageName: "@officekit/ppt",
  exclusions: Object.freeze(["mcp"]),
  sourceFiles,
  publicSurface,
  capabilityFamilies,
  deferredCoreContracts,
  implementationMilestones,
  parityRisks
});

export function getPptAdapterManifest() {
  return structuredClone(pptAdapterManifest);
}

export function summarizePptAdapter() {
  return {
    packageName: pptAdapterManifest.packageName,
    surfaceCount: pptAdapterManifest.publicSurface.length,
    riskCount: pptAdapterManifest.parityRisks.length,
    previewModes: ["html", "svg"]
  };
}

const pptAdapterContract = Object.freeze({
  mutationFamilies: Object.freeze([
    "add",
    "set",
    "remove",
    "move",
    "swap",
    "copy-from",
    "raw-set",
    "batch"
  ]),
  queryFamilies: Object.freeze([
    "get",
    "query",
    "view-text",
    "view-annotated",
    "view-outline",
    "view-stats",
    "view-issues",
    "view-html",
    "view-svg",
    "check-shape-text-overflow"
  ]),
  canonicalPaths: Object.freeze([
    "/presentation",
    "/slide[1]",
    "/slide[1]/shape[1]",
    "/slide[1]/table[1]",
    "/slide[1]/placeholder[title]",
    "/slide[1]/chart[1]"
  ]),
  previewModes: Object.freeze(["html", "svg"]),
  laneNotes: Object.freeze([
    "Keep layout/theme inheritance local until shared contracts are exercised by more than one format",
    "Treat HTML/SVG preview fidelity and overflow checks as validation surfaces, not optional polish"
  ])
});

export function getPptAdapterContract() {
  return structuredClone(pptAdapterContract);
}

export function summarizePptAdapterContract() {
  return {
    canonicalPathCount: pptAdapterContract.canonicalPaths.length,
    previewModes: [...pptAdapterContract.previewModes]
  };
}
