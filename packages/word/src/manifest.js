const sourceFiles = Object.freeze([
  "OfficeCLI/src/officecli/Handlers/WordHandler.cs",
  "OfficeCLI/src/officecli/Handlers/Word/WordHandler.Add.cs",
  "OfficeCLI/src/officecli/Handlers/Word/WordHandler.FormFields.cs",
  "OfficeCLI/src/officecli/Handlers/Word/WordHandler.HtmlPreview.cs",
  "OfficeCLI/src/officecli/Handlers/Word/WordHandler.Mutations.cs",
  "OfficeCLI/src/officecli/Handlers/Word/WordHandler.Query.cs",
  "OfficeCLI/src/officecli/Handlers/Word/WordHandler.Set.cs",
  "OfficeCLI/src/officecli/Handlers/Word/WordHandler.View.cs"
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
  "ViewAsForms"
]);

const capabilityFamilies = Object.freeze({
  structure: Object.freeze([
    "body-navigation",
    "paragraphs",
    "runs",
    "tables",
    "headers-footers",
    "sections"
  ]),
  formatting: Object.freeze([
    "styles",
    "doc-defaults",
    "compatibility-settings",
    "section-layout"
  ]),
  mediaAndPreview: Object.freeze([
    "images",
    "charts",
    "html-preview",
    "form-fields"
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
  "shared-filesystem-helpers",
  "preview-server-orchestration"
]);

const implementationMilestones = Object.freeze([
  "Model document/body traversal and path resolution without locking shared-core APIs",
  "Port add/set/remove/move/copy flows for paragraphs, runs, and tables",
  "Port section layout, style, and document settings families",
  "Recreate HTML preview and forms output with fixture-backed parity checks"
]);

const parityRisks = Object.freeze([
  "Compatibility and doc-default behavior may be encoded across multiple set/navigation partials",
  "HTML preview requires chart, table, shape, and text rendering to stay aligned",
  "Form-field extraction must preserve agent-usable structured output"
]);

export const wordAdapterManifest = Object.freeze({
  packageName: "@officekit/word",
  exclusions: Object.freeze(["mcp"]),
  sourceFiles,
  publicSurface,
  capabilityFamilies,
  deferredCoreContracts,
  implementationMilestones,
  parityRisks
});

export function getWordAdapterManifest() {
  return structuredClone(wordAdapterManifest);
}

export function summarizeWordAdapter() {
  return {
    packageName: wordAdapterManifest.packageName,
    surfaceCount: wordAdapterManifest.publicSurface.length,
    riskCount: wordAdapterManifest.parityRisks.length,
    previewModes: ["html", "forms"]
  };
}

const wordAdapterContract = Object.freeze({
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
    "view-forms",
    "view-html"
  ]),
  canonicalPaths: Object.freeze([
    "/body",
    "/styles",
    "/settings",
    "/numbering",
    "/header[1]",
    "/footer[1]",
    "/chart[1]"
  ]),
  previewModes: Object.freeze(["html", "forms"]),
  laneNotes: Object.freeze([
    "Keep Word navigation and document settings local until at least one other format reuses them",
    "Treat forms output as a first-class agent-facing contract rather than a preview side effect"
  ])
});

export function getWordAdapterContract() {
  return structuredClone(wordAdapterContract);
}

export function summarizeWordAdapterContract() {
  return {
    canonicalPathCount: wordAdapterContract.canonicalPaths.length,
    previewModes: [...wordAdapterContract.previewModes]
  };
}
