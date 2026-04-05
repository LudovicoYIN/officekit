import type { SupportedFormat } from "./formats.js";

export const officeCliLineageStatement =
  "officekit is a Node.js + Bun migration of OfficeCLI; v1 targets capability parity except MCP.";

export const excludedCapabilityFamilies = ["mcp"] as const;
export type ExcludedCapabilityFamily = (typeof excludedCapabilityFamilies)[number];

export const capabilityFamilies = [
  "create",
  "add",
  "set",
  "get",
  "query",
  "remove",
  "view",
  "raw",
  "watch",
  "check",
  "import",
  "batch",
  "install",
  "skills",
  "config",
  "help",
] as const;

export type CapabilityFamily = (typeof capabilityFamilies)[number];

export type PackageLane =
  | "packages/core"
  | "packages/cli"
  | "packages/word"
  | "packages/excel"
  | "packages/ppt"
  | "packages/preview"
  | "packages/skills"
  | "packages/install"
  | "packages/docs"
  | "packages/parity-tests";

export interface CapabilityLedgerEntry {
  capability: CapabilityFamily;
  ownerPackage: PackageLane | readonly PackageLane[];
  verificationLayer: "unit" | "integration" | "preview" | "docs" | "parity";
  status: "scaffolded" | "planned";
  notes: string;
}

export const capabilityLedger: readonly CapabilityLedgerEntry[] = [
  {
    capability: "create",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "scaffolded",
    notes: "CLI and shared routing contracts are scaffolded; format-specific document creation lands in format packages.",
  },
  {
    capability: "add",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "scaffolded",
    notes: "Verb contract is registered with parity-aware routing; mutations are deferred to format packages.",
  },
  {
    capability: "set",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "scaffolded",
    notes: "Shared command routing exists so property mutation semantics can be validated consistently later.",
  },
  {
    capability: "get",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "scaffolded",
    notes: "Read-oriented command contracts are wired and emit execution plans in scaffold mode.",
  },
  {
    capability: "query",
    ownerPackage: ["packages/cli", "packages/core", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "scaffolded",
    notes: "Selector/query semantics are reserved for shared core plus format adapters.",
  },
  {
    capability: "remove",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "scaffolded",
    notes: "Mutation verb is routed, but document changes remain unimplemented pending format packages.",
  },
  {
    capability: "view",
    ownerPackage: ["packages/cli", "packages/preview", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "preview",
    status: "scaffolded",
    notes: "Preview/view flows are explicitly reserved for the preview lane while keeping CLI contracts stable.",
  },
  {
    capability: "raw",
    ownerPackage: ["packages/cli", "packages/core", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "planned",
    notes: "Raw XML/query helpers remain planned; routing contract exists to preserve parity intent.",
  },
  {
    capability: "watch",
    ownerPackage: ["packages/cli", "packages/preview"],
    verificationLayer: "preview",
    status: "scaffolded",
    notes: "Watch is modeled as a preview lane capability with a stable CLI entry point.",
  },
  {
    capability: "check",
    ownerPackage: ["packages/cli", "packages/core", "packages/parity-tests"],
    verificationLayer: "parity",
    status: "scaffolded",
    notes: "Parity/check command surface is reserved for verification and inventory output.",
  },
  {
    capability: "import",
    ownerPackage: ["packages/cli", "packages/excel", "packages/ppt", "packages/word"],
    verificationLayer: "integration",
    status: "planned",
    notes: "Import remains planned until format packages land concrete handlers.",
  },
  {
    capability: "batch",
    ownerPackage: ["packages/cli", "packages/core", "packages/parity-tests"],
    verificationLayer: "parity",
    status: "planned",
    notes: "Batch execution contract is reserved for multi-step parity fixtures.",
  },
  {
    capability: "install",
    ownerPackage: ["packages/cli", "packages/install"],
    verificationLayer: "docs",
    status: "scaffolded",
    notes: "Install/update/config flow remains a first-class surface, even before implementation.",
  },
  {
    capability: "skills",
    ownerPackage: ["packages/cli", "packages/skills"],
    verificationLayer: "docs",
    status: "scaffolded",
    notes: "Skills command shape is preserved in the CLI contract for later packaging work.",
  },
  {
    capability: "config",
    ownerPackage: ["packages/cli", "packages/install"],
    verificationLayer: "docs",
    status: "scaffolded",
    notes: "Shared configuration command entrypoint is scaffolded.",
  },
  {
    capability: "help",
    ownerPackage: ["packages/cli", "packages/docs"],
    verificationLayer: "docs",
    status: "scaffolded",
    notes: "Help output carries lineage and parity scope explicitly from day one.",
  },
] as const satisfies readonly CapabilityLedgerEntry[];

export function summarizeParity() {
  return {
    lineage: officeCliLineageStatement,
    supportedFormats: ["word", "excel", "powerpoint"] satisfies readonly SupportedFormat[],
    excluded: excludedCapabilityFamilies,
    capabilityCount: capabilityLedger.length,
    scaffoldedCount: capabilityLedger.filter((entry) => entry.status === "scaffolded").length,
    plannedCount: capabilityLedger.filter((entry) => entry.status === "planned").length,
    ledger: capabilityLedger,
  };
}
