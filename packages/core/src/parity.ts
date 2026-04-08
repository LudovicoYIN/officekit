import type { SupportedFormat } from "./formats.js";

export const capabilityFamilies = [
  "create",
  "add",
  "set",
  "get",
  "query",
  "remove",
  "view",
  "raw",
  "raw-set",
  "add-part",
  "merge",
  "watch",
  "unwatch",
  "open",
  "close",
  "check",
  "validate",
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
  status: "scaffolded" | "planned" | "implemented";
  notes: string;
}

export const capabilityLedger: readonly CapabilityLedgerEntry[] = [
  {
    capability: "create",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "implemented",
    notes: "Format packages handle document creation; CLI routing is in place.",
  },
  {
    capability: "add",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "implemented",
    notes: "Format packages handle add mutations for paragraphs, runs, tables, and other elements.",
  },
  {
    capability: "set",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "implemented",
    notes: "Format packages handle property mutations for all supported document types.",
  },
  {
    capability: "get",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "implemented",
    notes: "Read-oriented command contracts emit execution plans backed by format package implementations.",
  },
  {
    capability: "query",
    ownerPackage: ["packages/cli", "packages/core", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "implemented",
    notes: "Selector/query semantics are backed by shared core and format adapter implementations.",
  },
  {
    capability: "remove",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "implemented",
    notes: "Format packages handle remove mutations for paragraphs, runs, tables, and other elements.",
  },
  {
    capability: "view",
    ownerPackage: ["packages/cli", "packages/preview", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "preview",
    status: "implemented",
    notes: "Preview/view flows are backed by preview lane and format package implementations.",
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
    status: "implemented",
    notes: "Watch preview server is backed by preview lane implementation.",
  },
  {
    capability: "unwatch",
    ownerPackage: ["packages/cli", "packages/preview"],
    verificationLayer: "preview",
    status: "scaffolded",
    notes: "Unwatch stops the watch preview server for a document.",
  },
  {
    capability: "open",
    ownerPackage: ["packages/cli", "packages/core"],
    verificationLayer: "integration",
    status: "scaffolded",
    notes: "Open starts a resident process to keep the document in memory for faster subsequent commands.",
  },
  {
    capability: "close",
    ownerPackage: ["packages/cli", "packages/core"],
    verificationLayer: "integration",
    status: "scaffolded",
    notes: "Close stops the resident process for the document.",
  },
  {
    capability: "raw-set",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "scaffolded",
    notes: "Raw-set modifies raw XML in a document part using XPath.",
  },
  {
    capability: "add-part",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "scaffolded",
    notes: "Add-part creates a new document part (chart, header, footer) and returns its relationship ID.",
  },
  {
    capability: "merge",
    ownerPackage: ["packages/cli", "packages/word", "packages/excel", "packages/ppt"],
    verificationLayer: "integration",
    status: "scaffolded",
    notes: "Merge replaces {{key}} placeholders in a template document with data values.",
  },
  {
    capability: "check",
    ownerPackage: ["packages/cli", "packages/core", "packages/parity-tests"],
    verificationLayer: "parity",
    status: "implemented",
    notes: "Parity/check command surface backed by verification and inventory output.",
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
    supportedFormats: ["word", "excel", "powerpoint"] satisfies readonly SupportedFormat[],
    capabilityCount: capabilityLedger.length,
    scaffoldedCount: capabilityLedger.filter((entry) => entry.status === "scaffolded").length,
    plannedCount: capabilityLedger.filter((entry) => entry.status === "planned").length,
    implementedCount: capabilityLedger.filter((entry) => entry.status === "implemented").length,
    ledger: capabilityLedger,
  };
}
