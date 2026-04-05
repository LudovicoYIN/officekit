import { describe, expect, test } from "bun:test";
import { buildExecutionPlan, renderHelpText } from "./router.js";
import { summarizeParity } from "./parity.js";

describe("execution plan routing", () => {
  test("routes Word create to packages/word", () => {
    const plan = buildExecutionPlan(["create", "demo.docx"]);
    expect(plan.format).toBe("word");
    expect(plan.targetPackage).toBe("packages/word");
  });

  test("routes Excel create to packages/excel", () => {
    const plan = buildExecutionPlan(["create", "demo.xlsx"]);
    expect(plan.format).toBe("excel");
    expect(plan.targetPackage).toBe("packages/excel");
  });

  test("routes PowerPoint create to packages/ppt", () => {
    const plan = buildExecutionPlan(["create", "demo.pptx"]);
    expect(plan.format).toBe("powerpoint");
    expect(plan.targetPackage).toBe("packages/ppt");
  });

  test("keeps MCP explicitly excluded", () => {
    expect(() => buildExecutionPlan(["mcp"])).toThrow(/excluded/i);
  });
});

describe("parity metadata", () => {
  test("summarizes explicit lineage and exclusions", () => {
    const summary = summarizeParity();
    expect(summary.lineage).toContain("migration of OfficeCLI");
    expect(summary.excluded).toContain("mcp");
    expect(summary.capabilityCount).toBeGreaterThan(10);
  });

  test("help text includes early vertical-slice examples", () => {
    const help = renderHelpText();
    expect(help).toContain("officekit create demo.docx --plan --json");
    expect(help).toContain("Excluded by design: mcp");
  });
});
