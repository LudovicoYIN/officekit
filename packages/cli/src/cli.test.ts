import { describe, expect, test } from "bun:test";
import { runCli } from "./index.js";

describe("officekit CLI scaffold", () => {
  test("returns a JSON execution plan for Word create", () => {
    const result = runCli(["create", "demo.docx", "--plan", "--json"]);
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain('"targetPackage": "packages/word"');
  });

  test("returns lineage summary for about", () => {
    const result = runCli(["about"]);
    expect(result.exitCode).toBe(0);
    expect(result.stdout).toContain("migration of OfficeCLI");
  });

  test("keeps unsupported MCP explicit", () => {
    const result = runCli(["mcp", "--json"]);
    expect(result.exitCode).toBe(1);
    expect(result.stderr).toContain("capability_excluded");
  });
});
