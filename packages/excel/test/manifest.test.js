import test from "node:test";
import assert from "node:assert/strict";

import {
  getExcelAdapterContract,
  getExcelAdapterManifest,
  summarizeExcelAdapterContract,
  summarizeExcelAdapter
} from "../src/index.js";

test("excel adapter manifest preserves parity-critical coverage", () => {
  const manifest = getExcelAdapterManifest();

  assert.equal(manifest.packageName, "@officekit/excel");
  assert.deepEqual(manifest.exclusions, ["mcp"]);
  assert.ok(manifest.publicSurface.includes("Import"));
  assert.ok(manifest.capabilityFamilies.calculations.includes("formulas"));
  assert.ok(manifest.capabilityFamilies.calculations.includes("charts"));
  assert.ok(manifest.capabilityFamilies.calculations.includes("pivots"));
  assert.ok(
    manifest.capabilityFamilies.workbook.includes("filtered-raw-sheet-views")
  );
});

test("excel adapter summary remains stable for lane handoff", () => {
  const summary = summarizeExcelAdapter();

  assert.equal(summary.surfaceCount, 18);
  assert.deepEqual(summary.previewModes, ["html"]);
});

test("excel adapter contract keeps workbook and chart paths explicit", () => {
  const contract = getExcelAdapterContract();
  const summary = summarizeExcelAdapterContract();

  assert.ok(contract.canonicalPaths.includes("/Sheet1"));
  assert.ok(contract.canonicalPaths.includes("/Sheet1/chart[1]"));
  assert.ok(contract.mutationFamilies.includes("import"));
  assert.equal(summary.canonicalPathCount, 7);
  assert.deepEqual(summary.previewModes, ["html"]);
});
