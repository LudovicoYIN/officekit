import test from "node:test";
import assert from "node:assert/strict";

import {
  getWordAdapterContract,
  getWordAdapterManifest,
  summarizeWordAdapterContract,
  summarizeWordAdapter
} from "../src/index.js";

test("word adapter manifest preserves parity-critical coverage", () => {
  const manifest = getWordAdapterManifest();

  assert.equal(manifest.packageName, "@officekit/word");
  assert.deepEqual(manifest.exclusions, ["mcp"]);
  assert.ok(manifest.publicSurface.includes("ViewAsHtml"));
  assert.ok(manifest.publicSurface.includes("ViewAsForms"));
  assert.ok(manifest.capabilityFamilies.structure.includes("tables"));
  assert.ok(
    manifest.capabilityFamilies.formatting.includes("compatibility-settings")
  );
  assert.ok(
    manifest.capabilityFamilies.mediaAndPreview.includes("form-fields")
  );
});

test("word adapter summary remains stable for lane handoff", () => {
  const summary = summarizeWordAdapter();

  assert.equal(summary.surfaceCount, 19);
  assert.deepEqual(summary.previewModes, ["html", "forms"]);
});

test("word adapter contract keeps path and preview expectations explicit", () => {
  const contract = getWordAdapterContract();
  const summary = summarizeWordAdapterContract();

  assert.ok(contract.canonicalPaths.includes("/body"));
  assert.ok(contract.canonicalPaths.includes("/chart[1]"));
  assert.ok(contract.queryFamilies.includes("view-forms"));
  assert.equal(summary.canonicalPathCount, 7);
  assert.deepEqual(summary.previewModes, ["html", "forms"]);
});
