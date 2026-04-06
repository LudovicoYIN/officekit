import test from "node:test";
import assert from "node:assert/strict";

import {
  getPptAdapterContract,
  getPptAdapterManifest,
  summarizePptAdapter,
  summarizePptAdapterContract
} from "../src/index.js";

test("ppt adapter manifest preserves parity-critical coverage", () => {
  const manifest = getPptAdapterManifest();

  assert.equal(manifest.packageName, "@officekit/ppt");
  assert.deepEqual(manifest.exclusions, ["mcp"]);
  assert.ok(manifest.publicSurface.includes("ViewAsSvg"));
  assert.ok(manifest.publicSurface.includes("CheckShapeTextOverflow"));
  assert.ok(manifest.capabilityFamilies.slides.includes("layout-inheritance"));
  assert.ok(
    manifest.capabilityFamilies.renderingAndValidation.includes("animations")
  );
  assert.ok(manifest.capabilityFamilies.shapesAndMedia.includes("media"));
});

test("ppt adapter summary remains stable for lane handoff", () => {
  const summary = summarizePptAdapter();

  assert.equal(summary.surfaceCount, 20);
  assert.deepEqual(summary.previewModes, ["html", "svg"]);
});

test("ppt adapter contract keeps slide and preview semantics explicit", () => {
  const contract = getPptAdapterContract();
  const summary = summarizePptAdapterContract();

  assert.ok(contract.canonicalPaths.includes("/slide[1]/shape[1]"));
  assert.ok(contract.canonicalPaths.includes("/slide[1]/placeholder[title]"));
  assert.ok(contract.queryFamilies.includes("check-shape-text-overflow"));
  assert.equal(summary.canonicalPathCount, 6);
  assert.deepEqual(summary.previewModes, ["html", "svg"]);
});
