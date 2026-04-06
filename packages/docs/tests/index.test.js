import { describe, expect, test } from "bun:test";

import {
  extractSection,
  listDocTopics,
  resolveCommandDoc,
  searchDocs,
} from "../src/index.js";

describe("docs package", () => {
  test("resolveCommandDoc returns full command docs", async () => {
    const previewDoc = await resolveCommandDoc("preview");
    expect(previewDoc).toContain("HTML shell");
  });

  test("extractSection returns a named subsection", async () => {
    const watchBrowserFlow = await resolveCommandDoc("watch", { section: "Browser flow" });
    expect(watchBrowserFlow).toContain("SSE delivers replacement events");
  });

  test("listDocTopics reports migrated topics", async () => {
    const topics = await listDocTopics();
    expect(topics.map((entry) => entry.topic)).toContain("lineage");
  });

  test("searchDocs finds lineage references", async () => {
    const matches = await searchDocs("migrated from OfficeCLI");
    expect(matches.some((entry) => entry.path === "reference/lineage.md")).toBe(true);
  });

  test("extractSection returns null for missing headings", () => {
    expect(extractSection("# title\n\ntext", "missing")).toBeNull();
  });
});
