import { afterEach, describe, expect, test } from "bun:test";
import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  buildPreviewHtml,
  extractBodyHtml,
  startPreviewServer,
  startPreviewSession,
} from "../src/index.js";

const servers = [];

afterEach(async () => {
  while (servers.length > 0) {
    await servers.pop().close();
  }
});

describe("preview package", () => {
  test("buildPreviewHtml creates a self-updating shell", () => {
    const html = buildPreviewHtml({ title: "Demo", bodyHtml: "<article>Hello</article>" });
    expect(html).toContain("<title>Demo</title>");
    expect(html).toContain("EventSource(\"/events\")");
    expect(html).toContain("<article>Hello</article>");
    expect(html).toContain('id="preview-status"');
    expect(html).toContain('id="preview-refresh"');
  });

  test("preview server updates the root document through POST messages", async () => {
    const server = await startPreviewServer();
    servers.push(server);

    const before = await fetch(`${server.url}/`).then((response) => response.text());
    expect(before).toContain("Waiting for first update");

    await fetch(`${server.url}/message`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ action: "full", html: "<section data-test='preview'>updated</section>" }),
    });

    const after = await fetch(`${server.url}/`).then((response) => response.text());
    expect(extractBodyHtml(after)).toContain("data-test='preview'");

    const health = await fetch(`${server.url}/health`).then((response) => response.json());
    expect(health.ok).toBe(true);
    expect(health.version).toBe(1);
  });

  test("preview session rerenders after file changes", async () => {
    const fixtureDir = await mkdtemp(path.join(tmpdir(), "officekit-preview-"));
    const watchedFile = path.join(fixtureDir, "sample.txt");
    await writeFile(watchedFile, "alpha", "utf8");

    const session = await startPreviewSession({
      filePath: watchedFile,
      render: async (filePath) => {
        const content = await Bun.file(filePath).text();
        return `<article>${content}</article>`;
      },
    });
    servers.push(session);

    let html = await fetch(`${session.url}/`).then((response) => response.text());
    expect(extractBodyHtml(html)).toContain("<article>alpha</article>");

    await writeFile(watchedFile, "beta", "utf8");
    await Bun.sleep(180);

    html = await fetch(`${session.url}/`).then((response) => response.text());
    expect(extractBodyHtml(html)).toContain("<article>beta</article>");
  });
});
