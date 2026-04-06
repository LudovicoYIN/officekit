import { readdir, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const packageDir = path.dirname(fileURLToPath(import.meta.url));
const defaultDocsDir = path.join(packageDir, "..", "content");

const commandFiles = {
  config: "commands/config.md",
  help: "commands/help.md",
  install: "commands/install.md",
  preview: "commands/preview.md",
  skills: "commands/skills.md",
  watch: "commands/watch.md",
  lineage: "reference/lineage.md",
};

/**
 * @param {string} command
 * @param {{docsDir?: string, section?: string}} [options]
 */
export async function resolveCommandDoc(command, { docsDir = defaultDocsDir, section } = {}) {
  const relativePath = commandFiles[command];
  if (!relativePath) {
    return null;
  }

  const content = await readFile(path.join(docsDir, relativePath), "utf8");
  if (!section) {
    return content;
  }

  return extractSection(content, section);
}

/**
 * @param {{docsDir?: string}} [options]
 */
export async function listDocTopics({ docsDir = defaultDocsDir } = {}) {
  const results = [];

  for (const [topic, relativePath] of Object.entries(commandFiles)) {
    const fullPath = path.join(docsDir, relativePath);
    const content = await readFile(fullPath, "utf8");
    results.push({
      topic,
      path: relativePath,
      title: content.split("\n")[0].replace(/^#\s*/, "").trim(),
    });
  }

  return results;
}

/**
 * @param {string} query
 * @param {{docsDir?: string}} [options]
 */
export async function searchDocs(query, { docsDir = defaultDocsDir } = {}) {
  const lowered = query.toLowerCase();
  const results = [];
  const directories = await readdir(docsDir, { withFileTypes: true });

  for (const entry of directories) {
    if (!entry.isDirectory()) continue;
    const nestedDir = path.join(docsDir, entry.name);
    const nestedEntries = await readdir(nestedDir, { withFileTypes: true });

    for (const nestedEntry of nestedEntries) {
      if (!nestedEntry.isFile()) continue;
      const relativePath = path.join(entry.name, nestedEntry.name);
      const content = await readFile(path.join(docsDir, relativePath), "utf8");
      if (content.toLowerCase().includes(lowered)) {
        results.push({
          path: relativePath,
          title: content.split("\n")[0].replace(/^#\s*/, "").trim(),
        });
      }
    }
  }

  return results;
}

/**
 * @param {string} content
 * @param {string} heading
 */
export function extractSection(content, heading) {
  const lines = content.split("\n");
  const normalizedHeading = heading.trim().toLowerCase();
  const collected = [];
  let activeLevel = 0;
  let active = false;

  for (const line of lines) {
    const match = /^(#{1,6})\s+(.*)$/.exec(line);
    if (!match) {
      if (active) {
        collected.push(line);
      }
      continue;
    }

    const level = match[1].length;
    const title = match[2].trim().toLowerCase();

    if (active && level <= activeLevel) {
      break;
    }

    if (!active && title === normalizedHeading) {
      active = true;
      activeLevel = level;
      collected.push(line);
    }
  }

  return collected.length > 0 ? collected.join("\n").trim() : null;
}
