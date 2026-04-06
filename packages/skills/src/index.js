import { mkdir, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { homedir } from "node:os";
import path from "node:path";

export const supportedAgents = [
  {
    key: "claude",
    aliases: ["claude", "claude-code"],
    displayName: "Claude Code",
    detectPaths: [".claude"],
    skillPaths: [".claude/skills"],
  },
  {
    key: "copilot",
    aliases: ["copilot", "github-copilot"],
    displayName: "GitHub Copilot",
    detectPaths: [".copilot"],
    skillPaths: [".copilot/skills"],
  },
  {
    key: "codex",
    aliases: ["codex", "openai-codex"],
    displayName: "Codex",
    detectPaths: [".codex", ".agents"],
    skillPaths: [".codex/skills", ".agents/skills"],
  },
  {
    key: "cursor",
    aliases: ["cursor"],
    displayName: "Cursor",
    detectPaths: [".cursor"],
    skillPaths: [".cursor/skills"],
  },
  {
    key: "windsurf",
    aliases: ["windsurf"],
    displayName: "Windsurf",
    detectPaths: [".windsurf"],
    skillPaths: [".windsurf/skills"],
  },
];

export const bundledSkills = {
  officekit: {
    description: "Base officekit skill with install, preview, and docs guidance.",
    folderName: "officekit",
    files: {
      "SKILL.md": `---
name: officekit
description: Use officekit to create, inspect, preview, and modify Office documents through the Bun + Node.js migration of OfficeCLI.
---

# officekit

This skill belongs to **officekit**, the Bun + Node.js migration of OfficeCLI.

- Prefer docs/help over guessing command/property details.
- Use preview/watch workflows for developer-visible verification.
- Use install/skills flows to bootstrap local agent integrations.
`,
    },
  },
  "officekit-docx": {
    description: "Word-focused officekit skill.",
    folderName: "officekit-docx",
    files: {
      "SKILL.md": `---
name: officekit-docx
description: Word-oriented officekit workflow guidance migrated from OfficeCLI.
---

# officekit-docx

Use this skill when the task is primarily about Word document creation or editing.
`,
    },
  },
  "officekit-xlsx": {
    description: "Excel-focused officekit skill.",
    folderName: "officekit-xlsx",
    files: {
      "SKILL.md": `---
name: officekit-xlsx
description: Excel-oriented officekit workflow guidance migrated from OfficeCLI.
---

# officekit-xlsx

Use this skill when the task is primarily about Excel workbook creation or editing.
`,
    },
  },
  "officekit-pptx": {
    description: "PowerPoint-focused officekit skill.",
    folderName: "officekit-pptx",
    files: {
      "SKILL.md": `---
name: officekit-pptx
description: PowerPoint-oriented officekit workflow guidance migrated from OfficeCLI.
---

# officekit-pptx

Use this skill when the task is primarily about PowerPoint deck creation or editing.
`,
    },
  },
};

/**
 * @param {{homeDir?: string, exists?: typeof existsSync}} [options]
 */
export function detectInstalledAgents({ homeDir = homedir(), exists = existsSync } = {}) {
  return supportedAgents
    .map((agent) => {
      const detectPath = agent.detectPaths.find((candidate) => exists(path.join(homeDir, candidate)));
      if (!detectPath) return null;

      const preferredSkillPath =
        agent.skillPaths.find((candidate) => candidate.startsWith(detectPath)) ?? agent.skillPaths[0];

      return {
        key: agent.key,
        displayName: agent.displayName,
        aliases: [...agent.aliases],
        detectPath,
        skillDir: path.join(homeDir, preferredSkillPath),
      };
    })
    .filter(Boolean);
}

export function listSkillBundles() {
  return Object.entries(bundledSkills).map(([name, bundle]) => ({
    name,
    description: bundle.description,
    folderName: bundle.folderName,
  }));
}

/**
 * @param {object} options
 * @param {string[]} [options.bundleNames]
 * @param {string} [options.homeDir]
 * @param {string[]} [options.agentKeys]
 */
export async function installSkillBundles({
  bundleNames = ["officekit"],
  homeDir = homedir(),
  agentKeys,
} = {}) {
  const detected = detectInstalledAgents({ homeDir });
  const targets = agentKeys?.length
    ? detected.filter((agent) => agentKeys.includes(agent.key) || agent.aliases.some((alias) => agentKeys.includes(alias)))
    : detected;

  /** @type {Array<{agent:string,bundle:string,files:string[]}>} */
  const installed = [];

  for (const agent of targets) {
    for (const name of bundleNames) {
      const bundle = bundledSkills[name];
      if (!bundle) {
        throw new Error(`Unknown skill bundle: ${name}`);
      }

      const bundleDir = path.join(agent.skillDir, bundle.folderName);
      await mkdir(bundleDir, { recursive: true });

      const writtenFiles = [];
      for (const [relativePath, content] of Object.entries(bundle.files)) {
        const destination = path.join(bundleDir, relativePath);
        await mkdir(path.dirname(destination), { recursive: true });
        await writeFile(destination, content, "utf8");
        writtenFiles.push(destination);
      }

      installed.push({
        agent: agent.key,
        bundle: name,
        files: writtenFiles,
      });
    }
  }

  return installed;
}
