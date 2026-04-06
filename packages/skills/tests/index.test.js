import { describe, expect, test } from "bun:test";
import { mkdir, mkdtemp, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  detectInstalledAgents,
  installSkillBundles,
  listSkillBundles,
} from "../src/index.js";

describe("skills package", () => {
  test("detectInstalledAgents finds modern codex and claude layouts", async () => {
    const homeDir = await mkdtemp(path.join(tmpdir(), "officekit-skills-"));
    await mkdir(path.join(homeDir, ".codex"), { recursive: true });
    await mkdir(path.join(homeDir, ".claude"), { recursive: true });

    const installed = detectInstalledAgents({ homeDir });
    expect(installed.map((entry) => entry.key).sort()).toEqual(["claude", "codex"]);
  });

  test("listSkillBundles exposes migrated bundle metadata", () => {
    const bundles = listSkillBundles();
    expect(bundles.map((entry) => entry.name)).toContain("officekit");
    expect(bundles.map((entry) => entry.name)).toContain("officekit-pptx");
  });

  test("installSkillBundles writes skill files into detected agent directories", async () => {
    const homeDir = await mkdtemp(path.join(tmpdir(), "officekit-skills-install-"));
    await mkdir(path.join(homeDir, ".codex"), { recursive: true });

    const installs = await installSkillBundles({
      homeDir,
      bundleNames: ["officekit", "officekit-pptx"],
    });

    expect(installs).toHaveLength(2);

    const baseSkill = await readFile(
      path.join(homeDir, ".codex", "skills", "officekit", "SKILL.md"),
      "utf8",
    );
    expect(baseSkill).toContain("migration of OfficeCLI");
  });
});
