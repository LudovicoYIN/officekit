import { describe, expect, test } from "bun:test";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  buildInstallPlan,
  buildPathInstruction,
  readConfig,
  resolveReleaseAsset,
  shouldCheckForUpdates,
  writeConfig,
} from "../src/index.js";

describe("install package", () => {
  test("resolveReleaseAsset matches platform naming expectations", () => {
    expect(resolveReleaseAsset({ platform: "darwin", arch: "arm64" })).toBe("officekit-mac-arm64");
    expect(resolveReleaseAsset({ platform: "linux", arch: "x64", libc: "musl" })).toBe("officekit-linux-alpine-x64");
    expect(resolveReleaseAsset({ platform: "win32", arch: "arm64" })).toBe("officekit-win-arm64.exe");
  });

  test("buildPathInstruction uses fish-specific syntax", () => {
    expect(buildPathInstruction({ installDir: "/tmp/bin", shell: "/usr/bin/fish", platform: "linux" })).toBe(
      "fish_add_path /tmp/bin",
    );
  });

  test("config helpers persist config and respect update staleness", async () => {
    const homeDir = await mkdtemp(path.join(tmpdir(), "officekit-install-"));
    await writeConfig({ lastUpdateCheck: "2026-04-01T00:00:00.000Z" }, { homeDir });
    const config = await readConfig({ homeDir });
    expect(config.lineage).toBe("migrated-from-officecli");
    expect(
      shouldCheckForUpdates({
        config,
        now: new Date("2026-04-05T00:00:00.000Z"),
      }),
    ).toBe(true);
  });

  test("buildInstallPlan exposes install and config targets", () => {
    const plan = buildInstallPlan({
      platform: "linux",
      arch: "x64",
      homeDir: "/tmp/officekit-home",
      shell: "/bin/zsh",
    });

    expect(plan.installDir).toBe("/tmp/officekit-home/.local/bin");
    expect(plan.pathInstruction).toContain('export PATH="/tmp/officekit-home/.local/bin:$PATH"');
    expect(plan.configPath).toBe("/tmp/officekit-home/.officekit/config.json");
  });
});
