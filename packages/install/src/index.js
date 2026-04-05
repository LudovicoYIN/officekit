import { mkdir, readFile, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import path from "node:path";

export const defaultConfig = {
  autoUpdate: true,
  latestVersion: null,
  lastUpdateCheck: null,
  lineage: "migrated-from-officecli",
};

/**
 * @param {object} input
 * @param {"darwin"|"linux"|"win32"} input.platform
 * @param {"x64"|"arm64"} input.arch
 * @param {"gnu"|"musl"} [input.libc]
 * @param {string} [input.binaryName]
 */
export function resolveReleaseAsset({ platform, arch, libc = "gnu", binaryName = "officekit" }) {
  if (platform === "darwin") {
    return `${binaryName}-mac-${arch}`;
  }

  if (platform === "linux") {
    if (libc === "musl") {
      return `${binaryName}-linux-alpine-${arch}`;
    }
    return `${binaryName}-linux-${arch}`;
  }

  if (platform === "win32") {
    return `${binaryName}-win-${arch}.exe`;
  }

  throw new Error(`Unsupported platform: ${platform}`);
}

/**
 * @param {object} [options]
 * @param {"darwin"|"linux"|"win32"} [options.platform]
 * @param {string} [options.homeDir]
 * @param {string} [options.localAppData]
 */
export function resolveInstallDirectory({
  platform = process.platform,
  homeDir = homedir(),
  localAppData = process.env.LOCALAPPDATA,
} = {}) {
  if (platform === "win32") {
    return localAppData ? path.join(localAppData, "Officekit", "bin") : path.join(homeDir, "AppData", "Local", "Officekit", "bin");
  }

  return path.join(homeDir, ".local", "bin");
}

/**
 * @param {object} [options]
 * @param {string} [options.homeDir]
 */
export function resolveConfigPath({ homeDir = homedir() } = {}) {
  return path.join(homeDir, ".officekit", "config.json");
}

/**
 * @param {object} [options]
 * @param {string} [options.homeDir]
 */
export async function readConfig({ homeDir = homedir() } = {}) {
  const configPath = resolveConfigPath({ homeDir });
  try {
    const raw = await readFile(configPath, "utf8");
    return { ...defaultConfig, ...JSON.parse(raw) };
  } catch {
    return { ...defaultConfig };
  }
}

/**
 * @param {Record<string, unknown>} config
 * @param {object} [options]
 * @param {string} [options.homeDir]
 */
export async function writeConfig(config, { homeDir = homedir() } = {}) {
  const configPath = resolveConfigPath({ homeDir });
  await mkdir(path.dirname(configPath), { recursive: true });
  const merged = { ...defaultConfig, ...config };
  await writeFile(configPath, JSON.stringify(merged, null, 2), "utf8");
  return configPath;
}

/**
 * @param {object} options
 * @param {string} options.installDir
 * @param {string} [options.shell]
 * @param {"darwin"|"linux"|"win32"} [options.platform]
 */
export function buildPathInstruction({
  installDir,
  shell = process.env.SHELL ?? "",
  platform = process.platform,
} = {}) {
  if (platform === "win32") {
    return `Add ${installDir} to your user PATH.`;
  }

  if (shell.endsWith("/fish")) {
    return `fish_add_path ${installDir}`;
  }

  return `export PATH="${installDir}:$PATH"`;
}

/**
 * @param {object} options
 * @param {Date} [options.now]
 * @param {{lastUpdateCheck?: string | null, autoUpdate?: boolean}} [options.config]
 * @param {number} [options.intervalHours]
 */
export function shouldCheckForUpdates({
  now = new Date(),
  config = defaultConfig,
  intervalHours = 24,
} = {}) {
  if (config.autoUpdate === false) return false;
  if (!config.lastUpdateCheck) return true;

  const last = new Date(config.lastUpdateCheck);
  return now.getTime() - last.getTime() >= intervalHours * 60 * 60 * 1000;
}

/**
 * @param {object} [options]
 * @param {"darwin"|"linux"|"win32"} [options.platform]
 * @param {"x64"|"arm64"} [options.arch]
 * @param {"gnu"|"musl"} [options.libc]
 * @param {string} [options.homeDir]
 * @param {string} [options.shell]
 */
export function buildInstallPlan({
  platform = process.platform,
  arch = process.arch === "arm64" ? "arm64" : "x64",
  libc = "gnu",
  homeDir = homedir(),
  shell = process.env.SHELL ?? "",
} = {}) {
  const installDir = resolveInstallDirectory({ platform, homeDir });
  const binaryName = platform === "win32" ? "officekit.exe" : "officekit";
  return {
    binaryName,
    installDir,
    assetName: resolveReleaseAsset({ platform, arch, libc, binaryName: "officekit" }),
    configPath: resolveConfigPath({ homeDir }),
    pathInstruction: buildPathInstruction({ installDir, platform, shell }),
  };
}
