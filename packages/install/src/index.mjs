export const POSIX_DEFAULT_INSTALL_DIR = '~/.local/bin';
export const WINDOWS_DEFAULT_INSTALL_DIR = '%USERPROFILE%\AppData\Local\Microsoft\WinGet\Links';

export function detectReleaseAsset({ platform, arch }) {
  const table = {
    'darwin:arm64': 'officekit-macos-arm64.tar.gz',
    'darwin:x64': 'officekit-macos-x64.tar.gz',
    'linux:arm64': 'officekit-linux-arm64.tar.gz',
    'linux:x64': 'officekit-linux-x64.tar.gz',
    'win32:arm64': 'officekit-windows-arm64.zip',
    'win32:x64': 'officekit-windows-x64.zip'
  };

  const asset = table[`${platform}:${arch}`];
  if (!asset) {
    throw new Error(`Unsupported platform/arch combination: ${platform}/${arch}`);
  }

  return asset;
}

export function getDefaultShellRc({ platform, env = {} } = {}) {
  if (platform === 'darwin') return '~/.zshrc';
  if (env.ZSH_VERSION) return '~/.zshrc';
  return '~/.bashrc';
}

export function buildPathExportLine(installDir = POSIX_DEFAULT_INSTALL_DIR) {
  return `export PATH="${installDir}:$PATH"`;
}

export function createInstallManifest({ version = '0.0.0-migration' } = {}) {
  return {
    productName: 'officekit',
    lineage: 'Migrated from OfficeCLI',
    version,
    releaseAssets: {
      macosArm64: detectReleaseAsset({ platform: 'darwin', arch: 'arm64' }),
      macosX64: detectReleaseAsset({ platform: 'darwin', arch: 'x64' }),
      linuxArm64: detectReleaseAsset({ platform: 'linux', arch: 'arm64' }),
      linuxX64: detectReleaseAsset({ platform: 'linux', arch: 'x64' }),
      windowsArm64: detectReleaseAsset({ platform: 'win32', arch: 'arm64' }),
      windowsX64: detectReleaseAsset({ platform: 'win32', arch: 'x64' })
    },
    postInstallSteps: ['Verify `officekit --version` once CLI packaging lands', 'Run `officekit skills install` to install the base skill surface']
  };
}

export function renderPosixInstallScript({ installDir = POSIX_DEFAULT_INSTALL_DIR } = {}) {
  return `#!/bin/sh
set -eu

# officekit is migrated from OfficeCLI and will install release assets once publishing is wired.
INSTALL_DIR=${installDir}
PATH_LINE='${buildPathExportLine(installDir)}'

echo "officekit install scaffold"
echo "Install dir: $INSTALL_DIR"
echo "Add to shell profile: $PATH_LINE"
echo "After installation, run: officekit skills install"
`;
}
