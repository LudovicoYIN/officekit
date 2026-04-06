import test from 'node:test';
import assert from 'node:assert/strict';
import { buildPathExportLine, createInstallManifest, detectReleaseAsset, getDefaultShellRc, renderPosixInstallScript } from '../src/index.mjs';

test('install manifest preserves platform asset coverage and lineage', () => {
  const manifest = createInstallManifest({ version: '1.0.0-canary' });
  assert.equal(manifest.productName, 'officekit');
  assert.match(manifest.lineage, /OfficeCLI/);
  assert.equal(manifest.releaseAssets.linuxX64, 'officekit-linux-x64.tar.gz');
  assert.equal(detectReleaseAsset({ platform: 'darwin', arch: 'arm64' }), 'officekit-macos-arm64.tar.gz');
});

test('install helpers emit shell profile guidance', () => {
  assert.equal(getDefaultShellRc({ platform: 'darwin' }), '~/.zshrc');
  assert.equal(getDefaultShellRc({ platform: 'linux', env: {} }), '~/.bashrc');
  assert.equal(buildPathExportLine('~/.bun/bin'), 'export PATH="~/.bun/bin:$PATH"');
  assert.match(renderPosixInstallScript(), /officekit install scaffold/);
  assert.match(renderPosixInstallScript(), /officekit skills install/);
  assert.throws(() => detectReleaseAsset({ platform: 'freebsd', arch: 'x64' }), /Unsupported platform/);
});
