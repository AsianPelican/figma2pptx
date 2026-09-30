import {expect, test} from 'bun:test';
import {existsSync, readFileSync, rmSync} from 'node:fs';
import {join} from 'node:path';
import {BRIDGE_ORIGIN} from '../bridge/protocol.ts';
import {configurePlugin} from '../configure.ts';
import {pluginFixture} from './plugin-fixture.ts';

const root = join(import.meta.dir, '..');

test('manifest matches the current documented Figma development-plugin fields', () => {
  const manifest = JSON.parse(readFileSync(join(root, 'manifest.json'), 'utf8'));
  expect(manifest).toEqual({
    name: 'figma2pptx Export',
    id: 'figma2pptx-local',
    api: '1.0.0',
    main: 'dist/code.js',
    ui: 'ui.html',
    editorType: ['figma'],
    enablePrivatePluginApi: true,
    permissions: [],
    documentAccess: 'dynamic-page',
    networkAccess: {
      allowedDomains: ['none'],
      devAllowedDomains: [BRIDGE_ORIGIN],
    },
  });
  expect(existsSync(join(root, manifest.main))).toBe(true);
  expect(existsSync(join(root, manifest.ui))).toBe(true);
  expect(readFileSync(join(root, manifest.ui), 'utf8')).toContain("const BRIDGE = '__FIGMA2PPTX_BRIDGE_ORIGIN__'");
  expect(readFileSync(join(root, manifest.ui), 'utf8')).toContain("const CONFIGURED = '__FIGMA2PPTX_IS_CONFIGURED__' === 'true'");
  expect(readFileSync(join(root, manifest.ui), 'utf8')).toContain('/*__FIGMA2PPTX_SHA256__*/');
  expect(readFileSync(join(root, manifest.ui), 'utf8')).toContain('/*__FIGMA2PPTX_RECONNECT__*/');
  expect(readFileSync(join(root, manifest.ui), 'utf8')).toContain('Plugin not configured. Build it and import plugin/local/manifest.json.');
  expect(readFileSync(join(root, manifest.ui), 'utf8')).toContain('Sections, groups, and components do not count.');
  expect(readFileSync(join(root, manifest.ui), 'utf8')).toContain("if (!connected && connectionStartedFor !== state.bridgeSecret) void startConnection(state.bridgeSecret)");
  expect(readFileSync(join(root, manifest.ui), 'utf8')).toContain("setBridge(false, 'Reconnecting…')");
  expect(readFileSync(join(root, manifest.ui), 'utf8')).toContain('<div id="pairing" class="row hidden">');
});

test('local build is a self-contained Figma import folder with a real manifest filename', () => {
  const testRoot = pluginFixture(root);
  const localRoot = join(testRoot, 'local');
  const origin = configurePlugin(testRoot);
  const manifest = JSON.parse(readFileSync(join(localRoot, 'manifest.json'), 'utf8'));

  expect(manifest.main).toBe('dist/code.js');
  expect(manifest.ui).toBe('dist/ui.html');
  expect(existsSync(join(localRoot, manifest.main))).toBe(true);
  expect(existsSync(join(localRoot, manifest.ui))).toBe(true);
  expect(readFileSync(join(localRoot, manifest.ui), 'utf8')).toContain(`const BRIDGE = '${origin}'`);
  expect(readFileSync(join(localRoot, manifest.ui), 'utf8')).not.toContain('__FIGMA2PPTX_BRIDGE_ORIGIN__');
  expect(readFileSync(join(localRoot, manifest.ui), 'utf8')).toContain(`const CONFIGURED = '${origin === BRIDGE_ORIGIN ? 'false' : 'true'}' === 'true'`);
  expect(readFileSync(join(localRoot, manifest.ui), 'utf8')).not.toContain('__FIGMA2PPTX_IS_CONFIGURED__');
  expect(readFileSync(join(localRoot, manifest.ui), 'utf8')).not.toContain('/*__FIGMA2PPTX_SHA256__*/');
  expect(readFileSync(join(localRoot, manifest.ui), 'utf8')).not.toContain('/*__FIGMA2PPTX_RECONNECT__*/');
  expect(readFileSync(join(localRoot, manifest.ui), 'utf8')).toContain('figma2pptxSha256Hex');
  expect(readFileSync(join(localRoot, manifest.ui), 'utf8')).toContain('figma2pptxCreateAutoConnector');
  rmSync(testRoot, {recursive: true, force: true});
});

test('configured local build enables bridge pairing without rewriting its own check', () => {
  const testRoot = pluginFixture(root);
  const localRoot = join(testRoot, 'local');
  configurePlugin(testRoot, 'https://bridge.example:38458');
  const ui = readFileSync(join(localRoot, 'dist', 'ui.html'), 'utf8');
  expect(ui).toContain("const BRIDGE = 'https://bridge.example:38458'");
  expect(ui).toContain("const CONFIGURED = 'true' === 'true'");
  rmSync(testRoot, {recursive: true, force: true});
});
