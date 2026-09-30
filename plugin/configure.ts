import {copyFileSync, mkdirSync, readFileSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {BRIDGE_ORIGIN, parseBridgeOrigin} from './bridge/protocol.ts';

const PLACEHOLDER = '__FIGMA2PPTX_BRIDGE_ORIGIN__';
const CONFIGURED_PLACEHOLDER = '__FIGMA2PPTX_IS_CONFIGURED__';
const SHA256_PLACEHOLDER = '/*__FIGMA2PPTX_SHA256__*/';
const RECONNECT_PLACEHOLDER = '/*__FIGMA2PPTX_RECONNECT__*/';

export function configurePlugin(root = import.meta.dir, value = process.env.FIGMA2PPTX_BRIDGE_ORIGIN || BRIDGE_ORIGIN): string {
  const origin = parseBridgeOrigin(value);
  const manifest = JSON.parse(readFileSync(join(root, 'manifest.json'), 'utf8'));
  const ui = readFileSync(join(root, 'ui.html'), 'utf8');
  const sha256 = readFileSync(join(root, 'src', 'sha256.js'), 'utf8');
  const reconnect = readFileSync(join(root, 'src', 'connection.js'), 'utf8');
  const localRoot = join(root, 'local');
  const localDist = join(localRoot, 'dist');
  if (!ui.includes(PLACEHOLDER)) throw Error(`UI template is missing ${PLACEHOLDER}`);
  if (!ui.includes(CONFIGURED_PLACEHOLDER)) throw Error(`UI template is missing ${CONFIGURED_PLACEHOLDER}`);
  if (!ui.includes(SHA256_PLACEHOLDER)) throw Error(`UI template is missing ${SHA256_PLACEHOLDER}`);
  if (!ui.includes(RECONNECT_PLACEHOLDER)) throw Error(`UI template is missing ${RECONNECT_PLACEHOLDER}`);
  manifest.ui = 'dist/ui.html';
  manifest.networkAccess.devAllowedDomains = [origin];
  mkdirSync(localDist, {recursive: true});
  copyFileSync(join(root, 'dist', 'code.js'), join(localDist, 'code.js'));
  writeFileSync(join(localDist, 'ui.html'), ui
    .replaceAll(PLACEHOLDER, origin)
    .replaceAll(CONFIGURED_PLACEHOLDER, String(origin !== BRIDGE_ORIGIN))
    .replace(SHA256_PLACEHOLDER, sha256)
    .replace(RECONNECT_PLACEHOLDER, reconnect));
  writeFileSync(join(localRoot, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
  return origin;
}

if (import.meta.main) {
  const origin = configurePlugin();
  console.log(origin === BRIDGE_ORIGIN
    ? 'Built placeholder plugin files; set FIGMA2PPTX_BRIDGE_ORIGIN and rebuild before importing into Figma.'
    : 'Built locally configured plugin files. Import plugin/local/manifest.json in Figma Desktop.');
}
