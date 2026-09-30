import {copyFileSync, existsSync, mkdirSync, readFileSync} from 'node:fs';
import {homedir} from 'node:os';
import {dirname, join, resolve} from 'node:path';
import {createHash} from 'node:crypto';
import {BRIDGE_ORIGIN, parseBridgeOrigin} from './bridge/protocol.ts';
import {configurePlugin} from './configure.ts';

type RefreshOptions = {
  pluginRoot?: string;
  destination?: string;
  origin?: string;
  secretPath?: string;
  fetcher?: (input: string | URL | Request, init?: RequestInit) => Promise<Response>;
};

const FILES = ['manifest.json', 'dist/code.js', 'dist/ui.html'] as const;
const RESERVED_HOST_SUFFIXES = ['.example', '.invalid', '.test'] as const;
const digest = (path: string) => createHash('sha256').update(readFileSync(path)).digest('hex');

export async function refreshDesktop(options: RefreshOptions = {}): Promise<void> {
  const pluginRoot = resolve(options.pluginRoot ?? import.meta.dir);
  const destination = resolve(options.destination ?? process.env.FIGMA2PPTX_PLUGIN_DIR ?? join(homedir(), 'Desktop', 'figma2pptx-plugin'));
  const origin = parseBridgeOrigin(options.origin ?? process.env.FIGMA2PPTX_BRIDGE_ORIGIN ?? '');
  const hostname = new URL(origin).hostname.toLowerCase();
  if (origin === BRIDGE_ORIGIN || RESERVED_HOST_SUFFIXES.some(suffix => hostname.endsWith(suffix))) {
    throw Error('desktop refresh requires the configured live bridge origin, not a placeholder or test origin');
  }
  const secretPath = resolve(options.secretPath ?? process.env.FIGMA2PPTX_BRIDGE_SECRET_FILE ?? join(homedir(), '.config', 'figma2pptx', 'bridge-secret'));
  if (!existsSync(secretPath)) throw Error(`bridge secret not found at ${secretPath}`);
  const secret = readFileSync(secretPath, 'utf8').trim();
  if (!/^[a-f0-9]{64}$/.test(secret)) throw Error(`invalid bridge secret at ${secretPath}`);

  configurePlugin(pluginRoot, origin);
  const localRoot = join(pluginRoot, 'local');
  const manifest = JSON.parse(readFileSync(join(localRoot, 'manifest.json'), 'utf8'));
  const ui = readFileSync(join(localRoot, 'dist', 'ui.html'), 'utf8');
  if (manifest.networkAccess?.devAllowedDomains?.length !== 1 || manifest.networkAccess.devAllowedDomains[0] !== origin) {
    throw Error('generated manifest origin does not match the configured live bridge');
  }
  if (!ui.includes(`const BRIDGE = '${origin}'`) || ui.includes('bridge.example') || ui.includes(BRIDGE_ORIGIN)) {
    throw Error('generated UI origin does not match the configured live bridge');
  }

  const response = await (options.fetcher ?? fetch)(`${origin}/v1/health`, {headers: {Authorization: `Bearer ${secret}`}});
  const body = await response.json().catch(() => ({})) as {ok?: boolean};
  if (response.status !== 200 || body.ok !== true) throw Error(`authenticated live bridge health check failed with ${response.status}`);

  for (const relative of FILES) {
    const source = join(localRoot, relative);
    const target = join(destination, relative);
    mkdirSync(dirname(target), {recursive: true});
    copyFileSync(source, target);
    if (digest(source) !== digest(target)) throw Error(`desktop refresh verification failed for ${relative}`);
  }
}

if (import.meta.main) {
  try {
    await refreshDesktop();
    console.log('Refreshed the Desktop plugin after authenticated live-bridge validation.');
  } catch (error) {
    console.error(`figma2pptx desktop refresh: ${error instanceof Error ? error.message : String(error)}`);
    process.exit(1);
  }
}
