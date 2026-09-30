import {readFileSync, statSync} from 'node:fs';
import {join} from 'node:path';
import {startBridge} from './server.ts';
import {runtimePaths} from './runtime.ts';
import {assertBridgeDependencies} from './preflight.ts';

export function loadBridgeEnvironment(path = process.env.FIGMA2PPTX_ENV_FILE || join(runtimePaths().config, 'bridge.env')): Record<string, string> {
  const mode = statSync(path).mode & 0o777;
  if (mode & 0o077) throw Error(`${path} must be owner-only (mode 600)`);
  const values: Record<string, string> = {};
  for (const raw of readFileSync(path, 'utf8').split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith('#')) continue;
    const match = /^([A-Z][A-Z0-9_]*)=(.*)$/.exec(line);
    if (!match) throw Error(`invalid environment line in ${path}`);
    if (!['FIGMA_TOKEN', 'FIGMA2PPTX_BRIDGE_HOST'].includes(match[1])) throw Error(`unsupported environment key ${match[1]} in ${path}`);
    let value = match[2];
    if ((value.startsWith("'") && value.endsWith("'")) || (value.startsWith('"') && value.endsWith('"'))) value = value.slice(1, -1);
    if (!value) throw Error(`${match[1]} is empty in ${path}`);
    values[match[1]] = value;
  }
  if (!values.FIGMA_TOKEN) throw Error(`FIGMA_TOKEN is missing from ${path}`);
  return values;
}

if (import.meta.main) {
  try {
    const paths = runtimePaths();
    process.env.FIGMA2PPTX_RUNTIME_DIR = paths.root;
    process.env.TMPDIR = paths.temp;
    Object.assign(process.env, loadBridgeEnvironment());
    const dependencies = assertBridgeDependencies();
    console.log(`Preflight: ${Object.entries(dependencies).map(([name, path]) => `${name}=${path}`).join(', ')}`);
    startBridge();
  } catch (error) {
    console.error(`figma2pptx bridge launch: ${error instanceof Error ? error.message : String(error)}`);
    process.exit(1);
  }
}
