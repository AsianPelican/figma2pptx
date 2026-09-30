import {accessSync} from 'node:fs';
import {delimiter, join} from 'node:path';
import {constants} from 'node:fs';

export const SERVICE_PATH = '/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin';
export const BRIDGE_EXECUTABLES = ['magick', 'fc-list', 'osascript', 'pdfimages', 'pdftoppm'] as const;

export function findExecutable(name: string, path = process.env.PATH ?? ''): string | undefined {
  for (const dir of path.split(delimiter).filter(Boolean)) {
    const candidate = join(dir, name);
    try { accessSync(candidate, constants.X_OK); return candidate; }
    catch {}
  }
}

export function assertBridgeDependencies(path = process.env.PATH ?? '', required: readonly string[] = BRIDGE_EXECUTABLES): Record<string, string> {
  const found: Record<string, string> = {};
  const missing: string[] = [];
  for (const name of required) {
    const executable = findExecutable(name, path);
    if (executable) found[name] = executable;
    else missing.push(name);
  }
  if (missing.length) {
    throw Error(`missing required executable${missing.length === 1 ? '' : 's'}: ${missing.join(', ')}. Install ImageMagick, fontconfig, and Poppler, and ensure the launch agent PATH includes their bin directories (PATH=${path || '(empty)'}).`);
  }
  return found;
}
