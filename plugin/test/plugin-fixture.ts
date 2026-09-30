import {copyFileSync, mkdirSync, mkdtempSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {dirname, join} from 'node:path';

const FILES = [
  'manifest.json',
  'ui.html',
  'dist/code.js',
  'src/sha256.js',
  'src/connection.js',
] as const;

export function pluginFixture(sourceRoot: string): string {
  const root = mkdtempSync(join(tmpdir(), 'figma2pptx-plugin-test-'));
  for (const relative of FILES) {
    const target = join(root, relative);
    mkdirSync(dirname(target), {recursive: true});
    copyFileSync(join(sourceRoot, relative), target);
  }
  return root;
}
