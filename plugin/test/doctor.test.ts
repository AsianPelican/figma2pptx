import {afterEach, expect, test} from 'bun:test';
import {mkdtempSync, rmSync, writeFileSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {sshContextHasFullDiskAccess} from '../bridge/doctor.ts';

const roots: string[] = [];

afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, {recursive: true, force: true});
});

test('SSH-context Full Disk Access is informational and probeable', () => {
  const root = mkdtempSync(join(tmpdir(), 'figma2pptx-doctor-'));
  roots.push(root);
  const readable = join(root, 'TCC.db');
  writeFileSync(readable, 'x');

  expect(sshContextHasFullDiskAccess(readable)).toBe(true);
  expect(sshContextHasFullDiskAccess(join(root, 'missing.db'))).toBe(false);
});
