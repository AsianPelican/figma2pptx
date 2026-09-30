import {test, expect} from 'bun:test';
import {mkdirSync, mkdtempSync, rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {parsePowerPointState, powerPointBusyReason, powerPointContainer, powerPointExportFailure} from '../src/powerpoint/export';

test('PowerPoint readiness parser keeps the probe deterministic', () => {
  expect(parsePowerPointState('1|2|0|1')).toEqual({running: true, presentations: 2, modalWindows: 0, sheetWindows: 1});
  expect(() => parsePowerPointState('garbage')).toThrow('unexpected PowerPoint readiness response');
});

test('an unrelated open deck is refused explicitly', () => {
  expect(powerPointBusyReason({running: true, presentations: 1, modalWindows: 0, sheetWindows: 0})).toContain('will not touch unrelated decks');
});

test('a modal PowerPoint is refused and an idle/not-running one is available', () => {
  expect(powerPointBusyReason({running: true, presentations: 0, modalWindows: 1, sheetWindows: 0})).toContain('modal dialog');
  expect(powerPointBusyReason({running: false, presentations: 0, modalWindows: 0, sheetWindows: 0})).toBeUndefined();
});

test('export timeouts become a bounded, actionable error', () => {
  expect(powerPointExportFailure({code: 'ETIMEDOUT'}, 195000).message).toContain('within 195 seconds');
  expect(powerPointExportFailure({stderr: 'execution error: AppleEvent timed out. (-1712)'}, 195000).message).toContain('repair/permission dialog');
});

test('PowerPoint staging never falls back outside its own container', () => {
  const home = mkdtempSync(join(tmpdir(), 'figma2pptx-powerpoint-home-'));
  try {
    expect(() => powerPointContainer(home)).toThrow('refusing to fall back to /tmp');
    const container = join(home, 'Library/Containers/com.microsoft.Powerpoint/Data/tmp');
    mkdirSync(container, {recursive: true});
    expect(powerPointContainer(home)).toBe(container);
  } finally {
    rmSync(home, {recursive: true, force: true});
  }
});
