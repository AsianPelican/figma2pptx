import {expect, test} from 'bun:test';
import {readFileSync} from 'node:fs';
import {join} from 'node:path';
import {parseRemoteArgs} from '../bridge/remote-cli.ts';

test('maps CLI selection and options to one remote conversion request', () => {
  const command = parseRemoteArgs([
    'https://www.figma.com/design/AbCdEf012345/My-Deck?node-id=12-34',
    '--pdf-preset', 'print', '--single-pass', '--no-embed-fonts', '--scale', '3', '--kern', '120',
    '-o', 'Deck copy.pptx',
  ]);
  expect(command.request).toEqual({
    fileKey: 'AbCdEf012345', frames: ['12:34'], frameNames: ['12:34'], outputName: 'Deck copy.pptx',
    pdf: 'print', passes: 1, embedFonts: false, allowFontFallback: false, offline: false, scale: 3, kern: '120',
  });
});

test('maps exact frame names and page without local cache paths', () => {
  const command = parseRemoteArgs(['AbCdEf012345', 'Cover', 'Evidence', '--page', 'Slides', '--pdf']);
  expect(command.request).toMatchObject({fileKey: 'AbCdEf012345', frames: ['Cover', 'Evidence'], page: 'Slides', pdf: 'screen'});
  expect(() => parseRemoteArgs(['AbCdEf012345', 'Cover', '--cache', '/tmp/cache'])).toThrow('cache lives on the bridge Mac');
});

test('the optional remote CLI stays a thin bridge client', () => {
  const entry = readFileSync(join(import.meta.dir, '..', 'bridge', 'remote-cli.ts'), 'utf8');
  expect(entry).toContain('remoteMain');
  expect(entry).not.toContain('convertFigma');
  expect(entry).not.toContain("../../src/index");
});
