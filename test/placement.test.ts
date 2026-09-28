import {test, expect} from 'bun:test';
import {placeLines} from '../src/measure/placement';
import type {PdfLine} from '../src/measure/pdflines';

const line = (t: string, x: number, base: number): PdfLine => ({
  t, x0: x, x1: x + t.length * 6, base, size: 12, font: 'Test',
  chars: [...t].map((c, i) => ({c, x: x + i * 6, y: base, x1: x + (i + 1) * 6})),
});

test('placement matches exact lines and reports component offsets against the one-pixel gate', () => {
  expect(placeLines([{t: 'Hello', x: 10, y: 20}], [line('Hello', 10.7, 19.2)])[0]).toMatchObject({status: 'placed', dx: 0.7, dy: -0.8, offsetPx: 0.8, exactLineBreak: true});
});

test('placement finds a Figma line merged into a longer PowerPoint PDF line', () => {
  const p = placeLines([{t: 'World', x: 46, y: 20}], [line('Hello World', 10, 20)])[0];
  expect(p).toMatchObject({status: 'placed', dx: 0, dy: 0, offsetPx: 0, exactLineBreak: false});
});

test('placement uses font size to distinguish duplicate copy at the same position', () => {
  const p = placeLines([{t: 'Same', x: 10, y: 25, size: 24}], [line('Same', 10, 24), {...line('Same', 10, 26), size: 24}])[0];
  expect(p).toMatchObject({status: 'placed', dy: 1});
});

test('leading-space lines verify the baseline without pretending the pen position is the first glyph', () => {
  const p = placeLines([{t: '       continuation', x: 10, y: 20, ignoreX: true}], [line('continuation', 300, 20.4)])[0];
  expect(p).toMatchObject({status: 'placed', dx: undefined, dy: 0.4, offsetPx: 0.4});
});

test('placement names unextractable text instead of counting it as an unexplained miss', () => {
  expect(placeLines([{t: 'Transparent', x: 10, y: 20}], [])[0]).toMatchObject({status: 'unmeasurable', exactLineBreak: false});
});
