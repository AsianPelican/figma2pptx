// The measured correction: find each Figma line in PowerPoint's PDF and derive one offset per text box.
import {test, expect} from 'bun:test';
import * as mupdf from 'mupdf';
import {measure} from '../src/measure/corrections';
import {pdfLines} from '../src/measure/pdflines';

// A 960x540 px page (720x405 pt) with Helvetica lines at given Figma-px pen positions.
function pdf(pages: {t: string, x: number, base: number, size?: number}[][]): Uint8Array {
  const doc = new mupdf.PDFDocument();
  const font = doc.addSimpleFont(new mupdf.Font('Helvetica'));
  for (const lines of pages) {
    const res = doc.addObject({Font: {F1: font}});
    const content = lines.map(l => `BT /F1 ${(l.size ?? 20) * 0.75} Tf 1 0 0 1 ${l.x * 0.75} ${405 - l.base * 0.75} Tm (${l.t}) Tj ET`).join('\n');
    doc.insertPage(-1, doc.addPage([0, 0, 720, 405], 0, res, content));
  }
  return doc.saveToBuffer('').asUint8Array();
}

const frames = [{id: '1:1', name: 'A', text: 2, shapes: 0, rasters: 0, plates: 0, rasterText: 0}, {id: '2:1', name: 'B', text: 1, shapes: 0, rasters: 0, plates: 0, rasterText: 0}];
const line = (node: string, frame: string, t: string, x: number, base: number) => ({node, frame, t, x, base, align: 'l'});

test('pdfLines reads character origins in Figma px', () => {
  const [l] = pdfLines(pdf([[{t: 'Hello', x: 100, base: 50}]]), 0, 1 / 0.75);
  expect(l.t).toBe('Hello');
  expect(l.x0).toBeCloseTo(100, 3);
  expect(l.base).toBeCloseTo(50, 3);
});

test('one offset per box brings PowerPoint back onto Figma; unmeasurable boxes get the median', () => {
  // PowerPoint drew box 1:2 2 px right and 3 px low, box 2:2 1 px high; box 1:3 is missing (rasterized text).
  const bytes = pdf([
    [{t: 'First line', x: 102, base: 103}, {t: 'Second line', x: 102, base: 133}],
    [{t: 'Other page', x: 49, base: 199}],
  ]);
  const report = {frames, textLines: [
    line('1:2', '1:1', 'First line', 100, 100), line('1:2', '1:1', 'Second line', 100, 130),
    line('1:3', '1:1', 'Not in the PDF', 300, 300),
    line('2:2', '2:1', 'Other page', 50, 200),
  ]};
  const m = measure(bytes, report, {});
  expect(m.measured).toBe(2);
  expect(m.unmeasurable).toBe(1);
  expect(m.corr['1:2'].dx).toBeCloseTo(-2, 2);
  expect(m.corr['1:2'].dy).toBeCloseTo(-3, 2);
  expect(m.corr['2:2'].dx).toBeCloseTo(1, 2);
  expect(m.corr['2:2'].dy).toBeCloseTo(1, 2);
  expect(m.corr['1:3']).toEqual({dx: 0, dy: m.medianDy});
  expect(m.measuredNodes.sort()).toEqual(['1:2', '2:2']);
  expect(m.spreadMax).toBeCloseTo(0, 2);
});

test('measuring a corrected build keeps the correction (it converges)', () => {
  // The pass-2 build already applied (-2, -3): PowerPoint now draws exactly on Figma's positions.
  const bytes = pdf([[{t: 'First line', x: 100, base: 100}]]);
  const m = measure(bytes, {frames, textLines: [line('1:2', '1:1', 'First line', 100, 100)]}, {'1:2': {dx: -2, dy: -3}});
  expect(m.corr['1:2'].dx).toBeCloseTo(-2, 2);
  expect(m.corr['1:2'].dy).toBeCloseTo(-3, 2);
});

test('leading spaces: the baseline is measured, x is not (Figma x is the pen before the spaces)', () => {
  const bytes = pdf([[{t: 'Indented', x: 120, base: 101}]]);
  const m = measure(bytes, {frames, textLines: [line('1:2', '1:1', '   Indented', 110, 100)]}, {});
  expect(m.corr['1:2'].dx).toBe(0);
  expect(m.corr['1:2'].dy).toBeCloseTo(-1, 2);
});

test('duplicate copy is associated with the matching font size before correction', () => {
  const bytes = pdf([[{t: 'Same', x: 100, base: 102, size: 20}, {t: 'Same', x: 100, base: 111, size: 40}]]);
  const report = {frames, textLines: [
    {...line('1:2', '1:1', 'Same', 100, 100), size: 20},
    {...line('1:3', '1:1', 'Same', 100, 110), size: 40},
  ]};
  const m = measure(bytes, report, {});
  expect(m.corr['1:2'].dy).toBeCloseTo(-2, 2);
  expect(m.corr['1:3'].dy).toBeCloseTo(-1, 2);
});
