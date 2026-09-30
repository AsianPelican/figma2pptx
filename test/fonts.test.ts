// Face mapping: the legacy family PowerPoint selects by, static faces only, and every fallback reported.
import {test, expect} from 'bun:test';
import {mapFaceIn, parseFcList, type FontRow} from '../src/convert/fonts';
import {FONT_ROWS} from './helpers';

test('an installed face maps to its legacy family plus bold/italic flags', () => {
  const rows: FontRow[] = [
    {fams: ['Geist', 'Geist SemiBold'], styles: ['SemiBold', 'Regular'], ps: 'Geist-SemiBold', weight: 180, file: '/f/Geist-SemiBold.ttf', index: 0},
    {fams: ['Geist'], styles: ['Bold'], ps: 'Geist-Bold', weight: 200, file: '/f/Geist-Bold.ttf', index: 0},
  ];
  expect(mapFaceIn(rows, 'Geist-SemiBold', 'Geist', 600, false)).toEqual({typeface: 'Geist SemiBold', b: 0, i: 0, ps: 'Geist-SemiBold', status: 'exact', note: undefined, file: '/f/Geist-SemiBold.ttf', index: 0});
  expect(mapFaceIn(rows, 'Geist-Bold', 'Geist', 700, false)).toMatchObject({typeface: 'Geist', b: 1, status: 'exact'});
});

test('a variable font is never used: a family installed only as one is reported variable-only', () => {
  const f = mapFaceIn(FONT_ROWS, 'Varia-SemiBold', 'Varia', 600, false);
  expect(f.status).toBe('variable-only');
  expect(f.note).toContain('variable font');
});

test('a static face wins over a variable named instance with the same PostScript name', () => {
  const rows: FontRow[] = [
    {fams: ['Geist'], styles: ['SemiBold'], ps: 'Geist-SemiBold', weight: 180, variable: true, file: '/f/Geist[wght].ttf'},
    {fams: ['Geist', 'Geist SemiBold'], styles: ['SemiBold', 'Regular'], ps: 'Geist-SemiBold', weight: 180, file: '/f/Geist-SemiBold.ttf'},
  ];
  expect(mapFaceIn(rows, 'Geist-SemiBold', 'Geist', 600, false)).toMatchObject({typeface: 'Geist SemiBold', status: 'exact', file: '/f/Geist-SemiBold.ttf'});
});

test('without a PostScript name (Figma bundled fonts) the nearest installed weight is used and flagged', () => {
  expect(mapFaceIn(FONT_ROWS, undefined, 'Arial', 700, false)).toMatchObject({typeface: 'Arial', b: 1, status: 'exact'});
  const semi = mapFaceIn(FONT_ROWS, undefined, 'Arial', 600, false);
  expect(semi).toMatchObject({typeface: 'Arial', b: 1, status: 'substitute'});
  expect(semi.note).toContain('nearest installed face Arial-BoldMT');
  expect(mapFaceIn(FONT_ROWS, undefined, 'Arial', 400, true)).toMatchObject({typeface: 'Arial', i: 1, status: 'exact'});
});

test('localized fontconfig style aliases do not hide the English bold and italic flags', () => {
  const rows: FontRow[] = [
    {fams: ['Test Sans'], styles: ['Bold', 'Negrita', 'Lodia'], ps: 'TestSans-Bold', weight: 200},
    {fams: ['Test Sans'], styles: ['Bold Italic', 'Negrita Cursiva', 'Lodi etzana'], ps: 'TestSans-BoldItalic', weight: 200},
  ];
  expect(mapFaceIn(rows, undefined, 'Test Sans', 700, false)).toMatchObject({ps: 'TestSans-Bold', b: 1, i: 0});
  expect(mapFaceIn(rows, undefined, 'Test Sans', 700, true)).toMatchObject({ps: 'TestSans-BoldItalic', b: 1, i: 1});
});

test('a family that is not installed is reported missing', () => {
  expect(mapFaceIn(FONT_ROWS, 'Nope-Bold', 'Nope', 700, false)).toMatchObject({typeface: 'Nope', b: 1, status: 'missing'});
});

test('fc-list output: every face of a variable font file is marked variable, named instances included', () => {
  const rows = parseFcList([
    'Geist|SemiBold|Geist-SemiBold|180|False|393216|/f/Geist[wght].ttf',
    'Geist|||[0 210]|True|0|/f/Geist[wght].ttf',
    'Geist,Geist SemiBold|SemiBold,Regular|Geist-SemiBold|180|False|0|/f/Geist-SemiBold.ttf',
    'Inter|Bold|Inter-Bold|200|False|2|/f/Inter.ttc',
  ].join('\n'));
  expect(rows.map(r => r.variable)).toEqual([true, true, false, false]);
  expect(rows[0].index).toBe(0); // named-instance bits dropped; the face index remains
  expect(rows[3].index).toBe(2);
  expect(rows[2].fams).toEqual(['Geist', 'Geist SemiBold']);
});
