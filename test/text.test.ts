// Text-line mapping and the PowerPoint corrections: line spacing, the zero-advance trailing space, rotation.
import {test, expect, describe} from 'bun:test';
import {DOMParser} from '@xmldom/xmldom';
import {collectSegs, groupLines, alignLines, paragraphs, paraPitch, lineSpacing, zeroAdvanceTrailingSpace, unrotatedSize, textShape, type Run} from '../src/convert/text';
import {parseTransform, I} from '../src/convert/geom';
import {tableFonts} from './helpers';

const svg = (s: string) => new DOMParser().parseFromString(`<svg xmlns="http://www.w3.org/2000/svg">${s}</svg>`, 'image/svg+xml').documentElement as any;
const style = {fontFamily: 'Arial', fontPostScriptName: 'ArialMT', fontWeight: 400, fontSize: 20, letterSpacing: 0, lineHeightPx: 24, textAlignHorizontal: 'LEFT'};
const noWarn = (s: string) => { throw Error('unexpected warning: ' + s); };

describe('laid-out lines', () => {
  test('tspans group by baseline, left to right, top to bottom, across several <text> runs', () => {
    const el = svg('<g><text fill="#FF0000"><tspan x="50" y="40">world</tspan><tspan x="0" y="20">Hello </tspan></text><text fill="#0000FF"><tspan x="0" y="40">big </tspan></text></g>');
    const {segs, rotTm} = collectSegs(el, I, noWarn);
    expect(rotTm).toBeNull();
    const lines = groupLines(segs);
    expect(lines.map(l => l.map(s => s.t).join(''))).toEqual(['Hello ', 'big world']);
    expect(lines[1][0].fill).toEqual(['0000FF', 1]);
  });

  test('the element transform moves baselines into slide space', () => {
    const {segs} = collectSegs(svg('<text transform="translate(10 5)"><tspan x="1" y="2">a</tspan></text>').firstChild, I, noWarn);
    expect([segs[0].x, segs[0].y]).toEqual([11, 7]);
  });

  test('fill-opacity multiplies into the run colour', () => {
    const {segs} = collectSegs(svg('<text fill="#336699" fill-opacity="0.5"><tspan x="0" y="0">a</tspan></text>').firstChild, I, noWarn);
    expect(segs[0].fill).toEqual(['336699', 0.5]);
  });

  test('\\n starts a paragraph, U+2028 is a line break inside one, soft wraps stay in the paragraph', () => {
    const n = {characters: 'One two\nThree\u2028four five\nSix', style};
    const el = svg('<text><tspan x="0" y="20">One </tspan><tspan x="0" y="44">two&#10;</tspan><tspan x="0" y="68">Three&#x2028;</tspan><tspan x="0" y="92">four five&#10;</tspan><tspan x="0" y="116">Six</tspan></text>').firstChild;
    const L = alignLines(groupLines(collectSegs(el, I, noWarn).segs), n, '1:1', noWarn);
    expect(L.map(l => [l.runs.map(r => r.t).join(''), l.hardBefore])).toEqual([['One ', false], ['two', false], ['Three', true], ['four five', false], ['Six', true]]);
    expect(paragraphs(L).map(p => p.length)).toEqual([2, 2, 1]);
  });

  test('style overrides split a line into runs with merged styles', () => {
    const n = {characters: 'Bold plain', style, characterStyleOverrides: [1, 1, 1, 1, 0], styleOverrideTable: {1: {fontWeight: 700, fontPostScriptName: 'Arial-BoldMT'}}};
    const el = svg('<text><tspan x="0" y="20">Bold plain</tspan></text>').firstChild;
    const [line] = alignLines(groupLines(collectSegs(el, I, noWarn).segs), n, '1:1', noWarn);
    expect(line.runs.map(r => [r.t, r.style.fontPostScriptName, r.style.fontSize])).toEqual([['Bold', 'Arial-BoldMT', 20], [' plain', 'ArialMT', 20]]);
  });

  test('a line that cannot be found in the characters is reported, not dropped', () => {
    const warns: string[] = [];
    const n = {characters: 'Something else', style};
    const el = svg('<text><tspan x="0" y="20">Missing</tspan></text>').firstChild;
    const [line] = alignLines(groupLines(collectSegs(el, I, noWarn).segs), n, '9:9', w => warns.push(w));
    expect(line.runs.map(r => r.t).join('')).toBe('Missing');
    expect(warns[0]).toContain('could not align "Missing"');
  });
});

describe('line spacing (PowerPoint for Mac rounds spcPts to whole points, spcPct to whole percent)', () => {
  test('picks exact points when they land closer', () => {
    expect(lineSpacing(48, 40)).toBe('<a:spcPts val="3600"/>'); // 36 pt exactly
  });
  test('picks percent when whole points would be off', () => {
    // 17.5 px pitch = 13.125 pt: spcPts gives 13 pt (17.33 px); 17.5 / (1.2 x 14) = 104.17% -> 104% gives 17.47 px
    expect(lineSpacing(17.5, 14)).toBe('<a:spcPct val="104000"/>');
  });
  test('paragraph pitch is the median baseline step, else the style line height', () => {
    const line = (y: number) => ({runs: [{t: 'x', style, fill: null, deco: null}], x: 0, y, hardBefore: false});
    expect(paraPitch([line(0), line(20), line(41), line(61)], 20)).toBe(20);
    expect(paraPitch([line(0)], 20)).toBe(24);
    expect(paraPitch([{...line(0), runs: [{t: 'x', style: {...style, lineHeightPx: undefined}, fill: null, deco: null}]}], 20)).toBe(24);
  });
});

describe('trailing space of centred and right-aligned soft-wrapped lines', () => {
  const fonts = tableFonts();
  const runs: Run[] = [{t: 'keep their ', style, fill: null, deco: null}];
  test('split off and given a negative spacing equal to its advance', () => {
    const out = zeroAdvanceTrailingSpace(runs, 'ctr', true, fonts);
    expect(out.map(r => r.t)).toEqual(['keep their', ' ']);
    expect(out[1].spcOverride).toBeCloseTo(-0.27783 * 20, 10);
  });
  test('left-aligned lines, last lines and lines without a trailing space are untouched', () => {
    expect(zeroAdvanceTrailingSpace(runs, 'l', true, fonts)).toEqual(runs);
    expect(zeroAdvanceTrailingSpace(runs, 'r', false, fonts)).toEqual(runs);
    const none: Run[] = [{t: 'no space', style, fill: null, deco: null}];
    expect(zeroAdvanceTrailingSpace(none, 'ctr', true, fonts)).toEqual(none);
  });
});

test('unrotatedSize recovers a rotated box from its axis-aligned bounds', () => {
  const [w, h] = unrotatedSize(24, 300, parseTransform('rotate(-90)'));
  expect(w).toBeCloseTo(300, 9);
  expect(h).toBeCloseTo(24, 9);
});

test('textShape: box at the node bounds plus the measured offset, no wrap, zero insets, tracking in 1/100 pt', () => {
  const n = {name: 'Label', characters: 'Hi', absoluteBoundingBox: {x: 110, y: 220, width: 40, height: 24}, style: {...style, letterSpacing: 1.5}};
  const el = svg('<text><tspan x="10" y="39">Hi</tspan></text>').firstChild;
  const r = textShape(el, I, n, '1:1', 1, 7, {frame: {x: 100, y: 200}, corr: {dx: 0.5, dy: -1}, kern: '100', fonts: tableFonts(), warn: noWarn, onFace: () => {}})!;
  expect(r.xml).toContain('<a:off x="100013" y="180975"/>'); // (10.5, 19) px
  expect(r.xml).toContain('<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" rtlCol="0">');
  expect(r.xml).toContain('spc="113"'); // 1.5 px = 1.125 pt
  expect(r.lines).toEqual([{node: '1:1', t: 'Hi', x: 10, base: 39, align: 'l', size: 20}]);
  expect(textShape(el, I, {...n, characters: '  '}, '1:1', 1, 7, {frame: {x: 100, y: 200}, corr: {dx: 0, dy: 0}, kern: '100', fonts: tableFonts(), warn: noWarn, onFace: () => {}})).toBeNull();
});
