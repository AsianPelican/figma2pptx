// Text-line mapping and the PowerPoint corrections: line spacing, the zero-advance trailing space, rotation.
import {test, expect, describe} from 'bun:test';
import {DOMParser} from '@xmldom/xmldom';
import {collectSegs, groupLines, sourceOrderedSegs, alignLines, paragraphs, paraPitch, lineSpacing, zeroAdvanceTrailingSpace, unrotatedSize, drawingmlCharacterSpacing, drawingmlKerning, textShape, measurableText, type Run} from '../src/convert/text';
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

  test('missing fill-opacity means fully opaque', () => {
    const {segs} = collectSegs(svg('<text fill="#336699"><tspan x="0" y="0">a</tspan></text>').firstChild, I, noWarn);
    expect(segs[0].fill).toEqual(['336699', 1]);
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

  test('per-character spacing overrides split the source into independently spaced runs', () => {
    const n = {
      characters: 'ABCD', style,
      characterStyleOverrides: [0, 1, 2, 0],
      styleOverrideTable: {
        1: {letterSpacing: {value: -8, unit: 'PERCENT'}},
        2: {letterSpacing: {value: 1.5, unit: 'PIXELS'}},
      },
    };
    const el = svg('<text><tspan x="0" y="20">ABCD</tspan></text>').firstChild;
    const [line] = alignLines(groupLines(collectSegs(el, I, noWarn).segs), n, '1:2', noWarn);
    expect(line.runs.map(r => [r.t, r.style.letterSpacing])).toEqual([
      ['A', 0],
      ['B', {value: -8, unit: 'PERCENT'}],
      ['C', {value: 1.5, unit: 'PIXELS'}],
      ['D', 0],
    ]);
  });

  test('styled SVG fragments are emitted in source character order', () => {
    const characters = 'Alpha (Beta, Gamma.)';
    const overrides = [...characters].map(c => c === 'B' || c === 'G' || c === ')' ? 0 : 1);
    const n = {characters, style, characterStyleOverrides: overrides, styleOverrideTable: {1: {fontWeight: 700, fontPostScriptName: 'Arial-BoldMT'}}};
    const el = svg('<g><text><tspan x="80" y="20">B</tspan><tspan x="130" y="20">G</tspan><tspan x="0" y="20">)</tspan></text><text><tspan x="0" y="20">Alpha (</tspan><tspan x="90" y="20">eta, </tspan><tspan x="140" y="20">amma.</tspan></text></g>');
    const [line] = alignLines(groupLines(collectSegs(el, I, noWarn).segs), n, '1:1', noWarn);
    expect(line.runs.map(r => r.t)).toEqual(['Alpha (', 'B', 'eta, ', 'G', 'amma.', ')']);
    expect(line.runs.map(r => r.style.fontPostScriptName)).toEqual(['Arial-BoldMT', 'ArialMT', 'Arial-BoldMT', 'ArialMT', 'Arial-BoldMT', 'ArialMT']);
    expect(line.runs.map(r => r.t).join('')).toBe(characters);
  });

  test('source ordering backtracks when a longer prefix is the wrong fragment', () => {
    const seg = (t: string): any => ({x: 0, y: 0, t, fill: null, deco: null});
    expect(sourceOrderedSegs([seg('a'), seg('ab'), seg('bc')], 'abcab')?.map(s => s.t)).toEqual(['a', 'bc', 'ab']);
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

describe('character spacing and kerning', () => {
  test('converts REST pixel spacing and plugin PIXELS/PERCENT spacing to 1/100 point', () => {
    expect(drawingmlCharacterSpacing({...style, letterSpacing: 1.5})).toBe(113);
    expect(drawingmlCharacterSpacing({...style, letterSpacing: {value: 1.5, unit: 'PIXELS'}})).toBe(113);
    expect(drawingmlCharacterSpacing({...style, letterSpacing: {value: -8, unit: 'PERCENT'}})).toBe(-120);
    expect(drawingmlCharacterSpacing({...style, letterSpacing: -8, letterSpacingUnit: 'PERCENT'})).toBe(-120);
  });

  test('uses Figma KERN when present and otherwise retains the converter default', () => {
    expect(drawingmlKerning(style, '100')).toBe('100');
    expect(drawingmlKerning({...style, opentypeFlags: {KERN: 0}}, '100')).toBe('0');
    expect(drawingmlKerning({...style, opentypeFlags: {KERN: 1}}, '0')).toBe('100');
    expect(drawingmlKerning({...style, fontFeatures: {KERN: false}}, '100')).toBe('0');
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
    expect(paraPitch([{...line(0), runs: []}], 20)).toBe(24);
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

test('textShape emits per-run percentage/pixel spacing and Figma kerning flags', () => {
  const n = {
    name: 'Adjusted letters', characters: 'ABCD',
    absoluteBoundingBox: {x: 0, y: 0, width: 100, height: 24}, style,
    characterStyleOverrides: [0, 1, 2, 3],
    styleOverrideTable: {
      1: {letterSpacing: {value: -8, unit: 'PERCENT'}},
      2: {letterSpacing: {value: 1.5, unit: 'PIXELS'}, opentypeFlags: {KERN: 0}},
      3: {opentypeFlags: {KERN: 1}},
    },
  };
  const el = svg('<text><tspan x="0" y="20">ABCD</tspan></text>').firstChild;
  const r = textShape(el, I, n, '4:1', 1, 11, {frame: {x: 0, y: 0}, corr: {dx: 0, dy: 0}, kern: '0', fonts: tableFonts(), warn: noWarn, onFace: () => {}})!;
  const props = [...r.xml.matchAll(/<a:rPr [^>]+/g)].map(match => match[0]);
  expect(props.map(value => value.match(/spc="(-?\d+)"/)?.[1])).toEqual(['0', '-120', '113', '0']);
  expect(props.map(value => value.match(/kern="(\d+)"/)?.[1])).toEqual(['0', '0', '0', '100']);
});

test('textShape maps Figma list type, level, hanging indent, and paragraph spacing', () => {
  const n = {
    name: 'Synthetic list', characters: 'Bullet\nNested\nFirst\nSecond\nPlain',
    absoluteBoundingBox: {x: 100, y: 200, width: 300, height: 150},
    style: {...style, listSpacing: 4, paragraphSpacing: 8},
    lineTypes: ['UNORDERED', 'UNORDERED', 'ORDERED', 'ORDERED', 'NONE'],
    lineIndentations: [1, 2, 1, 1, 0],
  };
  const el = svg('<text><tspan x="120" y="224">Bullet&#10;</tspan><tspan x="140" y="248">Nested&#10;</tspan><tspan x="120" y="272">First&#10;</tspan><tspan x="120" y="296">Second&#10;</tspan><tspan x="100" y="320">Plain</tspan></text>').firstChild;
  const r = textShape(el, I, n, '2:1', 1, 9, {frame: {x: 0, y: 0}, corr: {dx: 0, dy: 0}, kern: '100', fonts: tableFonts(), warn: noWarn, onFace: () => {}})!;
  expect(r.xml).toContain('<a:pPr algn="l" lvl="0" marL="190500" indent="-190500">');
  expect(r.xml).toContain('<a:pPr algn="l" lvl="1" marL="381000" indent="-190500">');
  expect(r.xml.match(/<a:buChar char="•"\/>/g)).toHaveLength(2);
  expect(r.xml.match(/<a:buAutoNum type="arabicPeriod"\/>/g)).toHaveLength(2);
  expect(r.xml.match(/<a:buSzPct val="100000"\/>/g)).toHaveLength(4);
  expect(r.xml.match(/<a:spcAft><a:spcPts val="300"\/><\/a:spcAft>/g)).toHaveLength(3);
  expect(r.xml).toContain('<a:spcAft><a:spcPts val="600"/></a:spcAft><a:buFont typeface="Arial"/><a:buSzPct val="100000"/><a:buAutoNum type="arabicPeriod"/>');
});

test('textShape preserves a blank source line but excludes it from placement records', () => {
  const n = {name: 'Synthetic blank line', characters: 'Top\n\nBottom', absoluteBoundingBox: {x: 0, y: 0, width: 200, height: 72}, style, lineTypes: ['NONE', 'NONE', 'NONE']};
  const el = svg('<text><tspan x="0" y="20">Top&#10;</tspan><tspan x="0" y="44">&#10;</tspan><tspan x="0" y="68">Bottom</tspan></text>').firstChild;
  const r = textShape(el, I, n, '3:1', 1, 10, {frame: {x: 0, y: 0}, corr: {dx: 0, dy: 0}, kern: '100', fonts: tableFonts(), warn: noWarn, onFace: () => {}})!;
  expect(r.xml.match(/<a:p>/g)).toHaveLength(2);
  expect(r.xml.match(/<a:br>/g)).toHaveLength(1);
  expect(r.lines.map(l => l.t)).toEqual(['Top', 'Bottom']);
});

test('a line with an outlined missing glyph is measured on its first native-text stretch', () => {
  // Hand-authored fragments give the glyph, space, and letters independent positions.
  const n = {characters: '≈ ABC', characterStyleOverrides: [], styleOverrideTable: {}, style};
  const el = svg('<g><text><tspan x="10" y="20">≈</tspan></text><text><tspan x="25" y="20"> </tspan></text><text><tspan x="30" y="20">ABC</tspan></text></g>');
  const [l] = alignLines(groupLines(collectSegs(el, I, noWarn).segs), n, '1:1', noWarn);
  expect(measurableText(l.runs, new Map([[0, {}]]), l.x)).toEqual({t: 'ABC', x: 30});
  // A span that starts with the space: the digits' pen x is unknown, so only the baseline is measured.
  const merged = svg('<g><text><tspan x="40" y="50">≈</tspan><tspan x="55" y="50"> 42</tspan></text></g>');
  const [m] = alignLines(groupLines(collectSegs(merged, I, noWarn).segs), {...n, characters: '≈ 42'}, '1:1', noWarn);
  expect(measurableText(m.runs, new Map([[0, {}]]), m.x)).toEqual({t: '42', x: 40, ignoreX: true});
  // Native text before the glyph keeps the line's own pen x; a glyph-only line has nothing to measure.
  const run = (t: string, sourceStart: number): Run => ({t, style, fill: null, deco: null, sourceStart, spanStart: sourceStart, spanX: 10 + sourceStart * 8});
  expect(measurableText([run('about ', 0), run('≈', 6), run(' ABC', 7)], new Map([[6, {}]]), 10)).toEqual({t: 'about ', x: 10});
  expect(measurableText([run('≈', 0)], new Map([[0, {}]]), 10)).toBeUndefined();
});
