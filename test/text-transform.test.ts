import {test, expect} from 'bun:test';
import {mkdtempSync, rmSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {strFromU8, unzipSync} from 'fflate';
import {buildDeck, type FrameSource} from '../src/convert/build';
import {textTransformPlan, transformStyle, type TextTransformOptions} from '../src/convert/text-transform';
import {fakeImages, FIXTURE_CACHE, FIXTURE_KEY, tableFonts} from './helpers';

const style = (family: string, ps: string | undefined, weight: number) => ({
  fontFamily: family, fontPostScriptName: ps, fontWeight: weight, fontSize: 24, italic: false,
  letterSpacing: 0, lineHeightPx: 28.8, textAlignHorizontal: 'LEFT',
});
const text = (id: string, name: string, y: number, characters: string, base: any, override?: any) => ({
  id, name, type: 'TEXT', visible: true, absoluteBoundingBox: {x: 40, y, width: 120, height: 80},
  absoluteRenderBounds: {x: 40, y, width: 120, height: 80}, characters, style: base,
  characterStyleOverrides: override ? [0, 1] : [], styleOverrideTable: override ? {1: override} : {}, lineTypes: ['NONE'],
});

const options: TextTransformOptions = {
  outline: s => s.fontFamily === 'Display Face',
  substitute: s => s.fontFamily === 'Legacy Sans' ? {...s, fontFamily: 'Replacement Sans', fontPostScriptName: undefined} : null,
};

const replacementFonts = tableFonts([
  {fams: ['Replacement Sans'], styles: ['Regular'], ps: 'ReplacementSans-Regular', weight: 80},
  {fams: ['Replacement Sans'], styles: ['Bold'], ps: 'ReplacementSans-Bold', weight: 200},
]);

test('host text policy outlines a whole node and substitutes configured live runs', () => {
  const outlined = text('1:2', 'Display', 20, 'AB', style('Display Face', 'Display-Regular', 400), {fontPostScriptName: 'Display-Semibold', fontWeight: 600});
  expect(textTransformPlan(outlined, options)).toEqual({outlinedRuns: 2, substitutedRuns: 0, outline: true});
  const mapped = transformStyle(style('Legacy Sans', 'Legacy-Bold', 700), options);
  expect(mapped.style).toMatchObject({fontFamily: 'Replacement Sans', fontPostScriptName: undefined, fontWeight: 700});
  expect(mapped.substituted).toBe(true);
  const mixed = text('1:3', 'Mixed', 20, 'AB', style('Display Face', 'Display-Regular', 400), {fontFamily: 'Replacement Sans'});
  expect(() => textTransformPlan(mixed, options)).toThrow('split the outlined text into its own Figma text node');
});

test('host text policy emits SVG with PNG fallback and keeps substitutions editable', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'figma2pptx-transform-test-'));
  const svg = join(dir, 'outlined.svg');
  writeFileSync(svg, '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80" viewBox="0 0 120 80"><path d="M0 0H120V80H0Z" fill="#123456"/></svg>');
  const fallback = join(FIXTURE_CACHE, FIXTURE_KEY, '1001', 'raster', '1-10@2.png');
  const outlined = text('1:2', 'Display', 20, 'AB', style('Display Face', 'Display-Regular', 400));
  const legacy = text('1:3', 'Legacy', 120, 'CD', style('Legacy Sans', 'Legacy-Regular', 400), {fontWeight: 700});
  const frame = {id: '1:1', name: 'Synthetic transform', type: 'FRAME', visible: true, clipsContent: true,
    absoluteBoundingBox: {x: 0, y: 0, width: 960, height: 540}, fills: [], strokes: [], effects: [], children: [outlined, legacy]};
  const source: FrameSource = {
    document: () => frame,
    svg: () => '<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540" viewBox="0 0 960 540"><g data-node-id="1:2"><text><tspan x="40" y="50">AB</tspan></text></g><g data-node-id="1:3"><text><tspan x="40" y="150">CD</tspan></text></g></svg>',
    ensureRasters: async () => {}, rasterFile: () => fallback,
    ensureOutlinedText: async ids => expect(ids).toEqual(['1:2']), outlinedTextFile: () => svg, outlinedTextFallbackFile: () => fallback,
    platePath: () => join(dir, 'plate.png'),
  };
  try {
    const {pptx, report} = await buildDeck(source, ['1:1'], {scale: 2, kern: '100', corr: {}, textTransform: options}, {fonts: replacementFonts, images: fakeImages});
    const zip = unzipSync(pptx), slide = strFromU8(zip['ppt/slides/slide1.xml']);
    expect(slide).toContain('<asvg:svgBlip');
    expect(slide).not.toContain('<a:t>AB</a:t>');
    expect(slide).toContain('<a:t>C</a:t>');
    expect(slide).toContain('<a:t>D</a:t>');
    expect(slide).toContain('typeface="Replacement Sans"');
    expect(slide).toContain(' b="1"');
    expect(Object.keys(zip).filter(k => k.startsWith('ppt/media/')).sort()).toEqual(['ppt/media/image1.png', 'ppt/media/image2.svg']);
    expect(report.frames[0].textTransform).toEqual({outlinedRuns: 1, substitutedRuns: 2});
  } finally { rmSync(dir, {recursive: true, force: true}); }
});

test('an unmatched host policy performs no outline fetch and leaves PPTX bytes unchanged', async () => {
  const fallback = join(FIXTURE_CACHE, FIXTURE_KEY, '1001', 'raster', '1-10@2.png');
  const ordinary = text('1:2', 'Ordinary', 20, 'AB', style('Replacement Sans', 'ReplacementSans-Regular', 400));
  const frame = {id: '1:1', name: 'Synthetic no-op', type: 'FRAME', visible: true, clipsContent: true,
    absoluteBoundingBox: {x: 0, y: 0, width: 960, height: 540}, fills: [], strokes: [], effects: [], children: [ordinary]};
  let outlineCalls = 0;
  const source: FrameSource = {
    document: () => frame,
    svg: () => '<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540" viewBox="0 0 960 540"><g data-node-id="1:2"><text><tspan x="40" y="50">AB</tspan></text></g></svg>',
    ensureRasters: async () => {}, rasterFile: () => fallback,
    ensureOutlinedText: async () => { outlineCalls++; }, outlinedTextFile: () => '', outlinedTextFallbackFile: () => fallback,
    platePath: () => '',
  };
  const plain = await buildDeck(source, ['1:1'], {scale: 2, kern: '100', corr: {}}, {fonts: replacementFonts, images: fakeImages});
  const transformed = await buildDeck(source, ['1:1'], {scale: 2, kern: '100', corr: {}, textTransform: options}, {fonts: replacementFonts, images: fakeImages});
  expect(outlineCalls).toBe(0);
  expect(transformed.pptx).toEqual(plain.pptx);
  expect(transformed.report.frames[0].textTransform).toEqual({outlinedRuns: 0, substitutedRuns: 0});
});

test('a source-font missing glyph becomes a Figma vector while adjacent text stays editable', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'figma2pptx-missing-glyph-test-'));
  const outline = join(dir, 'outlined.svg');
  writeFileSync(outline, '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80" viewBox="0 0 120 80"><path d="M0 20H12V30H0Z M25 20H60V60H25Z" fill="#123456"/></svg>');
  const fallback = join(FIXTURE_CACHE, FIXTURE_KEY, '1001', 'raster', '1-10@2.png');
  const metric = text('1:2', 'Synthetic metric', 20, '\u2243 42', style('Replacement Sans', 'ReplacementSans-Regular', 400));
  const frame = {id: '1:1', name: 'Synthetic missing glyph', type: 'FRAME', visible: true, clipsContent: true,
    absoluteBoundingBox: {x: 0, y: 0, width: 960, height: 540}, fills: [], strokes: [], effects: [], children: [metric]};
  let outlined: string[] = [];
  const source: FrameSource = {
    document: () => frame,
    svg: () => '<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540" viewBox="0 0 960 540"><g data-node-id="1:2"><text><tspan x="40" y="50">\u2243</tspan><tspan x="55" y="50"> 42</tspan></text></g></svg>',
    ensureRasters: async () => {}, rasterFile: () => fallback,
    ensureOutlinedText: async ids => { outlined = ids; }, outlinedTextFile: () => outline, outlinedTextFallbackFile: () => fallback,
    platePath: () => join(dir, 'plate.png'),
  };
  try {
    const fonts = tableFonts([
      {fams: ['Replacement Sans'], styles: ['Regular'], ps: 'ReplacementSans-Regular', weight: 80},
    ], [0x2243]);
    const {pptx, report} = await buildDeck(source, ['1:1'], {scale: 2, kern: '100', corr: {}}, {fonts, images: fakeImages});
    const slide = strFromU8(unzipSync(pptx)['ppt/slides/slide1.xml']);
    expect(outlined).toEqual(['1:2']);
    expect(slide).toContain('<a:t>\u00a0</a:t>');
    expect(slide).toContain('<a:t> 42</a:t>');
    expect(slide).toContain('<a:srcRect l="0" r="87500"/>');
    expect(slide).toContain('<asvg:svgBlip');
    expect(report.missingGlyphs).toEqual([{frame: '1:1', node: '1:2', name: 'Synthetic metric', index: 0, char: '\u2243', codePoint: 'U+2243', sourceFace: 'ReplacementSans-Regular', typeface: 'Replacement Sans', action: 'figma-vector-outline'}]);
    expect(report.textLines[0].unmeasurableReason).toBeUndefined();
    expect(report.textLines[0].measure).toEqual({t: '42', x: 40, ignoreX: true});
  } finally { rmSync(dir, {recursive: true, force: true}); }
});
