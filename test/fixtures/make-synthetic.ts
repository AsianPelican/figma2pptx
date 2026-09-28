// Writes the node trees of the synthetic fixture deck (test/fixtures/figma/SyntheticDeck0001): two frames that
// exercise styled runs, soft wraps, centring, bullets, rotation, dashes, holes, gradients, opacity and an image
// fill. The SVG exports next to them are hand-written to match. Run with `bun test/fixtures/make-synthetic.ts`.
import {writeFileSync} from 'node:fs';
const X = 100, Y = 200; // frame 1:2 origin; frame 1:40 sits at (1200, 200)
const box = (x: number, y: number, w: number, h: number, ox = X, oy = Y) => ({x: ox + x, y: oy + y, width: w, height: h});
const solid = (r: number, g: number, b: number, a = 1) => ({type: 'SOLID', visible: true, color: {r, g, b, a}});
const arial = (bold: boolean, size: number, extra: any = {}) => ({fontFamily: 'Arial', fontPostScriptName: bold ? 'Arial-BoldMT' : 'ArialMT', fontWeight: bold ? 700 : 400, fontSize: size, letterSpacing: 0, lineHeightPx: size * 1.2, textAlignHorizontal: 'LEFT', ...extra});
const text = (id: string, name: string, b: any, characters: string, style: any, more: any = {}) => ({id, name, type: 'TEXT', visible: true, absoluteBoundingBox: b, absoluteRenderBounds: b, fills: [solid(0.06, 0.09, 0.13)], strokes: [], effects: [], characters, style, characterStyleOverrides: [], styleOverrideTable: {}, lineTypes: characters.split('\n').map(() => 'NONE'), ...more});
const cover = {
  id: '1:2', name: 'Cover', type: 'FRAME', visible: true, clipsContent: true, absoluteBoundingBox: box(0, 0, 960, 540), absoluteRenderBounds: box(0, 0, 960, 540), fills: [solid(0.957, 0.945, 0.918)], strokes: [], effects: [],
  children: [
    {id: '1:3', name: 'Card', type: 'RECTANGLE', visible: true, absoluteBoundingBox: box(40, 40, 400, 200), cornerRadius: 16, fills: [solid(1, 1, 1)], strokes: [solid(0.118, 0.188, 0.976)], strokeWeight: 2, strokeDashes: [6, 4], effects: []},
    {id: '1:4', name: 'Ring', type: 'VECTOR', visible: true, absoluteBoundingBox: box(780, 50, 120, 120), fills: [solid(0.118, 0.188, 0.976)], strokes: [], effects: []},
    {id: '1:5', name: 'Rule', type: 'LINE', visible: true, absoluteBoundingBox: box(40, 260, 400, 0), fills: [], strokes: [{type: 'GRADIENT_LINEAR', visible: true}], strokeCap: 'ROUND', effects: []},
    {id: '1:6', name: 'Fade', type: 'RECTANGLE', visible: true, absoluteBoundingBox: box(40, 280, 400, 60), fills: [{type: 'GRADIENT_LINEAR', visible: true}], strokes: [], effects: []},
    text('1:7', 'Title', box(40, 360, 440, 96), 'Figma to PowerPoint, one to one', arial(true, 40, {letterSpacing: -0.8, lineHeightPx: 48}), {
      characterStyleOverrides: [...Array(21).fill(0), ...Array(10).fill(1)],
      styleOverrideTable: {1: {fontPostScriptName: 'ArialMT', fontWeight: 400, fills: [solid(0.118, 0.188, 0.976)]}},
    }),
    text('1:8', 'Centred', box(500, 250, 400, 90), 'Centred lines keep their trailing space but give it no advance', arial(false, 20, {textAlignHorizontal: 'CENTER', lineHeightPx: 30})),
    text('1:9', 'List', box(500, 380, 400, 120), 'First point\nSecond point\u2028continues\nPlain line', arial(false, 18, {lineHeightPx: 24}), {lineTypes: ['UNORDERED', 'UNORDERED', 'NONE']}),
    {id: '1:10', name: 'Photo', type: 'RECTANGLE', visible: true, absoluteBoundingBox: box(600, 40, 120, 80), absoluteRenderBounds: box(600, 40, 120, 80), fills: [{type: 'IMAGE', visible: true, scaleMode: 'FILL', imageRef: 'synthetic'}], strokes: [], effects: []},
    {id: '1:11', name: 'Dots', type: 'GROUP', visible: true, opacity: 0.5, absoluteBoundingBox: box(910, 460, 40, 40), fills: [], strokes: [], effects: [], children: [
      {id: '1:12', name: 'Dot', type: 'ELLIPSE', visible: true, absoluteBoundingBox: box(910, 460, 40, 40), fills: [solid(1, 0.4, 0)], strokes: [], effects: []},
    ]},
    {id: '1:13', name: 'Hidden', type: 'RECTANGLE', visible: false, absoluteBoundingBox: box(0, 0, 10, 10), fills: [solid(1, 0, 0)], strokes: [], effects: []},
  ],
};
const DX = 1200;
const details = {
  id: '1:40', name: 'Details', type: 'FRAME', visible: true, clipsContent: true, absoluteBoundingBox: box(0, 0, 960, 540, DX), absoluteRenderBounds: box(0, 0, 960, 540, DX), fills: [solid(1, 1, 1)], strokes: [], effects: [],
  children: [
    text('1:41', 'Side label', box(40, 120, 24, 300, DX), 'ROTATED LABEL', arial(true, 20, {letterSpacing: 2}), {rotation: 1.5707963}),
    text('1:42', 'Note', box(500, 100, 400, 28, DX), 'Right aligned, underlined', arial(false, 22, {textAlignHorizontal: 'RIGHT', textDecoration: 'UNDERLINE'}), {fills: [solid(0.118, 0.188, 0.976, 1)], opacity: 1}),
    {id: '1:43', name: 'Bracket', type: 'VECTOR', visible: true, absoluteBoundingBox: box(500, 200, 200, 120, DX), fills: [], strokes: [solid(0.06, 0.09, 0.13)], effects: []},
  ],
};
for (const f of [cover, details]) writeFileSync(`${import.meta.dir}/figma/SyntheticDeck0001/1001/nodes/${f.id.replace(':', '-')}.json`, JSON.stringify(f));
