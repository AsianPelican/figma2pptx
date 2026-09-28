#!/usr/bin/env bun
// Fidelity benchmark of a PowerPoint-exported PDF against Figma. A separate command: it measures, it never
// changes the conversion.
//
//   bun run bench <deck.pdf> [--report <deck.report.json>] [--ref <dir of NNN.pdf>] [--label name] [--diffs dir]
//
// The deck's report (written next to it by figma2pptx) names the Figma file, the frames and the cache holding
// their SVG exports. The PDF should be PowerPoint's raw export (--pdf-preset raw) for numbers comparable with
// earlier runs; an optimized PDF differs only in photo pixels.
//
// Metrics:
//   linesWithin075pt   Figma's PDF lines against PowerPoint's (lines1to1): text equal, left x and baseline within
//                      0.75 pt, width within 1.5 pt. Reference: Figma's own PDF export of each frame (fetched and
//                      cached, or --ref <dir> with 000.pdf, 001.pdf, ...).
//   lineBreaksExact    every laid-out Figma line (one SVG tspan row) appears as one PDF line with the same text.
//   linesPlaced1px     that line's first glyph is within 1 px of Figma's baseline and x.
//   meanDiffPct/pxOffPct  both rendered at 1440 px wide: mean absolute difference, and the share of pixels off
//                      by more than 32/255.
import {execFileSync} from 'node:child_process';
import {readFileSync, writeFileSync, mkdirSync, rmSync, existsSync} from 'node:fs';
import {join, resolve, basename} from 'node:path';
import {tmpdir} from 'node:os';
import {DOMParser} from '@xmldom/xmldom';
import {pdfLines, pageCount} from '../measure/pdflines';
import {compare} from './lines1to1';
import {placeLines, normalizeLine, type ReferenceLine} from '../measure/placement';
import {FigmaClient, readToken} from '../figma/api';
import {FigmaFile} from '../figma/source';

const args = process.argv.slice(2);
const opt = (k: string) => { const i = args.indexOf(k); return i >= 0 ? args[i + 1] : undefined; };
const pdf = args[0] && !args[0].startsWith('--') ? resolve(args[0]) : '';
if (!pdf || !existsSync(pdf)) { console.error('usage: bun run bench <deck.pdf> [--report deck.report.json] [--ref dir] [--label name] [--diffs dir]'); process.exit(2); }
const reportPath = opt('--report') ?? pdf.replace(/(_(raw-export|screen|standard|print))?\.pdf$/i, '') + '.report.json';
if (!existsSync(reportPath)) { console.error(`no conversion report at ${reportPath} (pass --report)`); process.exit(2); }
const R = JSON.parse(readFileSync(reportPath, 'utf8'));
const FRAMES: string[] = R.frameIds;
const label = opt('--label') ?? basename(pdf, '.pdf');
const DIFFS = opt('--diffs');
const t0 = Date.now();

// Figma's PDFs of the frames: a given directory, or fetched once into the cache.
let refs: string[];
const refDir = opt('--ref');
if (refDir) refs = FRAMES.map((_, i) => join(resolve(refDir), String(i).padStart(3, '0') + '.pdf'));
else refs = await (await FigmaFile.open(R.fileKey, resolve(R.cache, '../..'), new FigmaClient(readToken()))).framePdfs(FRAMES);
if (pageCount(readFileSync(pdf)) !== FRAMES.length) throw Error(`${pdf} has ${pageCount(readFileSync(pdf))} pages, the report lists ${FRAMES.length} frames`);

const tmp = join(tmpdir(), `figma2pptx-bench-${process.pid}`); mkdirSync(tmp, {recursive: true});
const ppm = (f: string) => { const d = readFileSync(f); let o = 0; const fields: number[] = []; while (fields.length < 3) { while (d[o] === 35 || d[o] <= 32) { if (d[o] === 35) while (d[o] !== 10) o++; o++; } let s = ''; while (d[o] > 32) s += String.fromCharCode(d[o++]); if (s !== 'P6') fields.push(+s); } o++; return {w: fields[0], h: fields[1], px: d.subarray(o)}; };
const rows: any[] = [];
let totA = 0, okA = 0, totB = 0, okBtext = 0, okBpos = 0, legacyTotal = 0, legacyWithin1px = 0, unmeasurable = 0, worstOffsetPx = 0;
const placementFailures: any[] = [];
try {
  execFileSync('pdftoppm', ['-scale-to-x', '1440', '-scale-to-y', '810', pdf, `${tmp}/c`], {stdio: 'ignore'});
  for (let i = 0; i < FRAMES.length; i++) {
    const nn = String(i).padStart(3, '0');
    execFileSync('pdftoppm', ['-scale-to-x', '1440', '-scale-to-y', '810', '-singlefile', refs[i], `${tmp}/r${nn}`], {stdio: 'ignore'});
    const A = ppm(`${tmp}/r${nn}.ppm`), cFile = `${tmp}/c-${String(i + 1).padStart(2, '0')}.ppm`;
    const B = ppm(existsSync(cFile) ? cFile : `${tmp}/c-${i + 1}.ppm`);
    let sum = 0, off = 0; const heat = Buffer.alloc(A.w * A.h * 3);
    for (let k = 0; k < A.w * A.h; k++) {
      const d0 = Math.abs(A.px[k * 3] - B.px[k * 3]), d1 = Math.abs(A.px[k * 3 + 1] - B.px[k * 3 + 1]), d2 = Math.abs(A.px[k * 3 + 2] - B.px[k * 3 + 2]);
      sum += d0 + d1 + d2; const mx = Math.max(d0, d1, d2); if (mx > 32) off++;
      const g = Math.round((A.px[k * 3] + A.px[k * 3 + 1] + A.px[k * 3 + 2]) / 3 * 0.25 + 180);
      heat[k * 3] = mx > 32 ? 255 : g; heat[k * 3 + 1] = mx > 32 ? 0 : g; heat[k * 3 + 2] = mx > 32 ? 60 : g;
    }
    if (DIFFS) { mkdirSync(DIFFS, {recursive: true}); writeFileSync(`${tmp}/h.ppm`, Buffer.concat([Buffer.from(`P6\n${A.w} ${A.h}\n255\n`), heat])); execFileSync('magick', [`${tmp}/h.ppm`, '-resize', '960x', `${DIFFS}/${label}-${nn}-diff.png`]); }
    // (a) lines against Figma's PDF
    const ca = compare(refs[i], pdf, i + 1) as any;
    const bad = ca.rows.filter((x: any) => x.missing || !x.exact || Math.abs(x.dx) > 0.75 || Math.abs(x.dBase) > 0.75 || Math.abs(x.dW) > 1.5);
    totA += ca.rows.length; okA += ca.rows.length - bad.length;
    // (b) Figma's laid-out lines from the SVG export
    const svg = new DOMParser().parseFromString(readFileSync(join(R.cache, 'svg', FRAMES[i].replace(/:/g, '-').replace(/;/g, '_') + '.svg'), 'utf8'), 'image/svg+xml') as any;
    const raw: (ReferenceLine & {el: any})[] = [];
    for (const t of Array.from(svg.getElementsByTagName('text')) as any[]) {
      for (const sp of Array.from(t.getElementsByTagName('tspan')) as any[]) {
        const x = +sp.getAttribute('x'), y = +sp.getAttribute('y'), s = sp.textContent || '';
        const owner = t.getAttribute('data-node-id') ? t : t.parentNode;
        const l = raw.find(l => Math.abs(l.y - y) < 0.01 && Math.abs(l.x - x) < 2000 && l.el === owner);
        if (l) { if (x < l.x) { l.t = s + l.t; l.x = x; } else l.t += s; }
        else raw.push({t: s, x, y, node: owner?.getAttribute?.('data-node-id') || undefined, el: owner});
      }
    }
    let figLines: ReferenceLine[];
    if (Array.isArray(R.textLines)) {
      // The conversion report records the transformed slide-space coordinates used to build the deck. Prefer
      // those to raw SVG x/y, which omit a parent transform on some Figma exports.
      figLines = R.textLines.filter((l: any) => l.frame === FRAMES[i]).map((l: any) => ({t: l.t, x: l.x, y: l.base, node: l.node, size: l.size, ignoreX: l.ignoreX ?? /^\s/.test(l.t), rotated: l.rotated}));
    } else figLines = raw;
    const P = pdfLines(readFileSync(pdf), i, 1 / 0.75);
    // Keep the prototype's original 319/334 metric byte-for-byte comparable. The richer placement result below
    // fixes its known false misses (transforms, leading spaces and merged PDF lines) and is the acceptance gate.
    const pchars = P.flatMap(l => l.chars.map(c => ({...c, base: l.base})));
    const legacy = raw.filter(l => normalizeLine(l.t));
    const legacyOk = legacy.filter(l => {
      const first = normalizeLine(l.t).replace(/\s/g, '')[0];
      return pchars.some(c => c.c === first && Math.abs(c.base - l.y) <= 1 && Math.abs(c.x - l.x) <= 1);
    }).length;
    legacyTotal += legacy.length; legacyWithin1px += legacyOk;
    const fl = figLines.filter(l => normalizeLine(l.t));
    const placed = placeLines(fl, P);
    const bt = placed.filter(x => x.exactLineBreak).length;
    const bp = placed.filter(x => x.status === 'placed' && x.offsetPx! <= 1).length;
    const missing = placed.filter(x => x.status === 'unmeasurable');
    const over = placed.filter(x => x.status === 'placed' && x.offsetPx! > 1);
    unmeasurable += missing.length;
    worstOffsetPx = Math.max(worstOffsetPx, ...placed.filter(x => x.status === 'placed').map(x => x.offsetPx!), 0);
    placementFailures.push(...[...over, ...missing].map(x => ({slide: i + 1, node: x.node, text: normalizeLine(x.t).slice(0, 80), dx: x.dx, dy: x.dy, offsetPx: x.offsetPx, cause: x.cause ?? `${x.offsetPx} px exceeds 1 px`})));
    totB += fl.length; okBtext += bt; okBpos += bp;
    rows.push({slide: i + 1, meanDiffPct: +(sum / (A.w * A.h * 3) / 255 * 100).toFixed(2), pxOffPct: +(off / (A.w * A.h) * 100).toFixed(2), linesA: `${ca.rows.length - bad.length}/${ca.rows.length}`, breaks: `${bt}/${fl.length}`, pos1px: `${bp}/${fl.length}`, legacyPos1px: `${legacyOk}/${legacy.length}`, worstOffsetPx: +Math.max(0, ...placed.filter(x => x.status === 'placed').map(x => x.offsetPx!)).toFixed(2), unmeasurable: missing.length, failures: [...over, ...missing].map(x => ({node: x.node, text: normalizeLine(x.t).slice(0, 80), dx: x.dx, dy: x.dy, offsetPx: x.offsetPx, cause: x.cause}))});
  }
} finally { rmSync(tmp, {recursive: true, force: true}); }
const summary = {label, pdf, pdfMB: +(readFileSync(pdf).length / 1e6).toFixed(2), meanDiffPct: +(rows.reduce((a, r) => a + r.meanDiffPct, 0) / rows.length).toFixed(2), pxOffPct: +(rows.reduce((a, r) => a + r.pxOffPct, 0) / rows.length).toFixed(2), linesWithin075pt: `${okA}/${totA}`, lineBreaksExact: `${okBtext}/${totB}`, linesPlaced1px: `${okBpos}/${totB}`, legacyLinesPlaced1px: `${legacyWithin1px}/${legacyTotal}`, worstOffsetPx: +worstOffsetPx.toFixed(2), unmeasurableLines: unmeasurable, failedPlacements: placementFailures.length, seconds: +((Date.now() - t0) / 1000).toFixed(1)};
console.log(JSON.stringify(summary));
for (const r of rows) console.log(JSON.stringify(r));
for (const f of placementFailures) console.error(JSON.stringify(f));
const outDir = resolve(import.meta.dir, '../../bench-results'); mkdirSync(outDir, {recursive: true});
writeFileSync(join(outDir, `${label}.json`), JSON.stringify({summary, rows, placementFailures}, null, 1));
console.log(`written ${join(outDir, label + '.json')}`);
