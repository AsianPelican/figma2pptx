// Fidelity benchmark of a PowerPoint-exported PDF against Figma.
//   bun bench.ts <candidate.pdf> <label> [--ref cache/figpdf] [--diffs dir]
// Pixel: both rendered at 1440x810; mean absolute difference and share of pixels off by more than 32/255.
// Lines (a): the prior worker's lines1to1 comparator (Figma PDF lines vs PowerPoint PDF lines; "ok" = text equal and
//   left x, baseline within 0.75 pt, width within 1.5 pt), so numbers are comparable with the Framedeck baseline.
// Lines (b): line-break fidelity against Figma's own layout: every laid-out line from the Figma SVG export
//   (one tspan row per line) must appear as one line in the candidate, same text, baseline and left edge within 1 px.
import {execFileSync} from 'node:child_process';
import {readFileSync, writeFileSync, mkdirSync, rmSync, existsSync} from 'node:fs';
import {DOMParser} from '@xmldom/xmldom';
import {pdfLines} from './pdflines';
import {compare} from '../tooling/lines1to1.ts';

const args = process.argv.slice(2);
const opt = (k: string, d: string) => { const i = args.indexOf(k); return i >= 0 ? args[i + 1] : d; };
const [pdf, label] = args;
const REF = opt('--ref', 'cache/figpdf'), DIFFS = opt('--diffs', '');
const FRAMES = opt('--frames', '').split(',');
const tmp = `tmp-bench-${Date.now()}`; mkdirSync(tmp);
const ppm = (f: string) => { const d = readFileSync(f); let o = 0; const fields: number[] = []; while (fields.length < 3) { while (d[o] === 35 || d[o] <= 32) { if (d[o] === 35) while (d[o] !== 10) o++; o++; } let s = ''; while (d[o] > 32) s += String.fromCharCode(d[o++]); if (s !== 'P6') fields.push(+s); } o++; return {w: fields[0], h: fields[1], px: d.subarray(o)}; };
execFileSync('pdftoppm', ['-scale-to-x', '1440', '-scale-to-y', '810', pdf, `${tmp}/c`], {stdio: 'ignore'});
const norm = (s: string) => s.normalize('NFKC').replace(/[’‘`´]/g, "'").replace(/[\u2028\s]+/g, ' ').trim();
const rows: any[] = [];
let totA = 0, okA = 0, totB = 0, okBtext = 0, okBpos = 0;
for (let i = 0; i < FRAMES.length; i++) {
  const nn = String(i).padStart(3, '0');
  execFileSync('pdftoppm', ['-scale-to-x', '1440', '-scale-to-y', '810', '-singlefile', `${REF}/${nn}.pdf`, `${tmp}/r${nn}`], {stdio: 'ignore'});
  const A = ppm(`${tmp}/r${nn}.ppm`), cFile = readFileSync.length && `${tmp}/c-${String(i + 1).padStart(2, '0')}.ppm`;
  const B = ppm(existsSync(cFile) ? cFile : `${tmp}/c-${i + 1}.ppm`);
  let sum = 0, off = 0; const heat = Buffer.alloc(A.w * A.h * 3);
  for (let k = 0; k < A.w * A.h; k++) {
    const d0 = Math.abs(A.px[k * 3] - B.px[k * 3]), d1 = Math.abs(A.px[k * 3 + 1] - B.px[k * 3 + 1]), d2 = Math.abs(A.px[k * 3 + 2] - B.px[k * 3 + 2]);
    sum += d0 + d1 + d2; const mx = Math.max(d0, d1, d2); if (mx > 32) off++;
    const g = Math.round((A.px[k * 3] + A.px[k * 3 + 1] + A.px[k * 3 + 2]) / 3 * 0.25 + 180);
    heat[k * 3] = mx > 32 ? 255 : g; heat[k * 3 + 1] = mx > 32 ? 0 : g; heat[k * 3 + 2] = mx > 32 ? 60 : g;
  }
  if (DIFFS) { mkdirSync(DIFFS, {recursive: true}); writeFileSync(`${tmp}/h.ppm`, Buffer.concat([Buffer.from(`P6\n${A.w} ${A.h}\n255\n`), heat])); execFileSync('magick', [`${tmp}/h.ppm`, '-resize', '960x', `${DIFFS}/${label}-${nn}-diff.png`]); }
  // (a)
  const ca = compare(REF, pdf, i + 1) as any;
  const bad = ca.rows.filter((x: any) => x.missing || !x.exact || Math.abs(x.dx) > 0.75 || Math.abs(x.dBase) > 0.75 || Math.abs(x.dW) > 1.5);
  totA += ca.rows.length; okA += ca.rows.length - bad.length;
  // (b)
  const svg = new DOMParser().parseFromString(readFileSync(`cache/svg/${FRAMES[i].replace(':', '-')}.svg`, 'utf8'), 'image/svg+xml') as any;
  const figLines: {t: string, x: number, y: number}[] = [];
  for (const t of Array.from(svg.getElementsByTagName('text')) as any[]) {
    for (const sp of Array.from(t.getElementsByTagName('tspan')) as any[]) {
      const x = +sp.getAttribute('x'), y = +sp.getAttribute('y'), s = sp.textContent || '';
      const l = figLines.find(l => Math.abs(l.y - y) < 0.01 && Math.abs(l.x - x) < 2000 && l.t !== undefined && (l as any).el === (t.getAttribute("data-node-id") ? t : t.parentNode));
      if (l) { if (x < l.x) { l.t = s + l.t; l.x = x; } else l.t += s; } else figLines.push({t: s, x, y, el: t.getAttribute("data-node-id") ? t : t.parentNode} as any);
    }
  }
  const P = pdfLines(readFileSync(pdf), i, 1 / 0.75);
  const pchars = P.flatMap(l => l.chars.map(c => ({...c, base: l.base})));
  let bt = 0, bp = 0; const misses: string[] = [];
  const fl = figLines.filter(l => norm(l.t));
  for (const l of fl) {
    const want = norm(l.t);
    // candidate line = the extracted PDF line nearest to Figma's baseline whose text contains this line's text as a whole line
    const cands = P.filter(p => Math.abs(p.base - l.y) < 12 && Math.abs(p.x0 - l.x) < 400);
    const hit = cands.find(p => norm(p.t) === want) || cands.find(p => norm(p.t).includes(want) && norm(p.t).startsWith(want.split(' ')[0]));
    if (hit && norm(hit.t) === want) bt++; else misses.push(want.slice(0, 40));
    // position: first glyph of the line
    const first = want.replace(/\s/g, '')[0];
    const g = pchars.filter(c => c.c === first && Math.abs(c.base - l.y) < 3 && Math.abs(c.x - l.x) < 3).sort((a, b) => Math.hypot(a.x - l.x, a.base - l.y) - Math.hypot(b.x - l.x, b.base - l.y))[0];
    if (g && Math.abs(g.base - l.y) <= 1 && Math.abs(g.x - l.x) <= 1) bp++;
  }
  totB += fl.length; okBtext += bt; okBpos += bp;
  rows.push({slide: i + 1, meanDiffPct: +(sum / (A.w * A.h * 3) / 255 * 100).toFixed(2), pxOffPct: +(off / (A.w * A.h) * 100).toFixed(2), linesA: `${ca.rows.length - bad.length}/${ca.rows.length}`, breaks: `${bt}/${fl.length}`, pos1px: `${bp}/${fl.length}`, misses: misses.slice(0, 4)});
}
rmSync(tmp, {recursive: true, force: true});
const summary = {label, pdf, pdfMB: +(readFileSync(pdf).length / 1e6).toFixed(2), meanDiffPct: +(rows.reduce((a, r) => a + r.meanDiffPct, 0) / rows.length).toFixed(2), pxOffPct: +(rows.reduce((a, r) => a + r.pxOffPct, 0) / rows.length).toFixed(2), linesWithin075pt: `${okA}/${totA}`, lineBreaksExact: `${okBtext}/${totB}`, linesPlaced1px: `${okBpos}/${totB}`};
console.log(JSON.stringify(summary));
for (const r of rows) console.log(JSON.stringify(r));
mkdirSync('bench', {recursive: true});
writeFileSync(`bench/${label}.json`, JSON.stringify({summary, rows}, null, 1));
