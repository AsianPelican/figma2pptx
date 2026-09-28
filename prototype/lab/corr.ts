// Measure where PowerPoint actually put every text line, against Figma's own baselines, and emit per-text-box
// offsets for the next build.  bun corr.ts <deck.pdf> <deck.report.json> <corr-in.json|-> <corr-out.json>
import {readFileSync, writeFileSync, existsSync} from 'node:fs';
import {pdfLines} from './pdflines';
const [pdf, rep, cin, cout] = process.argv.slice(2);
const R = JSON.parse(readFileSync(rep, 'utf8'));
const prev: Record<string, {dx: number, dy: number}> = cin !== '-' && existsSync(cin) ? JSON.parse(readFileSync(cin, 'utf8')) : {};
const frames: string[] = R.frames.map((f: any) => f.id);
const pages = new Map<number, any[]>();
const charsOf = (i: number) => { if (!pages.has(i)) pages.set(i, pdfLines(readFileSync(pdf), i, 1 / 0.75).flatMap(l => l.chars.map(c => ({...c, base: l.base})))); return pages.get(i)!; };
const byNode = new Map<string, any[]>();
for (const l of R.textLines) { if (!byNode.has(l.node)) byNode.set(l.node, []); byNode.get(l.node)!.push(l); }
const out: Record<string, {dx: number, dy: number}> = {};
let found = 0, missing = 0; const resid: number[] = [];
const midrange = (v: number[]) => (Math.min(...v) + Math.max(...v)) / 2;
for (const [node, lines] of byNode) {
  const page = frames.indexOf(lines[0].frame), C = charsOf(page);
  const p0 = prev[node] || {dx: 0, dy: 0};
  const dys: number[] = [], dxs: number[] = [];
  for (const l of lines) {
    const t = l.t.replace(/\s+$/, ''); const k = t.search(/\S/); if (k < 0) continue;
    const want = t.slice(k).replace(/\s/g, '').slice(0, 6);
    // expected position of this line in the current build = Figma position + the offset already applied
    const ex = l.x + p0.dx, ey = l.base + p0.dy;
    const cands = C.filter(c => c.c === want[0] && Math.abs(c.base - ey) < 14 && Math.abs(c.x - ex) < 60);
    let best: any = null, bd = 1e9;
    for (const c of cands) {
      const i = C.indexOf(c); let s = ''; for (let j = i; j < C.length && s.length < want.length; j++) if (C[j].c.trim()) s += C[j].c;
      if (s !== want) continue;
      const d = Math.hypot(c.x - ex, c.base - ey); if (d < bd) { bd = d; best = c; }
    }
    if (!best) continue;
    // Figma's line x is the pen position of the first character, including leading spaces; PDF gives the first glyph
    dys.push(l.base - (best.base - p0.dy));
    if (k === 0) dxs.push(l.x - (best.x - p0.dx));
  }
  if (!dys.length) { missing++; continue; }
  found++;
  out[node] = {dx: dxs.length ? +midrange(dxs).toFixed(3) : 0, dy: +midrange(dys).toFixed(3)};
  resid.push(Math.max(...dys) - Math.min(...dys));
}
// Boxes PowerPoint rasterized (semi-transparent text) cannot be measured; give them the median offset of measured boxes.
const med = (v: number[]) => v.sort((a, b) => a - b)[Math.floor(v.length / 2)];
const mdy = med(Object.values(out).map(o => o.dy));
for (const node of byNode.keys()) if (!out[node]) out[node] = {dx: 0, dy: mdy};
writeFileSync(cout, JSON.stringify(out, null, 1));
console.log(`measured ${found} boxes, ${missing} unmeasurable (median dy ${mdy.toFixed(2)} px applied); within-box baseline spread: median ${med([...resid]).toFixed(2)} px, max ${Math.max(...resid).toFixed(2)} px`);
