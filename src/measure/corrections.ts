// Measure where PowerPoint actually put every text line, against Figma's own baselines, and derive one offset
// per text box for the next build. PowerPoint's first-baseline position depends on font and line spacing in a
// non-linear, quantized way; measuring it is simpler and more exact than modelling it.
import {pdfLines} from './pdflines';
import type {BuildReport, Corrections} from '../convert/build';
import {placeLines, referenceLine, type Placement} from './placement';

export type Measurement = {
  corr: Corrections;
  measuredNodes: string[]; // boxes whose offset was measured (the rest got the median)
  measured: number; // boxes found in the PDF
  unmeasurable: number; // boxes PowerPoint rasterized (semi-transparent text), given the median offset
  medianDy: number;
  spreadMedian: number; // within-box baseline spread after applying one offset per box, px
  spreadMax: number;
  lines: Placement[]; // final observed line offsets; unextractable/rotated lines carry their cause
};

// `pdf` is PowerPoint's export of a build made with `prev` applied; one page per frame, in report order.
export function measure(pdf: Uint8Array, report: Pick<BuildReport, 'frames' | 'textLines'>, prev: Corrections): Measurement {
  const frames = report.frames.map(f => f.id);
  const pages = new Map<number, ReturnType<typeof pdfLines>>();
  const linesOf = (i: number) => { if (!pages.has(i)) pages.set(i, pdfLines(pdf, i, 1 / 0.75)); return pages.get(i)!; };
  const byNode = new Map<string, BuildReport['textLines']>();
  for (const l of report.textLines) { if (!byNode.has(l.node)) byNode.set(l.node, []); byNode.get(l.node)!.push(l); }
  const out: Corrections = {};
  let found = 0, missing = 0; const resid: number[] = [], placements: Placement[] = [];
  const midrange = (v: number[]) => (Math.min(...v) + Math.max(...v)) / 2;
  for (const [node, lines] of byNode) {
    const page = frames.indexOf(lines[0].frame);
    const p0 = prev[node] || {dx: 0, dy: 0};
    const dys: number[] = [], dxs: number[] = [];
    const observed = placeLines(lines.map(referenceLine), linesOf(page));
    placements.push(...observed);
    for (const p of observed) {
      if (p.status !== 'placed') continue;
      // p.dx/p.dy are current output minus Figma; preserve the previous total correction and remove the error.
      dys.push(p0.dy - p.dy!);
      if (p.dx != null) dxs.push(p0.dx - p.dx);
    }
    if (!dys.length) { missing++; continue; }
    found++;
    out[node] = {dx: dxs.length ? +midrange(dxs).toFixed(3) : 0, dy: +midrange(dys).toFixed(3)};
    resid.push(Math.max(...dys) - Math.min(...dys));
  }
  // Boxes PowerPoint rasterized (semi-transparent text) cannot be measured; give them the median offset of measured boxes.
  const med = (v: number[]) => v.sort((a, b) => a - b)[Math.floor(v.length / 2)];
  const mdy = med(Object.values(out).map(o => o.dy)) ?? 0;
  const unmeasured = new Set<string>();
  for (const node of byNode.keys()) if (!out[node]) { out[node] = {dx: 0, dy: mdy}; unmeasured.add(node); }
  return {corr: out, measuredNodes: Object.keys(out).filter(n => !unmeasured.has(n)), measured: found, unmeasurable: missing, medianDy: mdy, spreadMedian: resid.length ? med([...resid]) : 0, spreadMax: resid.length ? Math.max(...resid) : 0, lines: placements};
}
