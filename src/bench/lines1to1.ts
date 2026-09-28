// Per-line typography comparison: Figma's PDF of a frame (reference) against PowerPoint's PDF of the slide.
// A line passes when its text is equal, its left x and baseline are within 0.75 pt, and its width within 1.5 pt.
import * as m from 'mupdf';
import {readFileSync} from 'node:fs';

export type Line = {t: string, x0: number, x1: number, base: number, size: number, top: number, bot: number};

export function lines(pdf: Uint8Array, page: number, scale: number): Line[] {
  const d = m.Document.openDocument(Buffer.from(pdf), 'application/pdf'); const p = d.loadPage(page); const out: Line[] = []; let cur: any;
  p.toStructuredText('preserve-whitespace').walk({
    beginLine() { cur = {t: '', q: [] as number[][], o: null, s: 0}; },
    onChar(c: string, o: number[], f: any, s: number, q: number[]) { cur.t += c; if (c.trim()) { cur.q.push(q); if (!cur.o) cur.o = o; cur.s = Math.max(cur.s, s); } },
    endLine() {
      if (!cur.q.length) return;
      const xs = cur.q.flatMap((q: number[]) => [q[0], q[2], q[4], q[6]]), ys = cur.q.flatMap((q: number[]) => [q[1], q[3], q[5], q[7]]);
      out.push({t: cur.t, x0: Math.min(...xs) * scale, x1: Math.max(...xs) * scale, base: cur.o[1] * scale, size: cur.s * scale, top: Math.min(...ys) * scale, bot: Math.max(...ys) * scale});
    },
  } as any);
  return out;
}
export const norm = (s: string) => s.normalize('NFKC').replace(/ˇt/g, "'").replace(/[ˇ’‘`´]/g, "'").replace(/[\u2028\s]+/g, ' ').trim().toLowerCase();
export const key = (s: string) => norm(s).replace(/[^a-z0-9%$.,&]/g, '');
export function dice(a: string, b: string) {
  const g = (s: string) => { const mm = new Map<string, number>(); for (let i = 0; i < s.length - 1; i++) { const k = s.slice(i, i + 2); mm.set(k, (mm.get(k) || 0) + 1); } return mm; };
  const A = g(a), B = g(b); let o = 0, na = 0, nb = 0;
  for (const v of A.values()) na += v; for (const v of B.values()) nb += v; for (const [k, v] of A) o += Math.min(v, B.get(k) || 0);
  return na + nb ? 2 * o / (na + nb) : (a === b ? 1 : 0);
}
// Figma PDFs can carry a hidden duplicate text layer; keep the larger (visible) instance of a repeated line.
export function dedupe(F: Line[]) { return F.filter(a => !F.some(b => b !== a && key(b.t) === key(a.t) && key(a.t).length > 0 && Math.abs(b.x0 - a.x0) < 40 && Math.abs(b.base - a.base) < 40 && b.size > a.size + 0.5)); }

// `figmaPdf` is Figma's single-page PDF of the frame (1 px = 1 pt); `slide` is 1-based in the PowerPoint PDF.
export function compare(figmaPdf: string, pptPdf: string, slide: number) {
  const F = dedupe(lines(readFileSync(figmaPdf), 0, 0.75)), P = lines(readFileSync(pptPdf), slide - 1, 1);
  const used = new Set<number>(); const rows: any[] = [];
  for (const f of F) {
    let best = -1, bs = 0;
    P.forEach((p, i) => { if (used.has(i)) return; const s = dice(key(f.t), key(p.t)) - Math.min(0.3, Math.hypot(p.x0 - f.x0, p.base - f.base) / 600); if (s > bs) { bs = s; best = i; } });
    if (best >= 0 && bs > 0.75) { used.add(best); const p = P[best]; rows.push({t: f.t.trim().slice(0, 48), exact: key(f.t) === key(p.t), dx: +(p.x0 - f.x0).toFixed(1), dBase: +(p.base - f.base).toFixed(1), dW: +((p.x1 - p.x0) - (f.x1 - f.x0)).toFixed(1), size: +f.size.toFixed(1), pSize: +p.size.toFixed(1), ppt: key(f.t) === key(p.t) ? undefined : p.t.trim().slice(0, 48)}); }
    else rows.push({t: f.t.trim().slice(0, 48), missing: true});
  }
  const extra = P.filter((_, i) => !used.has(i)).map(p => p.t.trim().slice(0, 48));
  return {slide, figLines: F.length, pptLines: P.length, rows, extra};
}
