import type {PdfLine} from './pdflines';

export type ReferenceLine = {t: string, x: number, y: number, node?: string, size?: number, ignoreX?: boolean, rotated?: boolean};
export type Placement = ReferenceLine & {
  status: 'placed' | 'unmeasurable';
  dx?: number;
  dy?: number;
  offsetPx?: number;
  exactLineBreak: boolean;
  cause?: string;
};

export const normalizeLine = (s: string) => s.normalize('NFKC')
  .replace(/[’‘`´]/g, "'")
  .replace(/[\u2028\s]+/g, ' ')
  .trim();

const charKey = (s: string) => normalizeLine(s).toLowerCase().replace(/\s/g, '');

// Find the first glyph of `needle` inside one extracted PDF line while preserving the glyph coordinate. This
// also handles several adjacent Figma text nodes that PowerPoint combines into one extracted PDF line.
function firstGlyph(line: PdfLine, needle: string) {
  const chars = line.chars.flatMap(c => [...charKey(c.c)].map(k => ({k, c})));
  const hay = chars.map(x => x.k).join(''), want = charKey(needle);
  const at = hay.indexOf(want);
  return at >= 0 ? chars[at]?.c : undefined;
}

export function placeLines(reference: ReferenceLine[], pdf: PdfLine[]): Placement[] {
  return reference.map(l => {
    if (l.rotated) return {...l, status: 'unmeasurable' as const, exactLineBreak: false, cause: 'rotated text uses a local coordinate system'};
    const want = normalizeLine(l.t);
    let best: {line: PdfLine, glyph: NonNullable<ReturnType<typeof firstGlyph>>, score: number} | undefined;
    for (const p of pdf) {
      if (Math.abs(p.base - l.y) > 20 || (!l.ignoreX && Math.abs(p.x0 - l.x) > 450)) continue;
      const glyph = firstGlyph(p, want);
      if (!glyph) continue;
      const dx = glyph.x - l.x, dy = p.base - l.y;
      const distance = Math.hypot(l.ignoreX ? 0 : dx, dy);
      if (distance > 80) continue; // repeated copy elsewhere on the slide is not this line
      const sizePenalty = l.size == null ? 0 : Math.abs(p.size - l.size) * 20;
      const exact = normalizeLine(p.t) === want;
      const score = (exact ? 0 : 1000) + distance + sizePenalty;
      if (!best || score < best.score) best = {line: p, glyph, score};
    }
    if (!best) return {
      ...l,
      status: 'unmeasurable' as const,
      exactLineBreak: false,
      cause: 'text is not extractable from PowerPoint PDF (commonly transparent text rasterized by PowerPoint)',
    };
    const dx = l.ignoreX ? undefined : +(best.glyph.x - l.x).toFixed(2), dy = +(best.line.base - l.y).toFixed(2);
    return {
      ...l,
      status: 'placed' as const,
      dx,
      dy,
      offsetPx: +Math.max(dx == null ? 0 : Math.abs(dx), Math.abs(dy)).toFixed(2),
      exactLineBreak: normalizeLine(best.line.t) === want,
    };
  });
}
