import type {PdfLine} from './pdflines';
import type {LineRecord} from '../convert/text';

// `match`: the text to find in PowerPoint's PDF when it differs from `t` (glyphs drawn as vectors are not text there).
export type ReferenceLine = {t: string, match?: string, x: number, y: number, node?: string, size?: number, ignoreX?: boolean, rotated?: boolean, unmeasurableReason?: string};
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

// A conversion report's line as a placement reference: its measurable native text, pen x and baseline.
export function referenceLine(l: LineRecord): ReferenceLine {
  const m = l.measure;
  return {t: l.t, ...(m ? {match: m.t} : {}), x: m?.x ?? l.x, y: l.base, node: l.node, size: l.size, ignoreX: m ? !!m.ignoreX : (l.ignoreX ?? /^\s/.test(l.t)), rotated: l.rotated, unmeasurableReason: l.unmeasurableReason};
}

export function placeLines(reference: ReferenceLine[], pdf: PdfLine[]): Placement[] {
  return reference.map(l => {
    if (l.rotated) return {...l, status: 'unmeasurable' as const, exactLineBreak: false, cause: 'rotated text uses a local coordinate system'};
    const want = normalizeLine(l.match ?? l.t);
    let best: {line: PdfLine, glyph: NonNullable<ReturnType<typeof firstGlyph>>, score: number} | undefined;
    for (const p of pdf) {
      if (!l.ignoreX && Math.abs(p.x0 - l.x) > 450) continue;
      const glyph = firstGlyph(p, want);
      if (!glyph) continue;
      const dx = glyph.x - l.x, dy = glyph.y - l.y;
      if (Math.abs(dy) > 20) continue;
      const distance = Math.hypot(l.ignoreX ? 0 : dx, dy);
      if (distance > 80) continue; // repeated copy elsewhere on the slide is not this line
      const sizePenalty = l.size == null ? 0 : Math.abs(p.size - l.size) * 20;
      const exact = normalizeLine(p.t) === want;
      const score = (exact ? 0 : 1000) + distance + sizePenalty;
      if (!best || score < best.score) best = {line: p, glyph, score};
    }
    // Some display fonts export an altered/incomplete PDF text string even though the native text line and its
    // glyph geometry remain available. When one nearby line has the same size, use that geometry rather than
    // calling the line unmeasurable. Known rasterized lines never take this fallback.
    if (!best && !l.unmeasurableReason && l.size != null) {
      const size = l.size;
      const near = pdf.flatMap(line => {
        const glyph = line.chars.find(c => c.c.trim());
        if (!glyph || Math.abs(glyph.y - l.y) > 30 || (!l.ignoreX && Math.abs(glyph.x - l.x) > 30)) return [];
        if (Math.abs(line.size - size) > Math.max(0.75, size * 0.01)) return [];
        return [{line, glyph}];
      });
      if (near.length === 1) best = {...near[0], score: 2000};
    }
    if (!best) return {
      ...l,
      status: 'unmeasurable' as const,
      exactLineBreak: false,
      cause: l.unmeasurableReason ?? 'text is not extractable from PowerPoint PDF',
    };
    const dx = l.ignoreX ? undefined : +(best.glyph.x - l.x).toFixed(2), dy = +(best.glyph.y - l.y).toFixed(2);
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
