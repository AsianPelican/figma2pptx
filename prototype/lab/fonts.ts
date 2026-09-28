// Map a Figma face (PostScript name, or family + weight + italic) to the name PowerPoint selects it by:
// the legacy family (OpenType name ID 1) plus bold/italic flags from the legacy subfamily (name ID 2).
import {execFileSync} from 'node:child_process';
export type Face = {typeface: string, b: 0 | 1, i: 0 | 1, ps: string, installed: boolean, note?: string};
type Row = {fams: string[], styles: string[], ps: string, weight: number};
let rows: Row[] | null = null;
function load() {
  if (rows) return rows;
  const out = execFileSync('fc-list', ['--format', '%{family}|%{style}|%{postscriptname}|%{weight}\n'], {encoding: 'utf8', maxBuffer: 64e6});
  rows = out.split('\n').filter(Boolean).map(l => { const [f, s, ps, w] = l.split('|'); return {fams: f.split(','), styles: s.split(','), ps, weight: +w}; });
  return rows;
}
const cssWeightToFc: Record<number, number> = {100: 0, 200: 40, 300: 50, 400: 80, 500: 100, 600: 180, 700: 200, 800: 205, 900: 210};
export function mapFace(ps: string | null | undefined, family: string, weight: number, italic: boolean): Face {
  const R = load();
  let r = ps ? R.find(x => x.ps === ps) : undefined;
  let note: string | undefined;
  if (!r) {
    // No PostScript name (Figma's bundled Inter) or not installed: pick the installed static face of the family nearest in weight.
    const want = cssWeightToFc[Math.round(weight / 100) * 100] ?? 80;
    const cands = R.filter(x => x.fams.includes(family) && /Italic/i.test(x.styles.join()) === italic && !x.ps.includes('Variable'));
    cands.sort((a, b) => Math.abs(a.weight - want) - Math.abs(b.weight - want));
    r = cands[0];
    if (r && r.weight !== want) note = `${family} ${weight}${italic ? ' italic' : ''} not installed; nearest installed face ${r.ps}`;
  }
  if (!r) return {typeface: family, b: weight >= 700 ? 1 : 0, i: italic ? 1 : 0, ps: ps || '', installed: false, note: `${ps || family} not installed`};
  const fam = r.fams[r.fams.length - 1], sty = r.styles[r.styles.length - 1];
  return {typeface: fam, b: /Bold/i.test(sty) ? 1 : 0, i: /Italic|Oblique/i.test(sty) ? 1 : 0, ps: r.ps, installed: true, note};
}

// Advance of the space glyph in em, from the installed font file PowerPoint will use.
import * as mu from 'mupdf';
import {readFileSync} from 'node:fs';
const spaceCache = new Map<string, number>();
export function spaceEm(ps: string): number {
  if (spaceCache.has(ps)) return spaceCache.get(ps)!;
  let em = 0.25;
  try {
    const line = execFileSync('fc-list', ['--format', '%{file}|%{postscriptname}\n'], {encoding: 'utf8', maxBuffer: 64e6}).split('\n').find(l => l.endsWith('|' + ps));
    if (line) { const f = new mu.Font(ps, readFileSync(line.split('|')[0])); em = f.advanceGlyph(f.encodeCharacter(32), 0); }
  } catch {}
  spaceCache.set(ps, em); return em;
}
