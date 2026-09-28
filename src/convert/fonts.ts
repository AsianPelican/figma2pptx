// Map a Figma face (PostScript name, or family + weight + italic) to the name PowerPoint selects it by:
// the legacy family (OpenType name ID 1) plus bold/italic flags from the legacy subfamily (name ID 2).
//
// PowerPoint for Mac cannot select the weights of a variable font (it falls back to its default font), so only
// static faces count as installed. A variable-only family is reported, never silently used.
import {execFileSync} from 'node:child_process';
import {readFileSync} from 'node:fs';
import * as mu from 'mupdf';

// exact: the face Figma uses is installed as a static face.
// substitute: the family is installed, but not this weight or style; the nearest installed face is used.
// variable-only: the family is installed only as a variable font, which PowerPoint cannot render by weight.
// missing: the family is not installed at all.
export type FaceStatus = 'exact' | 'substitute' | 'variable-only' | 'missing';
export type Face = {typeface: string, b: 0 | 1, i: 0 | 1, ps: string, status: FaceStatus, note?: string, file?: string, index?: number};
// One installed face as fontconfig lists it: family and style names (typographic first, legacy last), PostScript
// name, fontconfig weight, and the font file (index within a collection); `variable` marks every face that
// comes from a variable font file, named instances included.
export type FontRow = {fams: string[], styles: string[], ps: string, weight: number, file?: string, index?: number, variable?: boolean};

export interface FontResolver {
  mapFace(ps: string | null | undefined, family: string, weight: number, italic: boolean): Face;
  // Advance of the space glyph in em, in the installed font file PowerPoint will use.
  spaceEm(ps: string): number;
}

const cssWeightToFc: Record<number, number> = {100: 0, 200: 40, 300: 50, 400: 80, 500: 100, 600: 180, 700: 200, 800: 205, 900: 210};

export function mapFaceIn(rows: FontRow[], ps: string | null | undefined, family: string, weight: number, italic: boolean): Face {
  const statics = rows.filter(x => !x.variable);
  let r = ps ? statics.find(x => x.ps === ps) : undefined;
  let note: string | undefined, status: FaceStatus = 'exact';
  if (!r) {
    // No PostScript name (Figma's bundled Inter) or not installed: pick the installed static face of the family nearest in weight.
    const want = cssWeightToFc[Math.round(weight / 100) * 100] ?? 80;
    const cands = statics.filter(x => x.fams.includes(family) && /Italic/i.test(x.styles.join()) === italic && !x.ps.includes('Variable'));
    cands.sort((a, b) => Math.abs(a.weight - want) - Math.abs(b.weight - want));
    r = cands[0];
    if (r && r.weight !== want) { status = 'substitute'; note = `${family} ${weight}${italic ? ' italic' : ''} not installed; nearest installed face ${r.ps}`; }
  }
  if (!r) {
    const variableOnly = rows.some(x => x.variable && x.fams.includes(family));
    return {
      typeface: family, b: weight >= 700 ? 1 : 0, i: italic ? 1 : 0, ps: ps || '',
      status: variableOnly ? 'variable-only' : 'missing',
      note: variableOnly ? `${family} is installed only as a variable font, which PowerPoint cannot render by weight; install its static faces` : `${ps || family} not installed`,
    };
  }
  const fam = r.fams[r.fams.length - 1], sty = r.styles[r.styles.length - 1];
  return {typeface: fam, b: /Bold/i.test(sty) ? 1 : 0, i: /Italic|Oblique/i.test(sty) ? 1 : 0, ps: r.ps, status, note, file: r.file, index: r.index};
}

export function parseFcList(out: string): FontRow[] {
  const rows = out.split('\n').filter(Boolean).map(l => {
    const [f, s, ps, w, v, idx, ...file] = l.split('|');
    return {fams: f.split(','), styles: s.split(','), ps, weight: +w, variable: v === 'True', index: +idx & 0xffff, file: file.join('|')};
  });
  // A variable font lists itself (variable=True) and each named instance (variable=False): mark them all.
  const variableFiles = new Set(rows.filter(r => r.variable).map(r => r.file));
  for (const r of rows) r.variable = variableFiles.has(r.file);
  return rows;
}

// The fonts installed on this machine, through fontconfig (`fc-list`).
export function fontconfigResolver(): FontResolver {
  let rows: FontRow[] | null = null;
  const spaceCache = new Map<string, number>();
  const load = () => rows ??= parseFcList(execFileSync('fc-list', ['--format', '%{family}|%{style}|%{postscriptname}|%{weight}|%{variable}|%{index}|%{file}\n'], {encoding: 'utf8', maxBuffer: 64e6}));
  return {
    mapFace: (ps, family, weight, italic) => mapFaceIn(load(), ps, family, weight, italic),
    spaceEm(ps) {
      if (spaceCache.has(ps)) return spaceCache.get(ps)!;
      let em = 0.25;
      try {
        const r = load().find(x => x.ps === ps && !x.variable) ?? load().find(x => x.ps === ps);
        if (r?.file) { const f = new mu.Font(ps, readFileSync(r.file)); em = f.advanceGlyph(f.encodeCharacter(32), 0); }
      } catch {}
      spaceCache.set(ps, em); return em;
    },
  };
}
