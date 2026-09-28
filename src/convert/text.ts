// Figma TEXT node + its SVG <text> element -> one native PowerPoint text box.
//
// The SVG export (svg_outline_text=false) carries Figma's own layout: one <tspan x y> per laid-out line at the
// exact baseline. The node JSON carries the characters and per-character styles. Walking both together gives
// each laid-out line with its styled runs; each line becomes one PowerPoint line (a:br for soft wraps, a new
// paragraph for \n), in a box that does not wrap, so PowerPoint cannot re-flow it.
import {mul, ap, parseTransform, type Mat} from './geom';
import {esc, EMU_PER_PX as E} from './pptx';
import {hex, clr, type Rgba} from './paint';
import type {Face, FontResolver} from './fonts';

export type TextSeg = {x: number, y: number, t: string, fill: Rgba | null, deco: string | null};
export type Run = {t: string, style: any, fill: Rgba | null, deco: string | null, spcOverride?: number};
export type TextLine = {runs: Run[], x: number, y: number, hardBefore: boolean};
// One laid-out Figma line as placed: its text, pen x and baseline in frame px.
export type LineRecord = {node: string, t: string, x: number, base: number, align: string, size?: number, ignoreX?: boolean, rotated?: boolean};

export type TextContext = {
  frame: {x: number, y: number}; // frame origin in Figma's absolute coordinates
  corr: {dx: number, dy: number}; // measured offset for this box (pass 2), px
  kern: string;
  fonts: FontResolver;
  warn: (s: string) => void;
  onFace: (key: string, face: Face) => void;
};

// Every tspan, in slide space. Rotated text is laid out in its own frame (the box is rotated afterwards), so
// its tspans keep their local coordinates and the rotation is returned.
export function collectSegs(el: any, m: Mat, warn: (s: string) => void): {segs: TextSeg[], rotTm: Mat | null} {
  const texts = (el.tagName === 'text' ? [el] : Array.from(el.getElementsByTagName('text'))) as any[];
  const segs: TextSeg[] = [];
  let rotTm: Mat | null = null;
  for (const t of texts) {
    const tm = mul(m, parseTransform(t.getAttribute('transform')));
    if (Math.abs(tm[1]) > 1e-3 || Math.abs(tm[2]) > 1e-3) rotTm = tm;
    const fillA = t.getAttribute('fill'), fo = +(t.getAttribute('fill-opacity') ?? 1);
    const f = hex(fillA, warn); if (f) f[1] *= fo;
    for (const sp of Array.from(t.getElementsByTagName('tspan')) as any[]) {
      const [x, y] = rotTm ? [+sp.getAttribute('x'), +sp.getAttribute('y')] : ap(tm, +sp.getAttribute('x'), +sp.getAttribute('y'));
      segs.push({x, y, t: sp.textContent || '', fill: f, deco: t.getAttribute('text-decoration')});
    }
  }
  return {segs, rotTm};
}

// Group tspans into lines by baseline, top to bottom, left to right within a line.
export function groupLines(segs: TextSeg[]): TextSeg[][] {
  const lines: TextSeg[][] = [];
  for (const s of [...segs].sort((a, b) => a.y - b.y || a.x - b.x)) { const l = lines.find(l => Math.abs(l[0].y - s.y) < 0.01); if (l) l.push(s); else lines.push([s]); }
  lines.sort((a, b) => a[0].y - b[0].y);
  return lines;
}

// Walk node.characters alongside the laid-out lines to recover per-character styles (split into runs) and
// which lines start a new paragraph (a \n before them in the source; U+2028 is a line break inside one).
export function alignLines(lines: TextSeg[][], n: any, nodeId: string, warn: (s: string) => void): TextLine[] {
  const chars = n.characters as string, ov = n.characterStyleOverrides || [], tbl = n.styleOverrideTable || {};
  let p = 0;
  const L: TextLine[] = [];
  for (const segsOfLine of lines) {
    let hardBefore = false; const runs: Run[] = [];
    segsOfLine.forEach((s, si) => {
      const txt = s.t.replace(/\n$/, '').replace(/\u2028$/, '');
      // skip newlines that precede this line in the source (paragraph breaks)
      while (p < chars.length && chars[p] !== txt[0] && /[\n\u2028]/.test(chars[p])) { if (si === 0) hardBefore = hardBefore || chars[p] === '\n'; p++; }
      let start = p;
      if (chars.slice(p, p + txt.length) !== txt) { const k = chars.indexOf(txt, p); if (k >= 0) start = k; else warn(`text ${nodeId}: could not align "${txt.slice(0, 30)}"`); }
      // split by style overrides
      let cur = '', curSt = -1;
      for (let k = 0; k < txt.length; k++) { const stId = ov[start + k] || 0; if (stId !== curSt && cur) { runs.push({t: cur, style: {...n.style, ...(tbl[curSt] || {})}, fill: s.fill, deco: s.deco}); cur = ''; } curSt = stId; cur += txt[k]; }
      if (cur) runs.push({t: cur, style: {...n.style, ...(tbl[curSt] || {})}, fill: s.fill, deco: s.deco});
      p = start + txt.length;
      while (p < chars.length && chars[p] === '\u2028') p++;
    });
    L.push({runs, x: segsOfLine[0].x, y: segsOfLine[0].y, hardBefore});
  }
  return L;
}

export function paragraphs(L: TextLine[]): TextLine[][] {
  const paras: TextLine[][] = [];
  for (const l of L) { if (!paras.length || l.hardBefore) paras.push([l]); else paras[paras.length - 1].push(l); }
  return paras;
}

export const maxSize = (l: TextLine, fallback: number) => l.runs.length ? Math.max(...l.runs.map(r => r.style.fontSize)) : fallback;

// Line pitch Figma used inside a paragraph (median baseline step); a one-line paragraph uses its line height.
export function paraPitch(para: TextLine[], fallbackSize: number): number {
  const pitches = para.slice(1).map((l, k) => l.y - para[k].y);
  return pitches.length ? pitches.sort((a, b) => a - b)[Math.floor(pitches.length / 2)] : (para[0].runs[0].style.lineHeightPx || maxSize(para[0], fallbackSize) * 1.2);
}

// PowerPoint for Mac rounds exact spacing (spcPts) to whole points, and percent spacing (spcPct, pitch =
// pct x 1.2 x size) to whole percent. Use whichever lands closer to Figma's pitch.
export function lineSpacing(pitch: number, size: number): string {
  const pts = Math.round(pitch * 0.75), errPts = Math.abs(pts / 0.75 - pitch);
  const pct = Math.round(pitch / (1.2 * size) * 100), errPct = Math.abs(pct / 100 * 1.2 * size - pitch);
  return errPct < errPts ? `<a:spcPct val="${pct * 1000}"/>` : `<a:spcPts val="${pts * 100}"/>`;
}

// Centred and right-aligned soft-wrapped lines: PowerPoint counts the line's trailing space in its width, Figma
// does not. Keep the space (the copy is unchanged) but give it zero advance, so both centre the same glyphs.
export function zeroAdvanceTrailingSpace(runs: Run[], align: string, softWrapped: boolean, fonts: FontResolver): Run[] {
  const out = [...runs];
  const lastR = out[out.length - 1];
  if ((align === 'ctr' || align === 'r') && softWrapped && lastR && /\s+$/.test(lastR.t) && lastR.t.trim()) {
    const tr = lastR.t.match(/\s+$/)![0];
    out[out.length - 1] = {...lastR, t: lastR.t.slice(0, -tr.length)};
    const fc = fonts.mapFace(lastR.style.fontPostScriptName, lastR.style.fontFamily, lastR.style.fontWeight, !!lastR.style.italic);
    out.push({...lastR, t: tr, spcOverride: -fonts.spaceEm(fc.ps) * lastR.style.fontSize});
  }
  return out;
}

// A rotated box's own size, from its axis-aligned bounds (W = w|cos| + h|sin|, H = w|sin| + h|cos|).
export function unrotatedSize(width: number, height: number, rotTm: Mat): [number, number] {
  let w = width, h = height;
  const t = Math.atan2(rotTm[1], rotTm[0]), c = Math.abs(Math.cos(t)), s = Math.abs(Math.sin(t)), det = c * c - s * s;
  if (Math.abs(det) > 1e-3) { w = (width * c - height * s) / det; h = (height * c - width * s) / det; }
  return [w, h];
}

const ALIGN: Record<string, string> = {LEFT: 'l', CENTER: 'ctr', RIGHT: 'r', JUSTIFIED: 'just'};

// The <p:sp> for one TEXT node, or null when it has no visible characters. `sid` is the shape id to use.
export function textShape(el: any, m: Mat, n: any, nodeId: string, opacity: number, sid: number, ctx: TextContext): {xml: string, lines: LineRecord[]} | null {
  const {segs, rotTm} = collectSegs(el, m, ctx.warn);
  if (!segs.length || !n.characters.trim()) return null;
  const paras = paragraphs(alignLines(groupLines(segs), n, nodeId, ctx.warn));
  const nb = n.absoluteBoundingBox, bx = rotTm ? 0 : nb.x - ctx.frame.x, by = rotTm ? 0 : nb.y - ctx.frame.y;
  const algn = ALIGN[n.style.textAlignHorizontal] || 'l';
  const bullets = (n.lineTypes || []) as string[];
  let xml = '', pi = 0;
  const lines: LineRecord[] = [];
  for (const para of paras) {
    const S = Math.max(...para.map(l => maxSize(l, n.style.fontSize)));
    const lnSpc = lineSpacing(paraPitch(para, n.style.fontSize), S);
    const isBullet = bullets[pi] && bullets[pi] !== 'NONE';
    const marL = Math.max(0, para[0].x - bx);
    const indentXml = algn === 'l' && marL > 0.05 ? ` marL="${Math.round(marL * E)}"${isBullet ? ` indent="${-Math.round(marL * E)}"` : ''}` : '';
    xml += `<a:p><a:pPr algn="${algn}"${indentXml}><a:lnSpc>${lnSpc}</a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft>${isBullet ? '<a:buFont typeface="Arial"/><a:buChar char="•"/>' : '<a:buNone/>'}</a:pPr>`;
    para.forEach((l, li) => {
      if (li) xml += `<a:br><a:rPr lang="en-US" sz="${Math.round(maxSize(l, n.style.fontSize) * 75)}"/></a:br>`;
      for (const r of zeroAdvanceTrailingSpace(l.runs, algn, li < para.length - 1, ctx.fonts)) {
        const s = r.style, face = ctx.fonts.mapFace(s.fontPostScriptName, s.fontFamily, s.fontWeight, !!s.italic);
        ctx.onFace(`${s.fontPostScriptName || s.fontFamily + ' ' + s.fontWeight}`, face);
        const fl = r.fill || ['000000', 1];
        xml += `<a:r><a:rPr lang="en-US" sz="${Math.round(s.fontSize * 75)}" b="${face.b}" i="${face.i}"${r.deco === 'underline' || s.textDecoration === 'UNDERLINE' ? ' u="sng"' : ''} spc="${Math.round((r.spcOverride ?? s.letterSpacing ?? 0) * 75)}" kern="${ctx.kern}" dirty="0"><a:solidFill>${clr(fl[0], fl[1] * opacity)}</a:solidFill><a:latin typeface="${esc(face.typeface)}"/><a:ea typeface="${esc(face.typeface)}"/><a:cs typeface="${esc(face.typeface)}"/></a:rPr><a:t>${esc(r.t)}</a:t></a:r>`;
      }
      const text = l.runs.map(r => r.t).join('');
      lines.push({node: nodeId, t: text, x: l.x, base: l.y, align: algn, size: maxSize(l, n.style.fontSize), ...(/^\s/.test(text) ? {ignoreX: true} : {}), ...(rotTm ? {rotated: true} : {})});
    });
    xml += `<a:endParaRPr lang="en-US" sz="${Math.round(S * 75)}"/></a:p>`;
    pi += 1; // lineTypes are per source line (paragraph) in Figma
  }
  let [w, h] = [nb.width, nb.height];
  let ox = bx + ctx.corr.dx, oy = by + ctx.corr.dy, rotAttr = '';
  if (rotTm) {
    [w, h] = unrotatedSize(nb.width, nb.height, rotTm);
    const c = ap(rotTm, w / 2, h / 2); ox = c[0] - w / 2; oy = c[1] - h / 2;
    rotAttr = ` rot="${Math.round(((Math.atan2(rotTm[1], rotTm[0]) * 180 / Math.PI) + 360) % 360 * 60000)}"`;
  }
  const sp = `<p:sp><p:nvSpPr><p:cNvPr id="${sid}" name="${esc(n.name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${rotAttr}><a:off x="${Math.round(ox * E)}" y="${Math.round(oy * E)}"/><a:ext cx="${Math.max(1, Math.round(w * E))}" cy="${Math.max(1, Math.round(h * E))}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>${xml}</p:txBody></p:sp>`;
  return {xml: sp, lines};
}
