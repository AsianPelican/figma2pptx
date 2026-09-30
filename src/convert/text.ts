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
export type Run = {
  t: string, style: any, fill: Rgba | null, deco: string | null, spcOverride?: number,
  sourceStart?: number, spanStart?: number, spanX?: number, spanNextX?: number, spanLength?: number,
};
export type TextLine = {runs: Run[], x: number, y: number, hardBefore: boolean};
// One laid-out Figma line as placed: its text, pen x and baseline in frame px.
// `measure`: the part PowerPoint still draws as native text when some of the line's glyphs are Figma vector outlines.
export type LineRecord = {node: string, t: string, x: number, base: number, align: string, size?: number, ignoreX?: boolean, rotated?: boolean, unmeasurableReason?: string, substituted?: boolean, measure?: MeasurableText};
export type MeasurableText = {t: string, x: number, ignoreX?: boolean};

export type TextContext = {
  frame: {x: number, y: number}; // frame origin in Figma's absolute coordinates
  corr: {dx: number, dy: number}; // measured offset for this box (pass 2), px
  kern: string;
  fonts: FontResolver;
  warn: (s: string) => void;
  mapStyle?: (style: any) => {style: any, substituted: boolean};
  onFace: (key: string, face: Face) => void;
  missingGlyphs?: Map<number, {advance: number}>; // source UTF-16 index -> Figma advance in px
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
    const fillA = t.getAttribute('fill'), fillOpacity = t.getAttribute('fill-opacity');
    const fo = fillOpacity === null || fillOpacity === '' ? 1 : +fillOpacity;
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

const segText = (s: TextSeg) => s.t.replace(/\n$/, '').replace(/\u2028$/, '');

// Figma may emit styled SVG fragments in paint order instead of character order. Recover the only order whose
// fragments cover the next source substring exactly; the node's `characters` field is the ordering authority.
export function sourceOrderedSegs(segs: TextSeg[], source: string): TextSeg[] | null {
  const byText = new Map<string, TextSeg[]>();
  for (const seg of segs) { const text = segText(seg); if (text) byText.set(text, [...(byText.get(text) || []), seg]); }
  const texts = [...byText.keys()].sort((a, b) => b.length - a.length);
  const counts = texts.map(t => byText.get(t)!.length), sequence: number[] = [], dead = new Set<string>();
  const visit = (at: number): boolean => {
    if (at === source.length) return counts.every(n => n === 0);
    const key = `${at}|${counts.join(',')}`;
    if (dead.has(key)) return false;
    for (let i = 0; i < texts.length; i++) {
      if (!counts[i] || !source.startsWith(texts[i], at)) continue;
      counts[i]--; sequence.push(i);
      if (visit(at + texts[i].length)) return true;
      sequence.pop(); counts[i]++;
    }
    dead.add(key);
    return false;
  };
  if (!visit(0)) return null;
  const used = new Map<string, number>();
  return sequence.map(i => { const text = texts[i], at = used.get(text) || 0; used.set(text, at + 1); return byText.get(text)![at]; });
}

// Walk node.characters alongside the laid-out lines to recover per-character styles (split into runs) and
// which lines start a new paragraph (a \n before them in the source; U+2028 is a line break inside one).
export function alignLines(lines: TextSeg[][], n: any, nodeId: string, warn: (s: string) => void): TextLine[] {
  const chars = n.characters as string, ov = n.characterStyleOverrides || [], tbl = n.styleOverrideTable || {};
  let p = 0;
  const L: TextLine[] = [];
  for (const segsOfLine of lines) {
    let hardBefore = false; const runs: Run[] = [];
    // Skip separators before the line, then reorder its SVG fragments against the exact next source substring.
    while (p < chars.length && /[\n\u2028]/.test(chars[p])) { hardBefore = hardBefore || chars[p] === '\n'; p++; }
    const lineLength = segsOfLine.reduce((sum, s) => sum + segText(s).length, 0);
    const ordered = sourceOrderedSegs(segsOfLine, chars.slice(p, p + lineLength));
    const sequence = ordered ?? segsOfLine;
    const visual = [...segsOfLine].sort((a, b) => a.x - b.x);
    const nextX = new Map(visual.slice(0, -1).map((s, k) => [s, visual[k + 1].x]));
    sequence.forEach(s => {
      const txt = segText(s);
      let start = p;
      if (chars.slice(p, p + txt.length) !== txt) { const k = chars.indexOf(txt, p); if (k >= 0) start = k; else warn(`text ${nodeId}: could not align "${txt.slice(0, 30)}"`); }
      // split by style overrides
      let cur = '', curSt = -1, curStart = start;
      const push = () => { if (cur) runs.push({t: cur, style: {...n.style, ...(tbl[curSt] || {})}, fill: s.fill, deco: s.deco, sourceStart: curStart, spanStart: start, spanX: s.x, spanNextX: nextX.get(s), spanLength: txt.length}); };
      for (let k = 0; k < txt.length; k++) { const stId = ov[start + k] || 0; if (stId !== curSt && cur) { push(); cur = ''; curStart = start + k; } curSt = stId; cur += txt[k]; }
      push();
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
  return pitches.length ? pitches.sort((a, b) => a - b)[Math.floor(pitches.length / 2)] : (para[0].runs[0]?.style.lineHeightPx || maxSize(para[0], fallbackSize) * 1.2);
}

// PowerPoint for Mac rounds exact spacing (spcPts) to whole points, and percent spacing (spcPct, pitch =
// pct x 1.2 x size) to whole percent. Use whichever lands closer to Figma's pitch.
export function lineSpacing(pitch: number, size: number): string {
  const pts = Math.round(pitch * 0.75), errPts = Math.abs(pts / 0.75 - pitch);
  const pct = Math.round(pitch / (1.2 * size) * 100), errPct = Math.abs(pct / 100 * 1.2 * size - pitch);
  return errPct < errPts ? `<a:spcPct val="${pct * 1000}"/>` : `<a:spcPts val="${pts * 100}"/>`;
}

// REST TypeStyle resolves letterSpacing to px. Plugin API callers retain the original LetterSpacing object,
// whose PERCENT value is relative to font size. Accept both, plus the flattened unit shape used by adapters.
export function drawingmlCharacterSpacing(style: any): number {
  const spacing = style?.letterSpacing;
  const value = Number(typeof spacing === 'object' && spacing !== null ? spacing.value : spacing);
  if (!Number.isFinite(value)) return 0;
  const unit = typeof spacing === 'object' && spacing !== null ? spacing.unit : style?.letterSpacingUnit;
  const px = String(unit || 'PIXELS').toUpperCase() === 'PERCENT'
    ? value / 100 * Number(style?.fontSize || 0)
    : value;
  return Math.round(px * 75); // px -> pt -> 1/100 pt
}

function figmaKernFlag(style: any): boolean | null {
  const features = style?.opentypeFlags ?? style?.fontFeatures ?? style?.openTypeFeatures;
  if (!features) return null;
  let value: unknown;
  if (Array.isArray(features)) {
    const match = features.find(feature => {
      if (typeof feature === 'string') return feature.toUpperCase() === 'KERN';
      const tag = feature?.tag ?? feature?.name ?? feature?.feature;
      return String(tag || '').toUpperCase() === 'KERN';
    });
    if (typeof match === 'string') value = true;
    else value = match?.value ?? match?.enabled;
  } else {
    value = features.KERN ?? features.kern;
  }
  return value === undefined || value === null ? null : value === true || value === 1 || value === '1';
}

// DrawingML kern is the minimum font size for pair kerning, in 1/100 pt. Zero disables it.
export function drawingmlKerning(style: any, fallback: string): string {
  const enabled = figmaKernFlag(style);
  if (enabled === null) return fallback;
  if (!enabled) return '0';
  return Number(fallback) > 0 ? fallback : '100';
}

// Centred and right-aligned soft-wrapped lines: PowerPoint counts the line's trailing space in its width, Figma
// does not. Keep the space (the copy is unchanged) but give it zero advance, so both centre the same glyphs.
export function zeroAdvanceTrailingSpace(runs: Run[], align: string, softWrapped: boolean, fonts: FontResolver, mapStyle: (style: any) => any = style => style): Run[] {
  const out = [...runs];
  const lastR = out[out.length - 1];
  if ((align === 'ctr' || align === 'r') && softWrapped && lastR && /\s+$/.test(lastR.t) && lastR.t.trim()) {
    const tr = lastR.t.match(/\s+$/)![0];
    out[out.length - 1] = {...lastR, t: lastR.t.slice(0, -tr.length)};
    const style = mapStyle(lastR.style);
    const fc = fonts.mapFace(style.fontPostScriptName, style.fontFamily, style.fontWeight, !!style.italic);
    out.push({...lastR, t: tr, spcOverride: -fonts.spaceEm(fc.ps) * lastR.style.fontSize});
  }
  return out;
}

// Replace a missing character with a supported blank having exactly Figma's advance. The visible character is
// overlaid as Figma's vector outline by build.ts; using a source-font blank prevents PowerPoint's wider,
// machine-dependent fallback glyph from moving everything that follows it.
export function replaceMissingGlyphs(runs: Run[], missing: Map<number, {advance: number}>, fonts: FontResolver, mapStyle: (style: any) => any = style => style): Run[] {
  const out: Run[] = [];
  for (const run of runs) {
    if (run.sourceStart === undefined) { out.push(run); continue; }
    let chunk = '', chunkStart = run.sourceStart, at = run.sourceStart;
    const flush = () => { if (chunk) out.push({...run, t: chunk, sourceStart: chunkStart}); chunk = ''; };
    for (const char of run.t) {
      const plan = missing.get(at);
      if (plan) {
        flush();
        const style = mapStyle(run.style), face = fonts.mapFace(style.fontPostScriptName, style.fontFamily, style.fontWeight, !!style.italic);
        out.push({...run, t: '\u00a0', sourceStart: at, spcOverride: plan.advance - fonts.spaceEm(face.ps) * style.fontSize});
        chunkStart = at + char.length;
      } else chunk += char;
      at += char.length;
    }
    flush();
  }
  return out;
}

// The first stretch of native text on a line whose missing glyphs are drawn as Figma vector outlines, anchored at
// its own Figma pen x, so the line's placement can still be measured in PowerPoint's PDF. A stretch that follows
// an outlined glyph is anchored only when its first visible character starts a Figma span; otherwise its x is
// unknown and only the baseline is measured.
export function measurableText(runs: Run[], missing: Map<number, unknown>, lineX: number): MeasurableText | undefined {
  let t = '', x = lineX, ignoreX = false, afterMissing = false;
  const done = () => t.trim() ? {t, x, ...(ignoreX || (!afterMissing && /^\s/.test(t)) ? {ignoreX: true} : {})} : undefined;
  for (const r of runs) {
    if (r.sourceStart === undefined) return undefined;
    let at = r.sourceStart;
    for (const char of r.t) {
      if (missing.has(at)) {
        if (t.trim()) return done();
        t = ''; afterMissing = true; ignoreX = false;
      } else if (afterMissing && !t && !char.trim()) {
        // whitespace between an outlined glyph and the next visible character is not part of the measured text
      } else {
        if (afterMissing && !t) {
          const anchored = at === r.sourceStart && r.spanStart === r.sourceStart && r.spanX !== undefined;
          x = anchored ? r.spanX! : lineX; ignoreX = !anchored;
        }
        t += char;
      }
      at += char.length;
    }
  }
  return done();
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
  const levels = (n.lineIndentations || []) as number[];
  let xml = '', pi = 0;
  const lines: LineRecord[] = [];
  for (const para of paras) {
    const S = Math.max(...para.map(l => maxSize(l, n.style.fontSize)));
    const lnSpc = lineSpacing(paraPitch(para, n.style.fontSize), S);
    const listType = bullets[pi] || 'NONE', isBullet = listType !== 'NONE';
    const level = isBullet ? Math.max(1, levels[pi] || 1) : 0;
    const marL = Math.max(0, para[0].x - bx);
    const hanging = isBullet ? marL / level : 0;
    const indentXml = algn === 'l' && marL > 0.05 ? ` marL="${Math.round(marL * E)}"${isBullet ? ` indent="${-Math.round(hanging * E)}"` : ''}` : '';
    const paraStyle = para[0].runs[0]?.style || n.style;
    const mappedParaStyle = ctx.mapStyle?.(paraStyle).style || paraStyle;
    const nextListType = bullets[pi + 1] || 'NONE';
    const spacingPx = isBullet && nextListType !== 'NONE' ? (paraStyle.listSpacing ?? n.style.listSpacing ?? 0) : (paraStyle.paragraphSpacing ?? n.style.paragraphSpacing ?? 0);
    let listXml = '<a:buNone/>';
    if (isBullet) {
      const face = ctx.fonts.mapFace(mappedParaStyle.fontPostScriptName, mappedParaStyle.fontFamily, mappedParaStyle.fontWeight, !!mappedParaStyle.italic);
      listXml = `<a:buFont typeface="${esc(face.typeface)}"/><a:buSzPct val="100000"/>${listType === 'ORDERED' ? '<a:buAutoNum type="arabicPeriod"/>' : '<a:buChar char="•"/>'}`;
    }
    xml += `<a:p><a:pPr algn="${algn}"${isBullet ? ` lvl="${level - 1}"` : ''}${indentXml}><a:lnSpc>${lnSpc}</a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="${Math.round(spacingPx * 75)}"/></a:spcAft>${listXml}</a:pPr>`;
    para.forEach((l, li) => {
      if (li) xml += `<a:br><a:rPr lang="en-US" sz="${Math.round(maxSize(l, n.style.fontSize) * 75)}"/></a:br>`;
      const trailing = zeroAdvanceTrailingSpace(l.runs, algn, li < para.length - 1, ctx.fonts, s => ctx.mapStyle?.(s).style || s);
      const emitted = ctx.missingGlyphs ? replaceMissingGlyphs(trailing, ctx.missingGlyphs, ctx.fonts, s => ctx.mapStyle?.(s).style || s) : trailing;
      for (const r of emitted) {
        const sourceStyle = r.style, mapped = ctx.mapStyle?.(sourceStyle) || {style: sourceStyle, substituted: false};
        const s = mapped.style, face = ctx.fonts.mapFace(s.fontPostScriptName, s.fontFamily, s.fontWeight, !!s.italic);
        ctx.onFace(`${sourceStyle.fontPostScriptName || sourceStyle.fontFamily + ' ' + sourceStyle.fontWeight}`, face);
        const fl = r.fill || ['000000', 1];
        const spc = r.spcOverride === undefined ? drawingmlCharacterSpacing(s) : Math.round(r.spcOverride * 75);
        xml += `<a:r><a:rPr lang="en-US" sz="${Math.round(s.fontSize * 75)}" b="${face.b}" i="${face.i}"${r.deco === 'underline' || s.textDecoration === 'UNDERLINE' ? ' u="sng"' : ''} spc="${spc}" kern="${drawingmlKerning(s, ctx.kern)}" dirty="0"><a:solidFill>${clr(fl[0], fl[1] * opacity)}</a:solidFill><a:latin typeface="${esc(face.typeface)}"/><a:ea typeface="${esc(face.typeface)}"/><a:cs typeface="${esc(face.typeface)}"/></a:rPr><a:t>${esc(r.t)}</a:t></a:r>`;
      }
      const text = l.runs.map(r => r.t).join('');
      const translucent = l.runs.some(r => (r.fill?.[1] ?? 1) * opacity < 0.999);
      const substituted = l.runs.some(r => ctx.mapStyle?.(r.style).substituted);
      const outlinedMissing = ctx.missingGlyphs && l.runs.some(r => r.sourceStart !== undefined && [...ctx.missingGlyphs!.keys()].some(i => i >= r.sourceStart! && i < r.sourceStart! + r.t.length));
      const measure = outlinedMissing ? measurableText(l.runs, ctx.missingGlyphs!, l.x) : undefined;
      const unmeasurableReason = outlinedMissing && !measure ? 'contains a Figma vector outline for a source-font missing glyph' : translucent ? 'PowerPoint PDF rasterizes text with opacity below 100%' : undefined;
      if (text.trim()) lines.push({node: nodeId, t: text, x: l.x, base: l.y, align: algn, size: maxSize(l, n.style.fontSize), ...(/^\s/.test(text) ? {ignoreX: true} : {}), ...(rotTm ? {rotated: true} : {}), ...(unmeasurableReason ? {unmeasurableReason} : {}), ...(substituted ? {substituted: true} : {}), ...(measure ? {measure} : {})});
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
