// Figma frames -> editable PPTX, 1:1 by construction.
//
// Inputs, per frame (see FrameSource):
//   the node JSON      text styles (PostScript face, size, tracking, leading), effects, fills, bounds
//   the SVG export     text kept live: exact line breaks and baselines (one <tspan> per laid-out line), and every
//                      vector's geometry, stroke, dash pattern and gradient in paint order, with node ids
//   PNG renders        for what PowerPoint cannot draw natively (image fills, effects, masks, angular and radial
//                      gradients), and for the plates behind background-blur panels
// Text becomes native PowerPoint text with Figma's own line breaks as hard breaks; vectors become custGeom shapes
// with native strokes, dashes, caps and alpha.
import {DOMParser} from '@xmldom/xmldom';
import {readFileSync, existsSync} from 'node:fs';
import {buildPptx, slideXml, esc, EMU_PER_PX as E, REL_IMAGE, type Media, type Slide} from './pptx';
import {parseTransform, parsePath, transformSegs, bbox, groupForFill, custGeom, rectPath, ellipsePath, mul, scaleOf, I, type Mat} from './geom';
import {paint} from './paint';
import {textShape, type LineRecord} from './text';
import type {Face, FaceStatus, FontResolver} from './fonts';
import {embedFaces, type EmbedResult} from './embed';
import type {ImageOps} from './images';

export interface FrameSource {
  document(frameId: string): any; // the frame's node JSON (`nodes[id].document` of /v1/files/:key/nodes)
  svg(frameId: string): string; // svg_outline_text=false, svg_include_node_id=true, svg_simplify_stroke=false
  ensureRasters(ids: string[], scale: number): Promise<void>; // PNG renders, use_absolute_bounds=true
  rasterFile(id: string, scale: number): string;
  platePath(id: string, scale: number): string; // where to write a derived blur plate
}

export type Corrections = Record<string, {dx: number, dy: number}>;
export type BuildOptions = {scale: number, kern: string, corr: Corrections, embedFonts?: boolean, progress?: (step: string) => void};
export type BuildEnv = {fonts: FontResolver, images: ImageOps};
export type FrameStats = {id: string, name: string, text: number, shapes: number, rasters: number, plates: number, rasterText: number};
export type BuildReport = {
  frames: FrameStats[];
  warnings: string[];
  fonts: Record<string, string>; // Figma face -> PowerPoint typeface (+b/+i) and any substitution note
  faces: {figma: string, typeface: string, b: 0 | 1, i: 0 | 1, status: FaceStatus, note?: string}[];
  embeddedFonts: EmbedResult[];
  textLines: (LineRecord & {frame: string})[];
  rasters: Record<string, string>; // node id -> why it is a picture
  blurPanels: string[];
};

const vis = (p: any) => p.visible !== false;
const count = (n: any): number => 1 + (n.children || []).reduce((a: number, c: any) => a + count(c), 0);
const hasText = (n: any): boolean => n.visible !== false && (n.type === 'TEXT' || (n.children || []).some(hasText));

// Why a node must be a picture, or null when PowerPoint can draw it natively.
export function rasterReason(n: any, isRoot: boolean): string | null {
  if (isRoot || n.type === 'TEXT') return null;
  if ((n.fills || []).some((f: any) => vis(f) && f.type === 'IMAGE')) return 'image fill';
  if ((n.effects || []).some((e: any) => vis(e) && e.type === 'BACKGROUND_BLUR')) return 'blur';
  if ((n.effects || []).some((e: any) => vis(e))) return 'effect ' + n.effects.filter(vis).map((e: any) => e.type).join('+');
  if ([...(n.fills || []), ...(n.strokes || [])].some((f: any) => vis(f) && /ANGULAR|DIAMOND|RADIAL/.test(f.type))) return 'gradient ' + [...(n.fills || []), ...(n.strokes || [])].filter(vis).map((f: any) => f.type).join('+');
  if ((n.children || []).some((c: any) => c.isMask && vis(c))) return 'mask';
  if (n.type === 'FRAME' && n.clipsContent && (n.children || []).length) return 'clipping frame';
  if (n.children && count(n) > 150 && !hasText(n)) return `dense group (${count(n)} nodes)`;
  return null;
}

export async function buildDeck(src: FrameSource, FRAMES: string[], opts: BuildOptions, env: BuildEnv): Promise<{pptx: Uint8Array, report: BuildReport}> {
  const SCALE = opts.scale, CORR = opts.corr, KERN = opts.kern;
  const report: BuildReport = {frames: [], warnings: [], fonts: {}, faces: [], embeddedFonts: [], textLines: [], rasters: {}, blurPanels: []};
  const faces = new Map<string, Face>();
  const warn = (s: string) => { if (!report.warnings.includes(s)) report.warnings.push(s); };

  // ---------- 1. node tree ----------
  const NODES: Record<string, {document: any}> = Object.fromEntries(FRAMES.map(f => [f, {document: src.document(f)}]));
  const byId = new Map<string, any>(), parentOf = new Map<string, any>();
  for (const f of FRAMES) (function idx(n: any, p: any) { byId.set(n.id, n); if (p) parentOf.set(n.id, p); for (const c of n.children || []) idx(c, n); })(NODES[f].document, null);

  // ---------- 2. decide what must be raster ----------
  const raster = new Map<string, string>(); // node id -> reason (raster roots)
  const blurNodes = new Set<string>();
  function scan(n: any, isRoot: boolean) {
    if (!vis(n)) return;
    const r = rasterReason(n, isRoot);
    if (r === 'blur') { blurNodes.add(n.id); return; }
    if (r) { raster.set(n.id, r); return; }
    for (const c of n.children || []) scan(c, false);
  }
  for (const f of FRAMES) scan(NODES[f].document, true);
  const ancestorsIn = (id: string, set: Map<string, string>): string | null => {
    let n = byId.get(id);
    while (n) { if (set.has(n.id)) return n.id; n = parentOf.get(n.id); }
    return null;
  };
  const blurMap = () => new Map([...blurNodes].map(b => [b, '']));
  const nextEl = (el: any) => { let s = el.nextSibling; while (s && s.nodeType !== 1) s = s.nextSibling; return s; };
  // SVG safety net: anything Figma had to express with clip-path, mask, filter or foreignObject is raster too.
  const svgDocs = new Map<string, any>();
  for (const f of FRAMES) {
    const doc = new DOMParser().parseFromString(src.svg(f), 'image/svg+xml') as any;
    svgDocs.set(f, doc);
    const walk = (el: any, owner: string | null, depth: number) => {
      if (el.nodeType !== 1) return;
      const id = el.getAttribute('data-node-id') || owner;
      const inRaster = id && ancestorsIn(id, raster);
      const special = (depth > 1 && (el.getAttribute('clip-path') || el.getAttribute('mask'))) || el.getAttribute('filter') || (el.tagName === 'foreignObject' && !nextEl(el)?.getAttribute('data-figma-bg-blur-radius'));
      if (special && id && !inRaster && !ancestorsIn(id, blurMap()) && byId.get(id)?.type !== 'TEXT' && id !== f) { raster.set(id, 'svg ' + (el.tagName === 'foreignObject' ? 'foreignObject' : el.getAttribute('filter') ? 'filter' : el.getAttribute('mask') ? 'mask' : 'clip-path')); return; }
      for (const c of Array.from(el.childNodes || []) as any[]) if (c.tagName !== 'defs') walk(c, id, depth + 1);
    };
    walk(doc.documentElement, null, 0);
  }
  for (const [id, why] of raster) if (hasText(byId.get(id))) warn(`raster ${id} "${byId.get(id).name}" (${why}) contains text, which is baked into the image`);

  // ---------- 3. rasters ----------
  const frameOf = (id: string) => { let n = byId.get(id); while (parentOf.get(n.id)) n = parentOf.get(n.id); return n; };
  const paintIndex = (id: string) => { const order: string[] = []; (function w(n: any) { order.push(n.id); for (const c of n.children || []) w(c); })(frameOf(id)); return order.indexOf(id); };
  // Blur plates need renders of everything painted beneath each panel.
  const under = new Map<string, string[]>();
  for (const id of blurNodes) {
    const n = byId.get(id), b = n.absoluteBoundingBox, list: string[] = [];
    let cur = n;
    while (parentOf.get(cur.id)) {
      const p = parentOf.get(cur.id);
      for (const s of p.children) { if (s.id === cur.id) break; const r = s.absoluteRenderBounds || s.absoluteBoundingBox; if (vis(s) && r && r.x < b.x + b.width && r.x + r.width > b.x && r.y < b.y + b.height && r.y + r.height > b.y) list.push(s.id); }
      cur = p;
    }
    // paint order: outermost ancestors' earlier siblings first
    under.set(id, list.reverse().sort((a, c) => paintIndex(a) - paintIndex(c)));
  }
  const toRender = new Set([...raster.keys(), ...[...under.values()].flat()]);
  if (toRender.size) opts.progress?.(`renders from Figma (${toRender.size} nodes)`);
  await src.ensureRasters([...raster.keys()], SCALE);
  await src.ensureRasters([...new Set([...under.values()].flat())].filter(id => !raster.has(id)), SCALE);
  await src.ensureRasters([...new Set([...under.values()].flat())].filter(id => raster.has(id)), SCALE);
  const rasterFile = (id: string) => src.rasterFile(id, SCALE);

  // ---------- 4. emit ----------
  const media: Media[] = [];
  const mediaByPath = new Map<string, string>();
  function addMedia(path: string): string {
    if (mediaByPath.has(path)) return mediaByPath.get(path)!;
    const ext = path.split('.').pop()!, name = `image${media.length + 1}.${ext}`;
    media.push({name, data: readFileSync(path), ct: ext === 'jpg' ? 'image/jpeg' : 'image/png'});
    mediaByPath.set(path, name); return name;
  }

  function frameEmit(frameId: string) {
    const frame = NODES[frameId].document, F = frame.absoluteBoundingBox;
    const doc: any = svgDocs.get(frameId);
    const defs = new Map<string, any>();
    for (const el of Array.from(doc.getElementsByTagName('*')) as any[]) if (el.getAttribute('id')) defs.set(el.getAttribute('id'), el);
    const rels: Slide['rels'] = []; let body = '', sid = 2;
    const done = new Set<string>();
    const stats = {text: 0, shapes: 0, rasters: 0, plates: 0, rasterText: 0};
    const relFor = (path: string) => { const name = addMedia(path); let r = rels.find(x => x.target === `../media/${name}`); if (!r) { r = {id: `rId${rels.length + 1}`, type: REL_IMAGE, target: `../media/${name}`}; rels.push(r); } return r.id; };
    const ancOpacity = (id: string, includeSelf: boolean) => { let a = 1, n = includeSelf ? byId.get(id) : parentOf.get(id); while (n && n.id !== frameId) { a *= n.opacity ?? 1; n = parentOf.get(n.id); } return a; };
    const pic = (name: string, path: string, x: number, y: number, w: number, h: number, alpha: number, geom = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>') => {
      body += `<p:pic><p:nvPicPr><p:cNvPr id="${sid++}" name="${esc(name)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${relFor(path)}">${alpha < 0.9995 ? `<a:alphaModFix amt="${Math.round(alpha * 100000)}"/>` : ''}</a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round(x * E)}" y="${Math.round(y * E)}"/><a:ext cx="${Math.round(w * E)}" cy="${Math.round(h * E)}"/></a:xfrm>${geom}</p:spPr></p:pic>`;
    };
    // A node render covers the node unclipped (bounding box, grown by effects such as shadows); the REST render
    // bounds are clipped by the frame, so pick whichever box matches the rendered pixel size.
    const placeRaster = (n: any, file: string) => {
      const [pw, ph] = env.images.size(file);
      const w = pw / SCALE, h = ph / SCALE;
      const boxes = [n.absoluteBoundingBox, n.absoluteRenderBounds].filter(Boolean);
      const b = boxes.sort((a: any, c: any) => (Math.abs(a.width - w) + Math.abs(a.height - h)) - (Math.abs(c.width - w) + Math.abs(c.height - h)))[0];
      if (Math.abs(b.width - w) + Math.abs(b.height - h) > 2) warn(`raster ${n.id} "${n.name}": ${w}x${h} matches neither bounding nor render box; placed at the bounding box`);
      return {x: b.x - F.x, y: b.y - F.y, w, h};
    };
    function emitRaster(id: string) {
      const n = byId.get(id), r = placeRaster(n, rasterFile(id));
      // Trim what lies off the slide; opaque renders are re-encoded as JPEG q85: same pixels on the slide, a
      // fraction of the bytes. An image fill can carry alpha, so only opaque renders become JPEG.
      const x0 = Math.max(0, r.x), y0 = Math.max(0, r.y), x1 = Math.min(F.width, r.x + r.w), y1 = Math.min(F.height, r.y + r.h);
      let file = rasterFile(id);
      const trim = x0 > r.x + 0.5 || y0 > r.y + 0.5 || x1 < r.x + r.w - 0.5 || y1 < r.y + r.h - 0.5;
      const opaque = env.images.isOpaque(file);
      if (opaque || trim) {
        const out = file.replace(/\.(jpg|png)$/, opaque ? '.slide.jpg' : '.slide.png');
        if (!existsSync(out)) env.images.slideCopy(file, out, trim ? {w: Math.round((x1 - x0) * SCALE), h: Math.round((y1 - y0) * SCALE), x: Math.round((x0 - r.x) * SCALE), y: Math.round((y0 - r.y) * SCALE)} : null, opaque);
        file = out;
      }
      if (trim) Object.assign(r, {x: x0, y: y0, w: x1 - x0, h: y1 - y0});
      pic(n.name, file, r.x, r.y, r.w, r.h, ancOpacity(id, false));
      stats.rasters++; if (hasText(n)) stats.rasterText++;
    }
    function emitPlate(id: string) {
      // Background blur: composite what Figma paints beneath the panel, blur it as Figma does (sigma = radius / 2,
      // as CSS blur()), clip to the panel.
      const n = byId.get(id), b = n.absoluteBoundingBox, eff = n.effects.find((e: any) => vis(e) && e.type === 'BACKGROUND_BLUR');
      const sigma = eff.radius / 2, pad = Math.ceil(sigma * 3), S = SCALE;
      const W = Math.round((b.width + 2 * pad) * S), H = Math.round((b.height + 2 * pad) * S);
      const out = src.platePath(id, S);
      if (!existsSync(out)) {
        const bg = frame.fills?.find((f: any) => vis(f) && f.type === 'SOLID');
        const bgc = bg ? '#' + [bg.color.r, bg.color.g, bg.color.b].map((v: number) => Math.round(v * 255).toString(16).padStart(2, '0')).join('') : 'white';
        const layers = under.get(id)!.map(u => { const q = placeRaster(byId.get(u), rasterFile(u)); return {file: rasterFile(u), x: Math.round((q.x + F.x - b.x + pad) * S), y: Math.round((q.y + F.y - b.y + pad) * S)}; });
        env.images.blurPlate(out, W, H, bgc, layers, sigma * S, {w: Math.round(b.width * S), h: Math.round(b.height * S), x: pad * S, y: pad * S});
      }
      const rad = n.cornerRadius || 0, adj = rad ? Math.round(rad / Math.min(b.width, b.height) * 100000) : 0;
      pic(n.name + ' (blur plate)', out, b.x - F.x, b.y - F.y, b.width, b.height, ancOpacity(id, false), rad ? `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${adj}"/></a:avLst></a:prstGeom>` : undefined);
      stats.plates++;
    }

    // --- vectors ---
    function emitVector(el: any, m: Mat, st: any, nodeId: string) {
      const tag = el.tagName; let d: string;
      const A = (k: string, dflt = 0) => el.getAttribute(k) != null ? +el.getAttribute(k) : dflt;
      if (tag === 'path') d = el.getAttribute('d');
      else if (tag === 'rect') d = rectPath(A('x'), A('y'), A('width'), A('height'), A('rx', A('ry')), A('ry', A('rx')));
      else if (tag === 'circle') d = ellipsePath(A('cx'), A('cy'), A('r'), A('r'));
      else if (tag === 'ellipse') d = ellipsePath(A('cx'), A('cy'), A('rx'), A('ry'));
      else if (tag === 'line') d = `M${A('x1')} ${A('y1')}L${A('x2')} ${A('y2')}`;
      else { warn('unhandled svg element ' + tag); return; }
      const mm = mul(m, parseTransform(el.getAttribute('transform')));
      const segs = transformSegs(parsePath(d), mm);
      if (!segs.length) return;
      const bb = bbox(segs);
      const fillAttr = st.fill, strokeAttr = st.stroke;
      const op = st.opacity;
      const fill = paint(fillAttr, op * st.fillOpacity, bb, mm, defs, warn);
      let ln = '<a:ln><a:noFill/></a:ln>';
      if (strokeAttr && strokeAttr !== 'none') {
        const w = st.strokeWidth * scaleOf(mm);
        const cap = st.linecap === 'round' ? 'rnd' : st.linecap === 'square' ? 'sq' : 'flat';
        const join = st.linejoin === 'round' ? '<a:round/>' : st.linejoin === 'bevel' ? '<a:bevel/>' : '<a:miter lim="800000"/>';
        let dash = '';
        // DrawingML dash lengths are in percent of the line width.
        if (st.dash && st.dash !== 'none') { const v = st.dash.split(/[\s,]+/).map(Number).map((x: number) => x * scaleOf(mm)); const ds = []; for (let k = 0; k < v.length; k += 2) ds.push(`<a:ds d="${Math.round(v[k] / w * 100000)}" sp="${Math.round((v[k + 1] ?? v[k]) / w * 100000)}"/>`); dash = `<a:custDash>${ds.join('')}</a:custDash>`; }
        ln = `<a:ln w="${Math.round(w * E)}" cap="${cap}">${paint(strokeAttr, op * st.strokeOpacity, bb, mm, defs, warn)}${dash}${join}</a:ln>`;
      }
      const groups = fillAttr && fillAttr !== 'none' ? groupForFill(segs, st.fillRule === 'evenodd') : [segs];
      const n = byId.get(nodeId);
      body += `<p:sp><p:nvSpPr><p:cNvPr id="${sid++}" name="${esc(n?.name || tag)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(bb.x0 * E)}" y="${Math.round(bb.y0 * E)}"/><a:ext cx="${Math.max(1, Math.round((bb.x1 - bb.x0) * E))}" cy="${Math.max(1, Math.round((bb.y1 - bb.y0) * E))}"/></a:xfrm>${custGeom(groups, bb, E, !fillAttr || fillAttr === 'none')}${fill}${ln}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
      stats.shapes++;
    }

    // --- text ---
    function emitText(el: any, m: Mat, nodeId: string, opacity: number) {
      const t = textShape(el, m, byId.get(nodeId), nodeId, opacity, sid, {
        frame: F, corr: CORR[nodeId] || {dx: 0, dy: 0}, kern: KERN, fonts: env.fonts, warn,
        onFace: (key, face) => { report.fonts[key] = `${face.typeface}${face.b ? ' +b' : ''}${face.i ? ' +i' : ''}${face.note ? ' (' + face.note + ')' : ''}`; faces.set(key, face); },
      });
      if (!t) return;
      sid++;
      body += t.xml;
      report.textLines.push(...t.lines.map(r => ({...r, frame: frameId})));
      stats.text++;
    }

    // --- walk the SVG in paint order ---
    const inherit = (el: any, st: any) => {
      const g = (k: string) => el.getAttribute(k);
      return {
        fill: g('fill') ?? st.fill, stroke: g('stroke') ?? st.stroke, strokeWidth: g('stroke-width') != null ? +g('stroke-width') : st.strokeWidth,
        dash: g('stroke-dasharray') ?? st.dash, linecap: g('stroke-linecap') ?? st.linecap, linejoin: g('stroke-linejoin') ?? st.linejoin,
        fillRule: g('fill-rule') ?? st.fillRule, fillOpacity: g('fill-opacity') != null ? +g('fill-opacity') : st.fillOpacity, strokeOpacity: g('stroke-opacity') != null ? +g('stroke-opacity') : st.strokeOpacity,
        opacity: st.opacity * (g('opacity') != null ? +g('opacity') : 1),
      };
    };
    function walk(el: any, m: Mat, st: any, owner: string | null) {
      if (el.nodeType !== 1 || el.tagName === 'defs' || el.tagName === 'clipPath' || el.tagName === 'mask' || el.tagName === 'pattern') return;
      const id = el.getAttribute('data-node-id') || owner;
      if (id && id !== frameId) {
        const rr = ancestorsIn(id, raster);
        if (rr) { if (!done.has(rr)) { done.add(rr); emitRaster(rr); } return; }
        const bl = ancestorsIn(id, blurMap());
        if (bl && !done.has('plate:' + bl)) { done.add('plate:' + bl); emitPlate(bl); }
        if (byId.get(id)?.type === 'TEXT') { if (!done.has(id)) { done.add(id); emitText(el, mul(m, I), id, inherit(el, st).opacity); } return; }
      }
      if (el.tagName === 'foreignObject') return;
      const st2 = inherit(el, st);
      if (el.tagName === 'g' || el.tagName === 'svg') {
        const m2 = el.tagName === 'g' ? mul(m, parseTransform(el.getAttribute('transform'))) : m;
        for (const c of Array.from(el.childNodes) as any[]) walk(c, m2, st2, id);
        return;
      }
      if (el.tagName === 'image' || el.tagName === 'use') { warn(`image element outside raster in ${id}`); return; }
      if (el.tagName === 'text') { warn(`text element not owned by a TEXT node (${id})`); return; }
      emitVector(el, m, st2, id!);
    }
    walk(doc.documentElement, I, {fill: 'black', stroke: null, strokeWidth: 1, dash: null, linecap: null, linejoin: null, fillRule: 'nonzero', fillOpacity: 1, strokeOpacity: 1, opacity: 1}, null);
    report.frames.push({id: frameId, name: frame.name, ...stats});
    return {xml: slideXml(body, frame.name), rels};
  }

  const slides = FRAMES.map((f, i) => { opts.progress?.(`slide ${i + 1}/${FRAMES.length}`); return frameEmit(f); });
  const box0 = NODES[FRAMES[0]].document.absoluteBoundingBox;
  for (const f of FRAMES.slice(1)) { const b = NODES[f].document.absoluteBoundingBox; if (Math.round(b.width) !== Math.round(box0.width) || Math.round(b.height) !== Math.round(box0.height)) warn(`frame ${f} is ${b.width}x${b.height}, the slide size comes from the first frame (${box0.width}x${box0.height})`); }
  report.faces = [...faces].map(([figma, f]) => ({figma, typeface: f.typeface, b: f.b, i: f.i, status: f.status, note: f.note}));
  if (opts.embedFonts) opts.progress?.('embedding fonts');
  const embed = opts.embedFonts ? embedFaces([...faces.values()]) : {fonts: [], results: []};
  report.embeddedFonts = embed.results;
  const pptx = buildPptx(slides, media, Math.round(box0.width), Math.round(box0.height), 'figma2pptx', embed.fonts);
  report.rasters = Object.fromEntries(raster); report.blurPanels = [...blurNodes];
  return {pptx, report};
}
