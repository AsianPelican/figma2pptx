// figma2pptx: Figma frames -> editable PPTX, 1:1 by construction.
//
//   bun figma2pptx.ts <fileKey> <frameId,frameId,...> <out.pptx> [--corr corrections.json] [--scale 2]
//
// Sources (all read-only Figma REST):
//   /v1/files/:key/nodes       the node tree: text styles (PostScript face, size, tracking, leading), effects, fills
//   /v1/images?format=svg      per frame, text kept live: exact line breaks and baselines (one <tspan> per laid-out
//                              line), and every vector's geometry, stroke, dash pattern and gradient in paint order
//   /v1/images?format=png|jpg  rasters for what PowerPoint cannot draw natively (image fills, effects, masks,
//                              angular gradients), and the plates behind background-blur panels
// Text becomes native PowerPoint text with Figma's own line breaks as hard breaks; vectors become custGeom shapes
// with native strokes, dashes, caps and alpha.
import {DOMParser} from '@xmldom/xmldom';
import {readFileSync, writeFileSync, existsSync, mkdirSync} from 'node:fs';
import {execFileSync} from 'node:child_process';
import {buildPptx, slideXml, esc, EMU_PER_PX as E, REL_IMAGE, type Media, type Slide} from './pptx';
import {parseTransform, parsePath, transformSegs, bbox, groupForFill, custGeom, rectPath, ellipsePath, mul, ap, scaleOf, I, type Mat} from './geom';
import {mapFace, spaceEm} from "./fonts";
import {figma} from './fetchsvg';

const args = process.argv.slice(2);
const opt = (k: string, d?: string) => { const i = args.indexOf(k); return i >= 0 ? args[i + 1] : d; };
const [KEY, IDS, OUT] = args;
const FRAMES = IDS.split(',');
const SCALE = +(opt('--scale', '2')!);
const CORR: Record<string, {dx: number, dy: number}> = opt('--corr') && existsSync(opt('--corr')!) ? JSON.parse(readFileSync(opt('--corr')!, 'utf8')) : {};
const KERN = opt('--kern', '100')!;
mkdirSync('cache/raster', {recursive: true});
const t0 = Date.now();
const report: any = {frames: [], warnings: [] as string[], fonts: {} as Record<string, string>, textLines: [] as any[]};
const warn = (s: string) => { if (!report.warnings.includes(s)) report.warnings.push(s); };

// ---------- 1. node tree ----------
const nodesPath = `cache/nodes-${KEY}-${FRAMES.length}.json`;
if (!existsSync(nodesPath)) writeFileSync(nodesPath, JSON.stringify(await figma(`files/${KEY}/nodes?ids=${IDS}`)));
const NODES = JSON.parse(readFileSync(nodesPath, 'utf8')).nodes;
const byId = new Map<string, any>(), parentOf = new Map<string, any>();
for (const f of FRAMES) (function idx(n: any, p: any) { byId.set(n.id, n); if (p) parentOf.set(n.id, p); for (const c of n.children || []) idx(c, n); })(NODES[f].document, null);

// ---------- 2. per-frame SVG ----------
if (FRAMES.some(f => !existsSync(`cache/svg/${f.replace(':', '-')}.svg`))) execFileSync('bun', ['fetchsvg.ts', KEY, IDS], {stdio: 'inherit'});

// ---------- 3. decide what must be raster ----------
const vis = (p: any) => p.visible !== false;
const count = (n: any): number => 1 + (n.children || []).reduce((a: number, c: any) => a + count(c), 0);
const hasText = (n: any): boolean => n.visible !== false && (n.type === 'TEXT' || (n.children || []).some(hasText));
function rasterReason(n: any, isRoot: boolean): string | null {
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
// SVG safety net: anything Figma had to express with clip-path, mask, filter or foreignObject is raster too.
const svgDocs = new Map<string, Document>();
for (const f of FRAMES) {
  const doc = new DOMParser().parseFromString(readFileSync(`cache/svg/${f.replace(':', '-')}.svg`, 'utf8'), 'image/svg+xml') as any;
  svgDocs.set(f, doc);
  const walk = (el: any, owner: string | null, depth: number) => {
    if (el.nodeType !== 1) return;
    const id = el.getAttribute('data-node-id') || owner;
    const inRaster = id && ancestorsIn(id, raster);
    const special = (depth > 1 && (el.getAttribute('clip-path') || el.getAttribute('mask'))) || el.getAttribute('filter') || (el.tagName === "foreignObject" && !nextEl(el)?.getAttribute("data-figma-bg-blur-radius"));
    if (special && id && !inRaster && !ancestorsIn(id, new Map([...blurNodes].map(b => [b, '']))) && byId.get(id)?.type !== 'TEXT' && id !== f) { raster.set(id, 'svg ' + (el.tagName === 'foreignObject' ? 'foreignObject' : el.getAttribute('filter') ? 'filter' : el.getAttribute('mask') ? 'mask' : 'clip-path')); return; }
    for (const c of Array.from(el.childNodes || []) as any[]) if (c.tagName !== 'defs') walk(c, id, depth + 1);
  };
  walk(doc.documentElement, null, 0);
}
function nextEl(el: any) { let s = el.nextSibling; while (s && s.nodeType !== 1) s = s.nextSibling; return s; }
function ancestorsIn(id: string, set: Map<string, string>): string | null {
  let n = byId.get(id);
  while (n) { if (set.has(n.id)) return n.id; n = parentOf.get(n.id); }
  return null;
}
for (const [id, why] of raster) if (hasText(byId.get(id))) warn(`raster ${id} "${byId.get(id).name}" (${why}) contains text, which is baked into the image`);

// ---------- 4. fetch rasters (cached) ----------
const frameOf = (id: string) => { let n = byId.get(id); while (parentOf.get(n.id)) n = parentOf.get(n.id); return n; };
const isOpaqueRect = (n: any) => n.type === 'RECTANGLE' && !n.cornerRadius && !n.rectangleCornerRadii && (n.opacity ?? 1) === 1 && !(n.rotation && Math.abs(n.rotation) > 1e-3) && !(n.strokes || []).some(vis) && (n.fills || []).some((f: any) => vis(f) && f.type === 'IMAGE' && (f.opacity ?? 1) === 1);
async function fetchRasters(ids: string[], scale: number) {
  const need = ids.filter(id => !existsSync(`cache/raster/${id.replace(':', '-')}@${scale}.png`) && !existsSync(`cache/raster/${id.replace(':', '-')}@${scale}.jpg`));
  for (const fmt of ['png', 'jpg']) {
    const list = fmt === "png" ? need : []; // always PNG: an image fill can carry alpha; opaque results become JPEG locally
    for (let i = 0; i < list.length; i += 10) {
      const chunk = list.slice(i, i + 10);
      const j: any = await figma(`images/${KEY}?ids=${chunk.join(",")}&format=${fmt}&scale=${scale}&use_absolute_bounds=true`);
      await Promise.all(chunk.map(async id => { const b = new Uint8Array(await (await fetch(j.images[id])).arrayBuffer()); writeFileSync(`cache/raster/${id.replace(':', '-')}@${scale}.${fmt}`, b); }));
    }
  }
}
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
function paintIndex(id: string) { const order: string[] = []; (function w(n: any) { order.push(n.id); for (const c of n.children || []) w(c); })(frameOf(id)); return order.indexOf(id); }
await fetchRasters([...raster.keys()], SCALE);
await fetchRasters([...new Set([...under.values()].flat())].filter(id => !raster.has(id)), SCALE);
await fetchRasters([...new Set([...under.values()].flat())].filter(id => raster.has(id)), SCALE);
const rasterFile = (id: string) => { const p = `cache/raster/${id.replace(':', '-')}@${SCALE}`; return existsSync(p + '.jpg') ? p + '.jpg' : p + '.png'; };

// ---------- 5. emit ----------
const media: Media[] = [];
const mediaByPath = new Map<string, string>();
function addMedia(path: string): string {
  if (mediaByPath.has(path)) return mediaByPath.get(path)!;
  const ext = path.split('.').pop()!, name = `image${media.length + 1}.${ext}`;
  media.push({name, data: readFileSync(path), ct: ext === 'jpg' ? 'image/jpeg' : 'image/png'});
  mediaByPath.set(path, name); return name;
}
const hex = (c: string | null): [string, number] | null => {
  if (!c || c === 'none') return null;
  if (c === 'white') return ['FFFFFF', 1]; if (c === 'black') return ['000000', 1];
  let m = c.match(/^#([0-9a-f]{3})$/i); if (m) return [m[1].split('').map(x => x + x).join('').toUpperCase(), 1];
  m = c.match(/^#([0-9a-f]{6})$/i); if (m) return [m[1].toUpperCase(), 1];
  m = c.match(/rgba?\(([^)]*)\)/); if (m) { const v = m[1].split(',').map(Number); return [v.slice(0, 3).map(x => Math.round(x).toString(16).padStart(2, '0')).join('').toUpperCase(), v[3] ?? 1]; }
  warn('unknown color ' + c); return ['000000', 1];
};
const clr = (h: string, a: number) => `<a:srgbClr val="${h}">${a < 0.9995 ? `<a:alpha val="${Math.round(a * 100000)}"/>` : ''}</a:srgbClr>`;

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
  // A node render covers the node unclipped (bounding box, grown by effects such as shadows); the REST render bounds
  // are clipped by the frame, so pick whichever box matches the rendered pixel size.
  const placeRaster = (n: any, file: string) => {
    const [pw, ph] = execFileSync("magick", ["identify", "-format", "%w %h", file], {encoding: "utf8"}).split(" ").map(Number);
    const w = pw / SCALE, h = ph / SCALE;
    const boxes = [n.absoluteBoundingBox, n.absoluteRenderBounds].filter(Boolean);
    const b = boxes.sort((a: any, c: any) => (Math.abs(a.width - w) + Math.abs(a.height - h)) - (Math.abs(c.width - w) + Math.abs(c.height - h)))[0];
    if (Math.abs(b.width - w) + Math.abs(b.height - h) > 2) warn(`raster ${n.id} "${n.name}": ${w}x${h} matches neither bounding nor render box; placed at the bounding box`);
    return {x: b.x - F.x, y: b.y - F.y, w, h};
  };
  const rb = (n: any) => { const r = n.absoluteRenderBounds || n.absoluteBoundingBox; return {x: r.x - F.x, y: r.y - F.y, w: r.width, h: r.height}; };
  function emitRaster(id: string) {
    const n = byId.get(id), r = placeRaster(n, rasterFile(id));
    // Trim what lies off the slide and re-encode Figma's JPEGs at q85: same pixels on the slide, a fraction of the bytes.
    const x0 = Math.max(0, r.x), y0 = Math.max(0, r.y), x1 = Math.min(F.width, r.x + r.w), y1 = Math.min(F.height, r.y + r.h);
    let file = rasterFile(id);
    const trim = x0 > r.x + 0.5 || y0 > r.y + 0.5 || x1 < r.x + r.w - 0.5 || y1 < r.y + r.h - 0.5;
    // opaque apart from anti-aliased edge pixels -> JPEG
    const opaque = +execFileSync("magick", [file, "-alpha", "extract", "-shave", "3x3", "-format", "%[fx:minima]", "info:"], {encoding: "utf8"}).trim() >= 0.999;
    if (opaque || trim) {
      const out = file.replace(/\.(jpg|png)$/, opaque ? ".slide.jpg" : ".slide.png");
      if (!existsSync(out)) execFileSync("magick", [file, ...(trim ? ["-crop", `${Math.round((x1 - x0) * SCALE)}x${Math.round((y1 - y0) * SCALE)}+${Math.round((x0 - r.x) * SCALE)}+${Math.round((y0 - r.y) * SCALE)}`, "+repage"] : []), ...(opaque ? ["-background", "white", "-flatten", "-quality", "85"] : []), out]);
      file = out;
    }
    if (trim) Object.assign(r, {x: x0, y: y0, w: x1 - x0, h: y1 - y0});
    pic(n.name, file, r.x, r.y, r.w, r.h, ancOpacity(id, false));
    stats.rasters++; if (hasText(n)) stats.rasterText++;
  }
  function emitPlate(id: string) {
    // Background blur: composite what Figma paints beneath the panel, blur it as Figma does, clip to the panel.
    const n = byId.get(id), b = n.absoluteBoundingBox, eff = n.effects.find((e: any) => vis(e) && e.type === 'BACKGROUND_BLUR');
    const sigma = eff.radius / 2, pad = Math.ceil(sigma * 3), S = SCALE;
    const W = Math.round((b.width + 2 * pad) * S), H = Math.round((b.height + 2 * pad) * S);
    const out = `cache/raster/plate-${id.replace(':', '-')}@${S}.png`;
    if (!existsSync(out)) {
      const bg = frame.fills?.find((f: any) => vis(f) && f.type === 'SOLID');
      const bgc = bg ? '#' + [bg.color.r, bg.color.g, bg.color.b].map((v: number) => Math.round(v * 255).toString(16).padStart(2, '0')).join('') : 'white';
      const cmd = ['-size', `${W}x${H}`, `xc:${bgc}`];
      for (const u of under.get(id)!) { const un = byId.get(u), q = placeRaster(un, rasterFile(u)), r = {x: q.x + F.x, y: q.y + F.y, width: q.w, height: q.h}; cmd.push("(", rasterFile(u), ")", "-geometry", `+${Math.round((r.x - b.x + pad) * S)}+${Math.round((r.y - b.y + pad) * S)}`, '-composite'); }
      cmd.push('-blur', `0x${sigma * S}`, '-crop', `${Math.round(b.width * S)}x${Math.round(b.height * S)}+${pad * S}+${pad * S}`, '+repage', out);
      execFileSync('magick', cmd);
    }
    const rad = n.cornerRadius || 0, adj = rad ? Math.round(rad / Math.min(b.width, b.height) * 100000) : 0;
    pic(n.name + ' (blur plate)', out, b.x - F.x, b.y - F.y, b.width, b.height, ancOpacity(id, false), rad ? `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${adj}"/></a:avLst></a:prstGeom>` : undefined);
    stats.plates++;
  }

  // --- vectors ---
  function paint(attr: string | null, opacity: number, segsBB: any, m: Mat): string {
    if (!attr || attr === 'none') return '<a:noFill/>';
    const u = attr.match(/url\(#([^)]+)\)/);
    if (!u) { const c = hex(attr)!; return `<a:solidFill>${clr(c[0], c[1] * opacity)}</a:solidFill>`; }
    const g = defs.get(u[1]);
    if (!g || g.tagName !== 'linearGradient') { warn('unsupported paint ' + (g?.tagName || u[1])); return '<a:noFill/>'; }
    const gm = mul(m, parseTransform(g.getAttribute('gradientTransform')));
    const [x1, y1] = ap(gm, +(g.getAttribute('x1') || 0), +(g.getAttribute('y1') || 0)), [x2, y2] = ap(gm, +(g.getAttribute('x2') || 1), +(g.getAttribute('y2') || 0));
    const dx = x2 - x1, dy = y2 - y1, L2 = dx * dx + dy * dy || 1;
    const stops = (Array.from(g.getElementsByTagName('stop')) as any[]).map(s => { const c = hex(s.getAttribute('stop-color') || '#000000')!; return {o: +(s.getAttribute('offset') || 0), c: c[0], a: c[1] * +(s.getAttribute('stop-opacity') ?? 1)}; });
    const tOf = (x: number, y: number) => ((x - x1) * dx + (y - y1) * dy) / L2;
    const ts = [[segsBB.x0, segsBB.y0], [segsBB.x1, segsBB.y0], [segsBB.x0, segsBB.y1], [segsBB.x1, segsBB.y1]].map(p => tOf(p[0], p[1]));
    const tmin = Math.min(...ts), tmax = Math.max(...ts);
    const at = (t: number) => { if (t <= stops[0].o) return stops[0]; for (let k = 1; k < stops.length; k++) if (t <= stops[k].o) { const a = stops[k - 1], b = stops[k], f = (t - a.o) / (b.o - a.o || 1); const mix = (i: number) => Math.round(parseInt(a.c.slice(i, i + 2), 16) * (1 - f) + parseInt(b.c.slice(i, i + 2), 16) * f).toString(16).padStart(2, '0'); return {o: t, c: (mix(0) + mix(2) + mix(4)).toUpperCase(), a: a.a * (1 - f) + b.a * f}; } return stops[stops.length - 1]; };
    const gs = [{...at(tmin), o: tmin}, ...stops.filter(s => s.o > tmin && s.o < tmax), {...at(tmax), o: tmax}];
    const ang = Math.round(((Math.atan2(dy, dx) * 180 / Math.PI) + 360) % 360 * 60000);
    return `<a:gradFill rotWithShape="1"><a:gsLst>${gs.map(s => `<a:gs pos="${Math.round((s.o - tmin) / (tmax - tmin || 1) * 100000)}">${clr(s.c, s.a * opacity)}</a:gs>`).join('')}</a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill>`;
  }
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
    const fill = paint(fillAttr, op * st.fillOpacity, bb, mm);
    let ln = '<a:ln><a:noFill/></a:ln>';
    if (strokeAttr && strokeAttr !== 'none') {
      const w = st.strokeWidth * scaleOf(mm);
      const cap = st.linecap === 'round' ? 'rnd' : st.linecap === 'square' ? 'sq' : 'flat';
      const join = st.linejoin === 'round' ? '<a:round/>' : st.linejoin === 'bevel' ? '<a:bevel/>' : '<a:miter lim="800000"/>';
      let dash = '';
      if (st.dash && st.dash !== 'none') { const v = st.dash.split(/[\s,]+/).map(Number).map((x: number) => x * scaleOf(mm)); const ds = []; for (let k = 0; k < v.length; k += 2) ds.push(`<a:ds d="${Math.round(v[k] / w * 100000)}" sp="${Math.round((v[k + 1] ?? v[k]) / w * 100000)}"/>`); dash = `<a:custDash>${ds.join('')}</a:custDash>`; }
      ln = `<a:ln w="${Math.round(w * E)}" cap="${cap}">${paint(strokeAttr, op * st.strokeOpacity, bb, mm)}${dash}${join}</a:ln>`;
    }
    const groups = fillAttr && fillAttr !== 'none' ? groupForFill(segs, st.fillRule === 'evenodd') : [segs];
    const n = byId.get(nodeId);
    body += `<p:sp><p:nvSpPr><p:cNvPr id="${sid++}" name="${esc(n?.name || tag)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(bb.x0 * E)}" y="${Math.round(bb.y0 * E)}"/><a:ext cx="${Math.max(1, Math.round((bb.x1 - bb.x0) * E))}" cy="${Math.max(1, Math.round((bb.y1 - bb.y0) * E))}"/></a:xfrm>${custGeom(groups, bb, E, !fillAttr || fillAttr === 'none')}${fill}${ln}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
    stats.shapes++;
  }

  // --- text ---
  function emitText(el: any, m: Mat, nodeId: string, opacity: number) {
    const n = byId.get(nodeId);
    const texts = (el.tagName === 'text' ? [el] : Array.from(el.getElementsByTagName('text'))) as any[];
    type Seg = {x: number, y: number, t: string, fill: [string, number] | null, deco: string | null};
    const segs: Seg[] = [];
    let rotTm: Mat | null = null; // rotated text: lay out in the text's own frame, then rotate the box
    for (const t of texts) {
      const tm = mul(m, parseTransform(t.getAttribute('transform')));
      if (Math.abs(tm[1]) > 1e-3 || Math.abs(tm[2]) > 1e-3) rotTm = tm;
      const fillA = t.getAttribute('fill'), fo = +(t.getAttribute('fill-opacity') ?? 1);
      const f = hex(fillA); if (f) f[1] *= fo;
      for (const sp of Array.from(t.getElementsByTagName('tspan')) as any[]) {
        const [x, y] = rotTm ? [+sp.getAttribute("x"), +sp.getAttribute("y")] : ap(tm, +sp.getAttribute("x"), +sp.getAttribute("y"));
        segs.push({x, y, t: sp.textContent || '', fill: f, deco: t.getAttribute('text-decoration')});
      }
    }
    if (!segs.length || !n.characters.trim()) return;
    // lines by baseline
    const lines: Seg[][] = [];
    for (const s of [...segs].sort((a, b) => a.y - b.y || a.x - b.x)) { const l = lines.find(l => Math.abs(l[0].y - s.y) < 0.01); if (l) l.push(s); else lines.push([s]); }
    lines.sort((a, b) => a[0].y - b[0].y);
    // walk node.characters to recover per-char style + paragraph boundaries
    const chars = n.characters as string, ov = n.characterStyleOverrides || [], tbl = n.styleOverrideTable || {};
    let p = 0;
    type Run = {t: string, style: any, fill: [string, number] | null, deco: string | null};
    type Line = {runs: Run[], x: number, y: number, hardBefore: boolean};
    const L: Line[] = [];
    for (const segsOfLine of lines) {
      let hardBefore = false; const runs: Run[] = [];
      segsOfLine.forEach((s, si) => {
        let txt = s.t.replace(/\n$/, '').replace(/\u2028$/, '');
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
        if (p < chars.length && chars[p] === '\n' && si === segsOfLine.length - 1) { /* consumed at next line */ }
      });
      L.push({runs, x: segsOfLine[0].x, y: segsOfLine[0].y, hardBefore});
    }
    // paragraphs
    const paras: Line[][] = [];
    for (const l of L) { if (!paras.length || l.hardBefore) paras.push([l]); else paras[paras.length - 1].push(l); }
    const nb = n.absoluteBoundingBox, bx = rotTm ? 0 : nb.x - F.x, by = rotTm ? 0 : nb.y - F.y;
    const algn = ({LEFT: 'l', CENTER: 'ctr', RIGHT: 'r', JUSTIFIED: 'just'} as any)[n.style.textAlignHorizontal] || 'l';
    const bullets = (n.lineTypes || []) as string[];
    const corr = CORR[nodeId] || {dx: 0, dy: 0};
    const maxSize = (l: Line) => l.runs.length ? Math.max(...l.runs.map(r => r.style.fontSize)) : n.style.fontSize;
    let xml = '', pi = 0;
    const lineRec: any[] = [];
    for (const para of paras) {
      // pitch Figma used inside this paragraph; single-line paragraphs use the style's line height
      const pitches = para.slice(1).map((l, k) => l.y - para[k].y);
      const pitch = pitches.length ? pitches.sort((a, b) => a - b)[Math.floor(pitches.length / 2)] : (para[0].runs[0].style.lineHeightPx || maxSize(para[0]) * 1.2);
      const S = Math.max(...para.map(maxSize));
      const pts = Math.round(pitch * 0.75), errPts = Math.abs(pts / 0.75 - pitch);
      const pct = Math.round(pitch / (1.2 * S) * 100), errPct = Math.abs(pct / 100 * 1.2 * S - pitch);
      const lnSpc = errPct < errPts ? `<a:spcPct val="${pct * 1000}"/>` : `<a:spcPts val="${pts * 100}"/>`;
      // gap before this paragraph beyond one pitch (blank lines, paragraph spacing)
      const isBullet = bullets[pi] && bullets[pi] !== 'NONE';
      const marL = Math.max(0, para[0].x - bx);
      const indentXml = algn === 'l' && marL > 0.05 ? ` marL="${Math.round(marL * E)}"${isBullet ? ` indent="${-Math.round(marL * E)}"` : ''}` : '';
      xml += `<a:p><a:pPr algn="${algn}"${indentXml}><a:lnSpc>${lnSpc}</a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft>${isBullet ? '<a:buFont typeface="Arial"/><a:buChar char="•"/>' : '<a:buNone/>'}</a:pPr>`;
      para.forEach((l, li) => {
        if (li) xml += `<a:br><a:rPr lang="en-US" sz="${Math.round(maxSize(l) * 75)}"/></a:br>`;
        // Centred and right-aligned lines: PowerPoint counts a soft-wrap line's trailing space in the line width, Figma does
        // not. Keep the space (the copy is unchanged) but give it zero advance so both centre the same glyphs.
        const runs: any[] = [...l.runs];
        const lastR = runs[runs.length - 1];
        if ((algn === 'ctr' || algn === 'r') && li < para.length - 1 && lastR && /\s+$/.test(lastR.t) && lastR.t.trim()) {
          const tr = lastR.t.match(/\s+$/)![0];
          runs[runs.length - 1] = {...lastR, t: lastR.t.slice(0, -tr.length)};
          const fc = mapFace(lastR.style.fontPostScriptName, lastR.style.fontFamily, lastR.style.fontWeight, !!lastR.style.italic);
          runs.push({...lastR, t: tr, spcOverride: -spaceEm(fc.ps) * lastR.style.fontSize});
        }
        for (const r of runs) {
          const s = r.style, face = mapFace(s.fontPostScriptName, s.fontFamily, s.fontWeight, !!s.italic);
          report.fonts[`${s.fontPostScriptName || s.fontFamily + ' ' + s.fontWeight}`] = `${face.typeface}${face.b ? ' +b' : ''}${face.i ? ' +i' : ''}${face.note ? ' (' + face.note + ')' : ''}`;
          const fl = r.fill || ['000000', 1];
          xml += `<a:r><a:rPr lang="en-US" sz="${Math.round(s.fontSize * 75)}" b="${face.b}" i="${face.i}"${r.deco === 'underline' || s.textDecoration === 'UNDERLINE' ? ' u="sng"' : ''} spc="${Math.round((r.spcOverride ?? s.letterSpacing ?? 0) * 75)}" kern="${KERN}" dirty="0"><a:solidFill>${clr(fl[0], fl[1] * opacity)}</a:solidFill><a:latin typeface="${esc(face.typeface)}"/><a:ea typeface="${esc(face.typeface)}"/><a:cs typeface="${esc(face.typeface)}"/></a:rPr><a:t>${esc(r.t)}</a:t></a:r>`;
        }
        lineRec.push({node: nodeId, t: l.runs.map(r => r.t).join(''), x: l.x, base: l.y, align: algn});
      });
      xml += `<a:endParaRPr lang="en-US" sz="${Math.round(S * 75)}"/></a:p>`;
      pi += 1 + 0; // lineTypes are per source line (paragraph) in Figma
    }
    report.textLines.push(...lineRec.map(r => ({...r, frame: frameId})));
    // rotated: recover the unrotated box size from its axis-aligned bounds (W = w|cos|+h|sin|, H = w|sin|+h|cos|)
    let w = nb.width, h = nb.height;
    if (rotTm) { const t = Math.atan2(rotTm[1], rotTm[0]), c = Math.abs(Math.cos(t)), s = Math.abs(Math.sin(t)), det = c * c - s * s; if (Math.abs(det) > 1e-3) { w = (nb.width * c - nb.height * s) / det; h = (nb.height * c - nb.width * s) / det; } }
    let ox = bx + corr.dx, oy = by + corr.dy, rotAttr = "";
    if (rotTm) { const c = ap(rotTm, w / 2, h / 2); ox = c[0] - w / 2; oy = c[1] - h / 2; rotAttr = ` rot="${Math.round(((Math.atan2(rotTm[1], rotTm[0]) * 180 / Math.PI) + 360) % 360 * 60000)}"`; }
    body += `<p:sp><p:nvSpPr><p:cNvPr id="${sid++}" name="${esc(n.name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${rotAttr}><a:off x="${Math.round(ox * E)}" y="${Math.round(oy * E)}"/><a:ext cx="${Math.max(1, Math.round(w * E))}" cy="${Math.max(1, Math.round(h * E))}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>${xml}</p:txBody></p:sp>`;
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
      const bl = ancestorsIn(id, new Map([...blurNodes].map(b => [b, ''])));
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
  // Slide background from the frame fill, then content.
  const root = doc.documentElement;
  walk(root, I, {fill: 'black', stroke: null, strokeWidth: 1, dash: null, linecap: null, linejoin: null, fillRule: 'nonzero', fillOpacity: 1, strokeOpacity: 1, opacity: 1}, null);
  report.frames.push({id: frameId, name: frame.name, ...stats});
  return {xml: slideXml(body, frame.name), rels};
}

const slides = FRAMES.map(frameEmit);
const W = Math.round(NODES[FRAMES[0]].document.absoluteBoundingBox.width), H = Math.round(NODES[FRAMES[0]].document.absoluteBoundingBox.height);
writeFileSync(OUT, buildPptx(slides, media, W, H, 'figma2pptx'));
report.rasters = Object.fromEntries(raster); report.blurPanels = [...blurNodes];
report.seconds = +((Date.now() - t0) / 1000).toFixed(1);
writeFileSync(OUT.replace(/\.pptx$/, '.report.json'), JSON.stringify(report, null, 1));
console.log(`${OUT}: ${slides.length} slides, ${media.length} media, ${report.seconds}s; ${report.warnings.length} warnings`);
for (const f of report.frames) console.log(' ', f.name, JSON.stringify(f));
for (const w of report.warnings) console.log('  WARN', w);
