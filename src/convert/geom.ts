// SVG path / transform parsing and DrawingML custGeom emission.
export type Mat = [number, number, number, number, number, number]; // a b c d e f (SVG order)
export const I: Mat = [1, 0, 0, 1, 0, 0];
export const mul = (m: Mat, n: Mat): Mat => [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
export const ap = (m: Mat, x: number, y: number): [number, number] => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
export const scaleOf = (m: Mat) => Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2]));

export function parseTransform(s: string | null): Mat {
  let m: Mat = I;
  if (!s) return m;
  for (const [, fn, args] of s.matchAll(/(\w+)\s*\(([^)]*)\)/g)) {
    const v = args.split(/[\s,]+/).filter(Boolean).map(Number);
    let t: Mat = I;
    if (fn === 'matrix') t = v as Mat;
    else if (fn === 'translate') t = [1, 0, 0, 1, v[0], v[1] || 0];
    else if (fn === 'scale') t = [v[0], 0, 0, v[1] ?? v[0], 0, 0];
    else if (fn === 'rotate') {
      const r = v[0] * Math.PI / 180, c = Math.cos(r), s2 = Math.sin(r);
      t = [c, s2, -s2, c, 0, 0];
      if (v.length > 2) t = mul(mul([1, 0, 0, 1, v[1], v[2]], t), [1, 0, 0, 1, -v[1], -v[2]]);
    }
    m = mul(m, t);
  }
  return m;
}

// Absolute segments: M, L, C (cubic), Z. Quadratics are elevated to cubics, arcs converted to cubics.
export type Seg = {t: 'M' | 'L' | 'C' | 'Z', p: number[]};
export function parsePath(d: string): Seg[] {
  const tok = d.match(/[a-zA-Z]|-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/g) || [];
  const out: Seg[] = [];
  let i = 0, cmd = '', x = 0, y = 0, sx = 0, sy = 0, lcx = 0, lcy = 0, lqx = 0, lqy = 0, prev = '';
  const num = () => +tok[i++];
  while (i < tok.length) {
    if (/[a-zA-Z]/.test(tok[i])) cmd = tok[i++];
    const rel = cmd === cmd.toLowerCase(), C = cmd.toUpperCase();
    const ox = rel ? x : 0, oy = rel ? y : 0;
    if (C === 'Z') { out.push({t: 'Z', p: []}); x = sx; y = sy; prev = 'Z'; continue; }
    if (C === 'M') { x = num() + ox; y = num() + oy; sx = x; sy = y; out.push({t: 'M', p: [x, y]}); cmd = rel ? 'l' : 'L'; prev = 'M'; continue; }
    if (C === 'L') { x = num() + ox; y = num() + oy; out.push({t: 'L', p: [x, y]}); }
    else if (C === 'H') { x = num() + (rel ? x : 0); out.push({t: 'L', p: [x, y]}); }
    else if (C === 'V') { y = num() + (rel ? y : 0); out.push({t: 'L', p: [x, y]}); }
    else if (C === 'C' || C === 'S') {
      let x1: number, y1: number;
      if (C === 'C') { x1 = num() + ox; y1 = num() + oy; } else { x1 = /[CS]/.test(prev) ? 2 * x - lcx : x; y1 = /[CS]/.test(prev) ? 2 * y - lcy : y; }
      const x2 = num() + ox, y2 = num() + oy, ex = num() + ox, ey = num() + oy;
      out.push({t: 'C', p: [x1, y1, x2, y2, ex, ey]}); lcx = x2; lcy = y2; x = ex; y = ey;
    } else if (C === 'Q' || C === 'T') {
      let qx: number, qy: number;
      if (C === 'Q') { qx = num() + ox; qy = num() + oy; } else { qx = /[QT]/.test(prev) ? 2 * x - lqx : x; qy = /[QT]/.test(prev) ? 2 * y - lqy : y; }
      const ex = num() + ox, ey = num() + oy;
      out.push({t: 'C', p: [x + 2 / 3 * (qx - x), y + 2 / 3 * (qy - y), ex + 2 / 3 * (qx - ex), ey + 2 / 3 * (qy - ey), ex, ey]});
      lqx = qx; lqy = qy; x = ex; y = ey;
    } else if (C === 'A') {
      const rx = num(), ry = num(), rot = num(), la = num(), sw = num(), ex = num() + ox, ey = num() + oy;
      for (const c of arcToCubic(x, y, rx, ry, rot, la, sw, ex, ey)) out.push({t: 'C', p: c});
      x = ex; y = ey;
    } else throw Error('path cmd ' + cmd);
    prev = C;
  }
  return out;
}

function arcToCubic(x1: number, y1: number, rx: number, ry: number, phi: number, fa: number, fs: number, x2: number, y2: number): number[][] {
  if (!rx || !ry) return [[x1, y1, x2, y2, x2, y2]];
  const r = phi * Math.PI / 180, cs = Math.cos(r), sn = Math.sin(r);
  const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2, x1p = cs * dx + sn * dy, y1p = -sn * dx + cs * dy;
  rx = Math.abs(rx); ry = Math.abs(ry);
  const lam = x1p * x1p / (rx * rx) + y1p * y1p / (ry * ry);
  if (lam > 1) { rx *= Math.sqrt(lam); ry *= Math.sqrt(lam); }
  const sgn = fa === fs ? -1 : 1;
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const co = sgn * Math.sqrt(Math.max(0, num / (rx * rx * y1p * y1p + ry * ry * x1p * x1p)));
  const cxp = co * rx * y1p / ry, cyp = -co * ry * x1p / rx;
  const cx = cs * cxp - sn * cyp + (x1 + x2) / 2, cy = sn * cxp + cs * cyp + (y1 + y2) / 2;
  const ang = (ux: number, uy: number, vx: number, vy: number) => Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
  let t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry), dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!fs && dt > 0) dt -= 2 * Math.PI; if (fs && dt < 0) dt += 2 * Math.PI;
  const n = Math.ceil(Math.abs(dt) / (Math.PI / 2)), d = dt / n, k = 4 / 3 * Math.tan(d / 4);
  const pt = (t: number) => [cx + rx * Math.cos(t) * cs - ry * Math.sin(t) * sn, cy + rx * Math.cos(t) * sn + ry * Math.sin(t) * cs];
  const dv = (t: number) => [-rx * Math.sin(t) * cs - ry * Math.cos(t) * sn, -rx * Math.sin(t) * sn + ry * Math.cos(t) * cs];
  const out: number[][] = [];
  for (let j = 0; j < n; j++) {
    const a = t1 + j * d, b = a + d, p0 = pt(a), p3 = pt(b), d0 = dv(a), d3 = dv(b);
    out.push([p0[0] + k * d0[0], p0[1] + k * d0[1], p3[0] - k * d3[0], p3[1] - k * d3[1], p3[0], p3[1]]);
  }
  return out;
}

export function rectPath(x: number, y: number, w: number, h: number, rx = 0, ry = rx): string {
  rx = Math.min(rx, w / 2); ry = Math.min(ry, h / 2);
  if (!rx || !ry) return `M${x} ${y}L${x + w} ${y}L${x + w} ${y + h}L${x} ${y + h}Z`;
  return `M${x + rx} ${y}L${x + w - rx} ${y}A${rx} ${ry} 0 0 1 ${x + w} ${y + ry}L${x + w} ${y + h - ry}A${rx} ${ry} 0 0 1 ${x + w - rx} ${y + h}L${x + rx} ${y + h}A${rx} ${ry} 0 0 1 ${x} ${y + h - ry}L${x} ${y + ry}A${rx} ${ry} 0 0 1 ${x + rx} ${y}Z`;
}
export const ellipsePath = (cx: number, cy: number, rx: number, ry: number) => `M${cx - rx} ${cy}A${rx} ${ry} 0 1 1 ${cx + rx} ${cy}A${rx} ${ry} 0 1 1 ${cx - rx} ${cy}Z`;

export function transformSegs(segs: Seg[], m: Mat): Seg[] {
  return segs.map(s => {
    const p: number[] = [];
    for (let k = 0; k < s.p.length; k += 2) p.push(...ap(m, s.p[k], s.p[k + 1]));
    return {t: s.t, p};
  });
}
export function bbox(segs: Seg[]) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const s of segs) for (let k = 0; k < s.p.length; k += 2) { x0 = Math.min(x0, s.p[k]); x1 = Math.max(x1, s.p[k]); y0 = Math.min(y0, s.p[k + 1]); y1 = Math.max(y1, s.p[k + 1]); }
  return {x0, y0, x1, y1};
}

// Split into subpaths, then group per non-zero winding: holes (contained, opposite orientation) join their container.
function subpaths(segs: Seg[]): Seg[][] {
  const out: Seg[][] = []; let cur: Seg[] = [];
  for (const s of segs) { if (s.t === 'M' && cur.length) { out.push(cur); cur = []; } cur.push(s); }
  if (cur.length) out.push(cur);
  return out.filter(sp => sp.length > 1);
}
function poly(sp: Seg[]): number[][] {
  const pts: number[][] = []; let lx = 0, ly = 0;
  for (const s of sp) {
    if (s.t === 'C') { for (let t = 0.25; t <= 1; t += 0.25) { const u = 1 - t; pts.push([u * u * u * lx + 3 * u * u * t * s.p[0] + 3 * u * t * t * s.p[2] + t * t * t * s.p[4], u * u * u * ly + 3 * u * u * t * s.p[1] + 3 * u * t * t * s.p[3] + t * t * t * s.p[5]]); } lx = s.p[4]; ly = s.p[5]; }
    else if (s.p.length) { pts.push([s.p[0], s.p[1]]); lx = s.p[0]; ly = s.p[1]; }
  }
  return pts;
}
const area = (p: number[][]) => { let a = 0; for (let i = 0; i < p.length; i++) { const q = p[(i + 1) % p.length]; a += p[i][0] * q[1] - q[0] * p[i][1]; } return a / 2; };
const inside = (x: number, y: number, p: number[][]) => { let c = false; for (let i = 0, j = p.length - 1; i < p.length; j = i++) if ((p[i][1] > y) !== (p[j][1] > y) && x < (p[j][0] - p[i][0]) * (y - p[i][1]) / (p[j][1] - p[i][1]) + p[i][0]) c = !c; return c; };
export function groupForFill(segs: Seg[], evenodd: boolean): Seg[][] {
  const sps = subpaths(segs);
  const seen = new Set<string>(); const uniq = sps.filter(sp => { const k = JSON.stringify(sp.map(s => s.p.map(v => v.toFixed(2)))); if (seen.has(k)) return false; seen.add(k); return true; });
  if (evenodd || uniq.length < 2) return [uniq.flat()];
  const P = uniq.map(poly), A = P.map(area);
  const parent = uniq.map((_, i) => {
    let best = -1;
    for (let j = 0; j < uniq.length; j++) if (j !== i && Math.abs(A[j]) > Math.abs(A[i]) && P[i].every(pt => inside(pt[0], pt[1], P[j])) && (best < 0 || Math.abs(A[j]) < Math.abs(A[best]))) best = j;
    return best;
  });
  const root = (i: number): number => parent[i] >= 0 && Math.sign(A[parent[i]]) !== Math.sign(A[i]) ? root(parent[i]) : i;
  const groups = new Map<number, Seg[]>();
  uniq.forEach((sp, i) => { const r = root(i); if (!groups.has(r)) groups.set(r, []); groups.get(r)!.push(...sp); });
  return [...groups.values()];
}

// custGeom in EMU relative to the bbox origin.
export function custGeom(groups: Seg[][], bb: {x0: number, y0: number, x1: number, y1: number}, E: number, fillNone: boolean) {
  const w = Math.max(1, Math.round((bb.x1 - bb.x0) * E)), h = Math.max(1, Math.round((bb.y1 - bb.y0) * E));
  const P = (x: number, y: number) => `<a:pt x="${Math.round((x - bb.x0) * E)}" y="${Math.round((y - bb.y0) * E)}"/>`;
  const paths = groups.map(g => `<a:path w="${w}" h="${h}"${fillNone ? ' fill="none"' : ''}>${g.map(s => s.t === 'M' ? `<a:moveTo>${P(s.p[0], s.p[1])}</a:moveTo>` : s.t === 'L' ? `<a:lnTo>${P(s.p[0], s.p[1])}</a:lnTo>` : s.t === 'C' ? `<a:cubicBezTo>${P(s.p[0], s.p[1])}${P(s.p[2], s.p[3])}${P(s.p[4], s.p[5])}</a:cubicBezTo>` : '<a:close/>').join('')}</a:path>`).join('');
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>`;
}
