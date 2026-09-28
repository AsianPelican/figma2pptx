// SVG paints to DrawingML fills: solid colours with alpha, and linear gradients remapped to a shape's box.
import {mul, ap, parseTransform, type Mat} from './geom';

export type Rgba = [string, number]; // RRGGBB, alpha 0..1
export type Box = {x0: number, y0: number, x1: number, y1: number};

export function hex(c: string | null, warn: (s: string) => void): Rgba | null {
  if (!c || c === 'none') return null;
  if (c === 'white') return ['FFFFFF', 1]; if (c === 'black') return ['000000', 1];
  let m = c.match(/^#([0-9a-f]{3})$/i); if (m) return [m[1].split('').map(x => x + x).join('').toUpperCase(), 1];
  m = c.match(/^#([0-9a-f]{6})$/i); if (m) return [m[1].toUpperCase(), 1];
  m = c.match(/rgba?\(([^)]*)\)/); if (m) { const v = m[1].split(',').map(Number); return [v.slice(0, 3).map(x => Math.round(x).toString(16).padStart(2, '0')).join('').toUpperCase(), v[3] ?? 1]; }
  warn('unknown color ' + c); return ['000000', 1];
}

export const clr = (h: string, a: number) => `<a:srgbClr val="${h}">${a < 0.9995 ? `<a:alpha val="${Math.round(a * 100000)}"/>` : ''}</a:srgbClr>`;

// A fill or stroke paint: `none`, a colour, or url(#id) of a linearGradient in `defs`. The gradient axis is
// transformed to slide space and the stops are clipped to the part of the axis the shape's box covers, because
// DrawingML stretches a linear gradient across the shape's box.
export function paint(attr: string | null, opacity: number, box: Box, m: Mat, defs: Map<string, any>, warn: (s: string) => void): string {
  if (!attr || attr === 'none') return '<a:noFill/>';
  const u = attr.match(/url\(#([^)]+)\)/);
  if (!u) { const c = hex(attr, warn)!; return `<a:solidFill>${clr(c[0], c[1] * opacity)}</a:solidFill>`; }
  const g = defs.get(u[1]);
  if (!g || g.tagName !== 'linearGradient') { warn('unsupported paint ' + (g?.tagName || u[1])); return '<a:noFill/>'; }
  const gm = mul(m, parseTransform(g.getAttribute('gradientTransform')));
  const [x1, y1] = ap(gm, +(g.getAttribute('x1') || 0), +(g.getAttribute('y1') || 0)), [x2, y2] = ap(gm, +(g.getAttribute('x2') || 1), +(g.getAttribute('y2') || 0));
  const dx = x2 - x1, dy = y2 - y1, L2 = dx * dx + dy * dy || 1;
  const stops = (Array.from(g.getElementsByTagName('stop')) as any[]).map(s => { const c = hex(s.getAttribute('stop-color') || '#000000', warn)!; return {o: +(s.getAttribute('offset') || 0), c: c[0], a: c[1] * +(s.getAttribute('stop-opacity') ?? 1)}; });
  const tOf = (x: number, y: number) => ((x - x1) * dx + (y - y1) * dy) / L2;
  const ts = [[box.x0, box.y0], [box.x1, box.y0], [box.x0, box.y1], [box.x1, box.y1]].map(p => tOf(p[0], p[1]));
  const tmin = Math.min(...ts), tmax = Math.max(...ts);
  const at = (t: number) => { if (t <= stops[0].o) return stops[0]; for (let k = 1; k < stops.length; k++) if (t <= stops[k].o) { const a = stops[k - 1], b = stops[k], f = (t - a.o) / (b.o - a.o || 1); const mix = (i: number) => Math.round(parseInt(a.c.slice(i, i + 2), 16) * (1 - f) + parseInt(b.c.slice(i, i + 2), 16) * f).toString(16).padStart(2, '0'); return {o: t, c: (mix(0) + mix(2) + mix(4)).toUpperCase(), a: a.a * (1 - f) + b.a * f}; } return stops[stops.length - 1]; };
  const gs = [{...at(tmin), o: tmin}, ...stops.filter(s => s.o > tmin && s.o < tmax), {...at(tmax), o: tmax}];
  const ang = Math.round(((Math.atan2(dy, dx) * 180 / Math.PI) + 360) % 360 * 60000);
  return `<a:gradFill rotWithShape="1"><a:gsLst>${gs.map(s => `<a:gs pos="${Math.round((s.o - tmin) / (tmax - tmin || 1) * 100000)}">${clr(s.c, s.a * opacity)}</a:gs>`).join('')}</a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill>`;
}
