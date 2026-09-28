// SVG geometry -> DrawingML: path parsing, transforms, winding groups, custGeom emission.
import {test, expect, describe} from 'bun:test';
import {parsePath, parseTransform, transformSegs, bbox, groupForFill, custGeom, rectPath, ellipsePath, mul, ap, scaleOf, I} from '../src/convert/geom';

const close = (a: number[], b: number[], eps = 1e-6) => a.length === b.length && a.every((v, i) => Math.abs(v - b[i]) < eps);

describe('parsePath', () => {
  test('absolute and relative lines, H/V, implicit lineto after M, Z returns to the start', () => {
    const s = parsePath('M10 10 20 10h5v5H10l-2 -2zM0 0');
    expect(s.map(x => x.t).join('')).toBe('MLLLLLZM');
    expect(s[1].p).toEqual([20, 10]);
    expect(s[2].p).toEqual([25, 10]);
    expect(s[3].p).toEqual([25, 15]);
    expect(s[4].p).toEqual([10, 15]);
    expect(s[5].p).toEqual([8, 13]);
  });
  test('smooth cubic reflects the previous control point', () => {
    const s = parsePath('M0 0C0 10 10 10 10 0S20 -10 20 0');
    expect(s[2].p).toEqual([10, -10, 20, -10, 20, 0]);
  });
  test('quadratics are elevated to exact cubics', () => {
    const s = parsePath('M0 0Q10 10 20 0');
    expect(close(s[1].p, [20 / 3, 20 / 3, 40 / 3, 20 / 3, 20, 0])).toBe(true);
  });
  test('an arc becomes quarter-turn cubics that stay on the circle', () => {
    const s = parsePath(ellipsePath(0, 0, 10, 10));
    const cubics = s.filter(x => x.t === 'C');
    expect(cubics.length).toBe(4); // two half arcs, two cubics each
    for (const c of cubics) expect(Math.hypot(c.p[4], c.p[5])).toBeCloseTo(10, 9);
  });
  test('exponents and packed numbers tokenize', () => {
    expect(parsePath('M1e1-2.5L.5.5').map(x => x.p)).toEqual([[10, -2.5], [0.5, 0.5]]);
  });
});

describe('transforms', () => {
  test('translate, scale, rotate about a point, matrix, composition order', () => {
    expect(ap(parseTransform('translate(5 7)'), 1, 1)).toEqual([6, 8]);
    expect(ap(parseTransform('scale(2 3)'), 1, 1)).toEqual([2, 3]);
    const r = ap(parseTransform('rotate(90 10 10)'), 20, 10);
    expect(close(r, [10, 20])).toBe(true);
    expect(parseTransform('matrix(1 2 3 4 5 6)')).toEqual([1, 2, 3, 4, 5, 6]);
    const m = parseTransform('translate(10 0) scale(2)'); // scale first, then translate
    expect(ap(m, 1, 0)).toEqual([12, 0]);
    expect(mul(I, m)).toEqual(m);
  });
  test('scaleOf is the geometric mean scale (stroke widths)', () => {
    expect(scaleOf(parseTransform('scale(2 8)'))).toBeCloseTo(4, 12);
    expect(scaleOf(parseTransform('rotate(33)'))).toBeCloseTo(1, 12);
  });
  test('transformSegs and bbox', () => {
    const s = transformSegs(parsePath(rectPath(0, 0, 10, 20)), parseTransform('translate(5 5) rotate(90)'));
    const b = bbox(s);
    expect(close([b.x0, b.y0, b.x1, b.y1], [-15, 5, 5, 15])).toBe(true);
  });
});

describe('groupForFill (DrawingML fills even-odd per path; SVG defaults to non-zero)', () => {
  const ring = parsePath('M0 50A50 50 0 1 0 100 50A50 50 0 1 0 0 50ZM25 50A25 25 0 1 1 75 50A25 25 0 1 1 25 50Z');
  test('a hole with opposite winding joins its container: one path, even-odd shows the hole', () => {
    expect(groupForFill(ring, false).length).toBe(1);
  });
  test('a same-winding inner shape is an overlap, not a hole: separate paths, both filled', () => {
    const overlap = parsePath('M0 0H100V100H0ZM25 25H75V75H25Z');
    expect(groupForFill(overlap, false).length).toBe(2);
  });
  test('an island inside a hole stays in the same even-odd group, which fills it again', () => {
    const island = parsePath('M0 0H100V100H0ZM10 10V90H90V10ZM40 40H60V60H40Z');
    expect(groupForFill(island, false).length).toBe(1);
  });
  test('even-odd sources stay one path', () => {
    expect(groupForFill(parsePath('M0 0H100V100H0ZM25 25H75V75H25Z'), true).length).toBe(1);
  });
  test('duplicate sub-paths are removed', () => {
    const dup = parsePath('M0 0H10V10H0ZM0 0H10V10H0Z');
    expect(groupForFill(dup, false)[0].filter(s => s.t === 'M').length).toBe(1);
  });
});

test('custGeom: EMU relative to the box origin, fill="none" for stroke-only paths', () => {
  const segs = parsePath('M10 10L20 10L20 30Z');
  const xml = custGeom([segs], bbox(segs), 9525, true);
  expect(xml).toBe('<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="95250" h="190500" fill="none"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="95250" y="0"/></a:lnTo><a:lnTo><a:pt x="95250" y="190500"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>');
});

test('rectPath clamps the corner radius to half the side', () => {
  const b = bbox(parsePath(rectPath(0, 0, 10, 4, 50)));
  expect(close([b.x0, b.y0, b.x1, b.y1], [0, 0, 10, 4], 1e-9)).toBe(true);
});
