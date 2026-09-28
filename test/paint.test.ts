import {test, expect} from 'bun:test';
import {DOMParser} from '@xmldom/xmldom';
import {hex, clr, paint} from '../src/convert/paint';
import {I, parseTransform} from '../src/convert/geom';

test('colours: names, #rgb, #rrggbb, rgb() and rgba()', () => {
  const warns: string[] = [];
  const w = (s: string) => warns.push(s);
  expect(hex('white', w)).toEqual(['FFFFFF', 1]);
  expect(hex('#1e3', w)).toEqual(['11EE33', 1]);
  expect(hex('#1E30F9', w)).toEqual(['1E30F9', 1]);
  expect(hex('rgba(255, 102, 0, 0.25)', w)).toEqual(['FF6600', 0.25]);
  expect(hex('none', w)).toBeNull();
  expect(hex('hsl(0,0%,0%)', w)).toEqual(['000000', 1]);
  expect(warns).toEqual(['unknown color hsl(0,0%,0%)']);
  expect(clr('000000', 1)).toBe('<a:srgbClr val="000000"></a:srgbClr>');
  expect(clr('000000', 0.5)).toBe('<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>');
});

test('a linear gradient is clipped to the part of its axis the shape covers', () => {
  const doc = new DOMParser().parseFromString('<svg xmlns="http://www.w3.org/2000/svg"><linearGradient id="g" x1="0" y1="0" x2="200" y2="0" gradientUnits="userSpaceOnUse"><stop stop-color="#000000"/><stop offset="1" stop-color="#FFFFFF"/></linearGradient></svg>', 'image/svg+xml') as any;
  const defs = new Map([['g', doc.getElementsByTagName('linearGradient')[0]]]);
  // the shape spans x 100..200: the second half of the axis, so from mid grey to white
  const xml = paint('url(#g)', 1, {x0: 100, y0: 0, x1: 200, y1: 10}, I, defs, () => {});
  expect(xml).toBe('<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="808080"></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"></a:srgbClr></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>');
});

test('the gradient axis follows the element transform (a vertical gradient is 90 degrees)', () => {
  const doc = new DOMParser().parseFromString('<svg xmlns="http://www.w3.org/2000/svg"><linearGradient id="g" x1="0" y1="0" x2="10" y2="0"><stop stop-color="#000000"/><stop offset="1" stop-color="#FFFFFF" stop-opacity="0.5"/></linearGradient></svg>', 'image/svg+xml') as any;
  const defs = new Map([['g', doc.getElementsByTagName('linearGradient')[0]]]);
  const xml = paint('url(#g)', 1, {x0: 0, y0: 0, x1: 1, y1: 10}, parseTransform('rotate(90)'), defs, () => {});
  expect(xml).toContain('<a:lin ang="5400000" scaled="0"/>');
  expect(xml).toContain('<a:alpha val="50000"/>');
});

test('unsupported paint servers are reported', () => {
  const warns: string[] = [];
  expect(paint('url(#nope)', 1, {x0: 0, y0: 0, x1: 1, y1: 1}, I, new Map(), s => warns.push(s))).toBe('<a:noFill/>');
  expect(warns).toEqual(['unsupported paint nope']);
});
