// Font embedding: reading the sfnt tables PowerPoint needs, the EOT wrapper, and the licence check.
import {test, expect} from 'bun:test';
import {mkdtempSync, writeFileSync, rmSync} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {readSfnt, toEot, embeddingAllowed} from '../src/convert/sfnt';
import {embedFaces} from '../src/convert/embed';
import {buildPptx} from '../src/convert/pptx';
import {unzipSync, strFromU8} from 'fflate';

// A minimal synthetic font: head, name (Windows Unicode records), OS/2 and a stand-in outline table.
function font(opts: {family: string, style: string, weight: number, fsType: number, italic?: boolean, extra?: string[]}): Uint8Array {
  const be = (n: number, len: number) => { const b = new Uint8Array(len); for (let i = 0; i < len; i++) b[len - 1 - i] = (n / 2 ** (8 * i)) & 255; return b; };
  const cat = (...a: Uint8Array[]) => { const o = new Uint8Array(a.reduce((s, x) => s + x.length, 0)); let k = 0; for (const x of a) { o.set(x, k); k += x.length; } return o; };
  const utf16be = (s: string) => cat(...[...s].map(c => be(c.charCodeAt(0), 2)));
  const names: [number, string][] = [[1, opts.family], [2, opts.style], [4, `${opts.family} ${opts.style}`], [5, 'Version 1.0']];
  const strs = names.map(([, s]) => utf16be(s));
  let off = 0;
  const recs = names.map(([id], k) => { const r = cat(be(3, 2), be(1, 2), be(0x409, 2), be(id, 2), be(strs[k].length, 2), be(off, 2)); off += strs[k].length; return r; });
  const name = cat(be(0, 2), be(names.length, 2), be(6 + 12 * names.length, 2), ...recs, ...strs);
  const head = new Uint8Array(54); head.set(be(0xdeadbeef, 4), 8);
  const os2 = new Uint8Array(96);
  os2.set(be(4, 2), 0); os2.set(be(opts.weight, 2), 4); os2.set(be(opts.fsType, 2), 8);
  os2.set([2, 11, 5, 3, 3, 4, 3, 2, 2, 4], 32); os2.set(be(0x80000003, 4), 42); os2.set(be(opts.italic ? 1 : 0x40, 2), 62); os2.set(be(1, 4), 78);
  const tables: [string, Uint8Array][] = [['OS/2', os2], ['head', head], ['name', name], ...(opts.extra ?? ['glyf']).map(t => [t, new Uint8Array(8)] as [string, Uint8Array])];
  let pos = 12 + 16 * tables.length;
  const dir = tables.map(([tag, d]) => { const e = cat(new TextEncoder().encode(tag.padEnd(4)), be(0, 4), be(pos, 4), be(d.length, 4)); pos += (d.length + 3) & ~3; return e; });
  const body = tables.map(([, d]) => cat(d, new Uint8Array(((d.length + 3) & ~3) - d.length)));
  return cat(be(0x00010000, 4), be(tables.length, 2), be(0, 6), ...dir, ...body);
}

test('readSfnt: names, weight, italic, licence flags, outline kind, variable', () => {
  const f = readSfnt(font({family: 'Test SemiBold', style: 'Italic', weight: 600, fsType: 8, italic: true}));
  expect(f.names[1]).toBe('Test SemiBold');
  expect(f.names[2]).toBe('Italic');
  expect([f.weight, f.fsType, f.italic, f.outlines, f.variable]).toEqual([600, 8, true, 'truetype', false]);
  expect(f.checkSumAdjustment).toBe(0xdeadbeef);
  expect(readSfnt(font({family: 'V', style: 'Regular', weight: 400, fsType: 0, extra: ['CFF ', 'fvar']}))).toMatchObject({outlines: 'cff', variable: true});
});

test('toEot: EOT 2.1 header, UTF-16LE names, the font data unchanged at the end', () => {
  const data = font({family: 'Test', style: 'Bold', weight: 700, fsType: 0});
  const eot = toEot(readSfnt(data)), dv = new DataView(eot.buffer, eot.byteOffset);
  expect(dv.getUint32(0, true)).toBe(eot.length);
  expect(dv.getUint32(4, true)).toBe(data.length);
  expect(dv.getUint32(8, true)).toBe(0x00020001);
  expect(dv.getUint32(12, true)).toBe(0); // not compressed, not encrypted
  expect(dv.getUint16(34, true)).toBe(0x504c);
  expect(dv.getUint32(28, true)).toBe(700);
  expect(dv.getUint32(60, true)).toBe(0xdeadbeef);
  expect(dv.getUint16(82, true)).toBe(8); // "Test" in UTF-16
  expect(new TextDecoder('utf-16le').decode(eot.subarray(84, 92))).toBe('Test');
  expect(Buffer.from(eot.subarray(eot.length - data.length)).equals(Buffer.from(data))).toBe(true);
});

test('embeddingAllowed follows the OS/2 fsType licence bits', () => {
  expect([0, 4, 8, 0x100].map(embeddingAllowed)).toEqual([true, true, true, true]);
  expect([2, 0x200].map(embeddingAllowed)).toEqual([false, false]);
});

test('embedFaces: one EOT per (typeface, slot); restricted, variable-only and unknown faces are reported', () => {
  const dir = mkdtempSync(join(tmpdir(), 'figma2pptx-fonts-'));
  try {
    const ok = join(dir, 'ok.ttf'), bold = join(dir, 'bold.ttf'), restricted = join(dir, 'r.ttf');
    writeFileSync(ok, font({family: 'Test SemiBold', style: 'Regular', weight: 600, fsType: 0}));
    writeFileSync(bold, font({family: 'Test', style: 'Bold', weight: 700, fsType: 8}));
    writeFileSync(restricted, font({family: 'Locked', style: 'Regular', weight: 400, fsType: 2}));
    const {fonts, results} = embedFaces([
      {typeface: 'Test SemiBold', b: 0, i: 0, ps: 'T-SB', status: 'exact', file: ok},
      {typeface: 'Test SemiBold', b: 0, i: 0, ps: 'T-SB', status: 'exact', file: ok},
      {typeface: 'Test', b: 1, i: 0, ps: 'T-B', status: 'exact', file: bold},
      {typeface: 'Locked', b: 0, i: 0, ps: 'L', status: 'exact', file: restricted},
      {typeface: 'Varia', b: 0, i: 0, ps: 'V', status: 'variable-only'},
    ]);
    expect(fonts.map(f => [f.typeface, Object.keys(f).filter(k => k !== 'typeface')])).toEqual([['Test SemiBold', ['regular']], ['Test', ['bold']]]);
    expect(results.map(r => [r.typeface, r.slot, r.embedded, r.reason])).toEqual([
      ['Test SemiBold', 'regular', true, undefined],
      ['Test', 'bold', true, undefined],
      ['Locked', 'regular', false, "the font's licence flags forbid embedding (fsType 0x2)"],
      ['Varia', 'regular', false, 'not installed as a static face'],
    ]);
    // In the package: fntdata parts, font relationships and the embedded font list in presentation.xml.
    const zip = unzipSync(buildPptx([], [], 960, 540, 't', fonts));
    expect(Object.keys(zip).filter(k => k.startsWith('ppt/fonts/'))).toEqual(['ppt/fonts/font1.fntdata', 'ppt/fonts/font2.fntdata']);
    const pres = strFromU8(zip['ppt/presentation.xml']);
    expect(pres).toContain(' embedTrueTypeFonts="1"');
    expect(pres).toContain('<p:embeddedFontLst><p:embeddedFont><p:font typeface="Test SemiBold"/><p:regular r:id="rIdF1"/></p:embeddedFont><p:embeddedFont><p:font typeface="Test"/><p:bold r:id="rIdF2"/></p:embeddedFont></p:embeddedFontLst></p:presentation>');
    expect(strFromU8(zip['[Content_Types].xml'])).toContain('<Default Extension="fntdata" ContentType="application/x-fontdata"/>');
    expect(strFromU8(zip['ppt/_rels/presentation.xml.rels'])).toContain('Target="fonts/font2.fntdata"');
  } finally { rmSync(dir, {recursive: true, force: true}); }
});
