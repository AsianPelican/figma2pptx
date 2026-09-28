// Just enough of the OpenType (sfnt) format to embed a font: the table directory (of a single font or one
// font of a collection), the naming table, OS/2 and head.
export type SfntInfo = {
  outlines: 'truetype' | 'cff';
  variable: boolean; // has an fvar table
  names: Record<number, string>; // name ID -> Windows (Unicode) name, falling back to Mac Roman
  fsType: number; // OS/2 embedding permissions
  weight: number; // OS/2 usWeightClass
  italic: boolean;
  panose: Uint8Array; // 10 bytes
  unicodeRange: number[]; // 4 x uint32
  codePageRange: number[]; // 2 x uint32
  checkSumAdjustment: number;
  data: Uint8Array; // the font as a standalone sfnt (extracted from a collection when needed)
};

const u16 = (b: Uint8Array, o: number) => (b[o] << 8) | b[o + 1];
const u32 = (b: Uint8Array, o: number) => ((b[o] << 24) >>> 0) + ((b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]);
const tag = (b: Uint8Array, o: number) => String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]);

type Table = {tag: string, checksum: number, offset: number, length: number};

function tables(b: Uint8Array, dir: number): Table[] {
  const n = u16(b, dir + 4);
  return Array.from({length: n}, (_, i) => { const o = dir + 12 + i * 16; return {tag: tag(b, o), checksum: u32(b, o + 4), offset: u32(b, o + 8), length: u32(b, o + 12)}; });
}

// A standalone sfnt holding the given tables (used to lift one font out of a .ttc).
function standalone(b: Uint8Array, dir: number, ts: Table[]): Uint8Array {
  const n = ts.length, head = 12 + 16 * n;
  let size = head; for (const t of ts) size += (t.length + 3) & ~3;
  const out = new Uint8Array(size), dv = new DataView(out.buffer);
  out.set(b.subarray(dir, dir + 12)); // sfnt version, numTables and search fields
  let off = head;
  ts.forEach((t, i) => {
    const o = 12 + i * 16;
    for (let k = 0; k < 4; k++) out[o + k] = t.tag.charCodeAt(k);
    dv.setUint32(o + 4, t.checksum); dv.setUint32(o + 8, off); dv.setUint32(o + 12, t.length);
    out.set(b.subarray(t.offset, t.offset + t.length), off);
    off += (t.length + 3) & ~3;
  });
  return out;
}

export function readSfnt(file: Uint8Array, index = 0): SfntInfo {
  let dir = 0;
  if (tag(file, 0) === 'ttcf') {
    const count = u32(file, 8);
    if (index >= count) throw Error(`font collection has ${count} fonts, no index ${index}`);
    dir = u32(file, 12 + 4 * index);
  }
  const ts = tables(file, dir);
  const t = (name: string) => ts.find(x => x.tag === name);
  const os2 = t('OS/2'), head = t('head'), name = t('name');
  if (!os2 || !head || !name) throw Error('font lacks OS/2, head or name table');
  const names: Record<number, string> = {};
  const nb = name.offset, count = u16(file, nb + 2), strings = nb + u16(file, nb + 4);
  for (let i = 0; i < count; i++) {
    const r = nb + 6 + i * 12, platform = u16(file, r), encoding = u16(file, r + 2), lang = u16(file, r + 4), id = u16(file, r + 6), len = u16(file, r + 8), off = strings + u16(file, r + 10);
    if (platform === 3 && (encoding === 1 || encoding === 0) && (lang === 0x409 || !(id in names))) {
      let s = ''; for (let k = 0; k < len; k += 2) s += String.fromCharCode(u16(file, off + k));
      names[id] = s;
    } else if (platform === 1 && encoding === 0 && !(id in names)) {
      names[id] = String.fromCharCode(...file.subarray(off, off + len));
    }
  }
  const o = os2.offset;
  return {
    outlines: t('CFF ') || t('CFF2') ? 'cff' : 'truetype',
    variable: !!t('fvar'),
    names,
    fsType: u16(file, o + 8),
    weight: u16(file, o + 4),
    italic: (u16(file, o + 62) & 1) === 1,
    panose: file.slice(o + 32, o + 42),
    unicodeRange: [0, 1, 2, 3].map(k => u32(file, o + 42 + 4 * k)),
    codePageRange: os2.length >= 86 ? [u32(file, o + 78), u32(file, o + 82)] : [0, 0],
    checkSumAdjustment: u32(file, head.offset + 8),
    data: dir ? standalone(file, dir, ts) : file,
  };
}

// OS/2 fsType: bit 1 (0x2) alone means "restricted license embedding": the font must not be embedded. Bit 9
// (0x200) means bitmap embedding only. Everything else (installable, preview & print, editable) allows it.
export function embeddingAllowed(fsType: number): boolean {
  if (fsType & 0x200) return false;
  return (fsType & 0xf) !== 0x2;
}

// Embedded OpenType (EOT, version 0x00020001), uncompressed and unencrypted: the container PowerPoint stores
// embedded fonts in (ppt/fonts/*.fntdata).
export function toEot(f: SfntInfo): Uint8Array {
  const utf16 = (s: string) => { const b = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { b[2 * i] = s.charCodeAt(i) & 255; b[2 * i + 1] = s.charCodeAt(i) >> 8; } return b; };
  const names = [f.names[1], f.names[2], f.names[5], f.names[4]].map(s => utf16(s || '')); // family, style, version, full
  // 80 bytes of fixed fields, then per name: padding (2), size (2), UTF-16LE bytes; then padding, RootStringSize.
  const size = 80 + names.reduce((a, n) => a + 4 + n.length, 0) + 4 + f.data.length;
  const out = new Uint8Array(size), dv = new DataView(out.buffer);
  dv.setUint32(0, size, true);
  dv.setUint32(4, f.data.length, true);
  dv.setUint32(8, 0x00020001, true);
  dv.setUint32(12, 0, true); // flags: not subset, not compressed, not XOR-encrypted
  out.set(f.panose.subarray(0, 10), 16);
  out[26] = 1; // DEFAULT_CHARSET
  out[27] = f.italic ? 1 : 0;
  dv.setUint32(28, f.weight, true);
  dv.setUint16(32, f.fsType, true);
  dv.setUint16(34, 0x504c, true);
  f.unicodeRange.forEach((v, k) => dv.setUint32(36 + 4 * k, v, true));
  f.codePageRange.forEach((v, k) => dv.setUint32(52 + 4 * k, v, true));
  dv.setUint32(60, f.checkSumAdjustment, true);
  let o = 80; // 64..79 reserved, zero
  for (const n of names) { o += 2; dv.setUint16(o, n.length, true); o += 2; out.set(n, o); o += n.length; }
  o += 2; dv.setUint16(o, 0, true); o += 2; // no root string
  out.set(f.data, o);
  return out;
}
