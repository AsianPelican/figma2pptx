// Size-optimize a PowerPoint PDF, touching only images, behind a render gate.
//
// - Photos (DCT/JPEG): downsampled to the preset ppi at their largest on-page size and re-encoded at the preset
//   JPEG quality; the smaller of old and new is kept. Only opaque plain JPEGs are touched: anything with a soft or
//   stencil mask or a Decode array stays byte-identical (PowerPoint stores some headline rasters that way).
// - Lossless images (Flate): never downsampled and never turned into JPEG, so type stays crisp; re-deflated at
//   level 9 with per-row PNG predictors, which is bit-for-bit lossless.
// - Identical objects deduplicated, all streams Flate-compressed (mupdf garbage=4).
// Text, vector paths and fonts are not rewritten: extracted text must match the input exactly, the page count
// and image placements must not change, and every page must render the same outside re-encoded photos.
import * as m from 'mupdf';
import {readFileSync, writeFileSync, mkdirSync, rmSync, readdirSync} from 'node:fs';
import {execFileSync} from 'node:child_process';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {deflateSync} from 'node:zlib';

export const PRESETS: Record<string, {photoPpi: number, jpegQ: number, chroma: string}> = {
  screen: {photoPpi: 150, jpegQ: 80, chroma: '4:2:0'},
  standard: {photoPpi: 200, jpegQ: 85, chroma: '4:2:0'},
  print: {photoPpi: 300, jpegQ: 90, chroma: '4:4:4'},
};

export type GatePage = {page: number, pass: boolean, reason?: string, worstBlockOutsidePhotos?: number, worstBlockInsideResampledPhotos?: number, photoZones?: number};
export type Gate = {pass: boolean, failed: number[], pages: GatePage[]};
export type OptimizeResult = {preset: string, pages: number, rawMB: number, optimizedMB: number, photos: number, photosResized: number, losslessRecompressed: number, gate: Gate, output: string};

const MB = (n: number) => +(n / 1e6).toFixed(2);

// Lossless re-deflate with per-row PNG predictors (None/Sub/Up/Average/Paeth, minimum-sum heuristic).
function pngFilter(px: Uint8Array, w: number, h: number, bpp: number) {
  const stride = w * bpp, out = new Uint8Array((stride + 1) * h);
  const prev = new Uint8Array(stride);
  for (let y = 0; y < h; y++) {
    const row = px.subarray(y * stride, (y + 1) * stride);
    let best = 0, bestSum = Infinity, bestBuf: Uint8Array | null = null;
    for (let t = 0; t < 5; t++) {
      const b = new Uint8Array(stride); let sum = 0;
      for (let x = 0; x < stride; x++) {
        const a = x >= bpp ? row[x - bpp] : 0, up = prev[x], c = x >= bpp ? prev[x - bpp] : 0; let v = row[x];
        if (t === 1) v -= a; else if (t === 2) v -= up; else if (t === 3) v -= (a + up) >> 1;
        else if (t === 4) { const p = a + up - c, pa = Math.abs(p - a), pb = Math.abs(p - up), pc = Math.abs(p - c); v -= pa <= pb && pa <= pc ? a : pb <= pc ? up : c; }
        b[x] = v & 255; sum += b[x] < 128 ? b[x] : 256 - b[x];
      }
      if (sum < bestSum) { bestSum = sum; best = t; bestBuf = b; }
    }
    out[y * (stride + 1)] = best; out.set(bestBuf!, y * (stride + 1) + 1); prev.set(row);
  }
  return out;
}

const pagesOf = (b: Uint8Array) => m.Document.openDocument(Buffer.from(b), 'application/pdf').countPages();
const textOf = (b: Uint8Array) => { const d = m.Document.openDocument(Buffer.from(b), 'application/pdf'); let t = ''; for (let i = 0; i < d.countPages(); i++) t += (d.loadPage(i) as any).toStructuredText().asText() + '\f'; return t; };
const imageCount = (p: string) => execFileSync('pdfimages', ['-list', p], {encoding: 'utf8'}).trim().split('\n').length - 2;

// Page areas (150 dpi px) covered by photos this run re-encoded, found by their new pixel size. Only there is
// JPEG/resampling difference expected; everywhere else the page must render essentially identically.
function photoZones(pdf: Uint8Array, dims: Set<string>) {
  const d = m.Document.openDocument(Buffer.from(pdf), 'application/pdf'); const zones: Record<number, number[][]> = {}; const k = 150 / 72;
  for (let p = 0; p < d.countPages(); p++) {
    const page = d.loadPage(p); const list: number[][] = [];
    const dev = new m.Device({fillImage(img: any, ctm: number[]) { if (!dims.has(`${img.getWidth()}x${img.getHeight()}`)) return; const xs = [ctm[4], ctm[0] + ctm[4], ctm[2] + ctm[4], ctm[0] + ctm[2] + ctm[4]], ys = [ctm[5], ctm[1] + ctm[5], ctm[3] + ctm[5], ctm[1] + ctm[3] + ctm[5]]; list.push([Math.min(...xs) * k, Math.min(...ys) * k, Math.max(...xs) * k, Math.max(...ys) * k]); }} as any);
    page.run(dev, m.Matrix.identity); dev.close(); zones[p + 1] = list;
  }
  return zones;
}

// Render gate: every page of both files is rendered at 150 dpi with poppler (independent of mupdf, which wrote
// the output) and compared in 12x12 px blocks (mean absolute difference, 0-255).
function renderGate(a: string, b: string, zones: Record<number, number[][]>): Gate {
  const dir = join(tmpdir(), `figma2pptx-gate-${process.pid}-${Date.now()}`); mkdirSync(dir, {recursive: true});
  try {
    execFileSync('pdftoppm', ['-r', '150', a, join(dir, 'a')]); execFileSync('pdftoppm', ['-r', '150', b, join(dir, 'b')]);
    const ppm = (f: string) => { const d = readFileSync(f); let o = 0; const fields: number[] = []; while (fields.length < 3) { while (d[o] === 35 || d[o] <= 32) { if (d[o] === 35) while (d[o] !== 10) o++; o++; } let s = ''; while (d[o] > 32) s += String.fromCharCode(d[o++]); if (s !== 'P6') fields.push(+s); } o++; return {w: fields[0], h: fields[1], px: d.subarray(o)}; };
    const pages: GatePage[] = []; const failed: number[] = [];
    for (const f of readdirSync(dir).filter(x => x.startsWith('a-')).sort()) {
      const n = +f.match(/(\d+)\.ppm$/)![1]; const A = ppm(join(dir, f)), B = ppm(join(dir, f.replace(/^a-/, 'b-')));
      if (A.w !== B.w || A.h !== B.h) { pages.push({page: n, pass: false, reason: 'size'}); failed.push(n); continue; }
      const bs = 12; let worst = 0, over8 = 0, worstPhoto = 0; const Z = zones[n] || [];
      const inZone = (x: number, y: number) => Z.some(z => x + bs > z[0] && y + bs > z[1] && x < z[2] && y < z[3]);
      for (let by = 0; by < A.h; by += bs) for (let bx = 0; bx < A.w; bx += bs) {
        let sum = 0, cnt = 0;
        for (let y = by; y < Math.min(by + bs, A.h); y++) for (let x = bx; x < Math.min(bx + bs, A.w); x++) { const k = (y * A.w + x) * 3; sum += Math.abs(A.px[k] - B.px[k]) + Math.abs(A.px[k + 1] - B.px[k + 1]) + Math.abs(A.px[k + 2] - B.px[k + 2]); cnt += 3; }
        const mean = sum / cnt; if (inZone(bx, by)) { if (mean > worstPhoto) worstPhoto = mean; continue; } if (mean > worst) worst = mean; if (mean > 8) over8++;
      }
      // Outside re-encoded photos: identical to within anti-aliasing (worst block <= 6/255).
      // Inside them: resampling noise is fine, gross damage (a block off by > 90/255, e.g. a lost or black image) is not.
      const pass = worst <= 6 && over8 === 0 && worstPhoto <= 90;
      pages.push({page: n, worstBlockOutsidePhotos: +worst.toFixed(1), worstBlockInsideResampledPhotos: +worstPhoto.toFixed(1), photoZones: Z.length, pass}); if (!pass) failed.push(n);
    }
    return {pass: failed.length === 0, failed, pages};
  } finally { rmSync(dir, {recursive: true, force: true}); }
}

// Optimize `rawPdf` into `outPath`. Throws, and writes nothing, when any check fails.
// maskedPhotos: also resample large JPEGs that carry a soft mask (their mask stays byte-identical).
export function optimizePdf(rawPdf: string, outPath: string, preset: string, opts: {maskedPhotos?: boolean} = {}): OptimizeResult {
  const P = PRESETS[preset];
  if (!P) throw Error(`unknown PDF preset ${preset} (${Object.keys(PRESETS).join(', ')})`);
  const rawBytes = readFileSync(rawPdf), slides = pagesOf(rawBytes);
  // Largest on-page size of every image object.
  const imgUse = new Map<number, {minPpi: number}>();
  for (const line of execFileSync('pdfimages', ['-list', rawPdf], {encoding: 'utf8'}).trim().split('\n').slice(2)) {
    const c = line.trim().split(/\s+/); const obj = +c[10], ppi = Math.min(+c[12], +c[13]);
    const u = imgUse.get(obj); if (!u || ppi < u.minPpi) imgUse.set(obj, {minPpi: ppi});
  }
  const rawText = textOf(rawBytes);
  const doc = m.Document.openDocument(Buffer.from(rawBytes), 'application/pdf') as m.PDFDocument;
  let photos = 0, photosResized = 0, lossless = 0; const changedDims = new Set<string>();
  for (let i = 1; i < doc.countObjects(); i++) {
    const ref = (doc as any).newIndirect(i, 0) as m.PDFObject; if (!ref.isStream()) continue;
    // Read through the stream ref only: resolve() on every object makes mupdf repair the Quartz xref and
    // silently drops image references on save.
    const d = ref; const st = d.get('Subtype'); if (st.isNull() || st.asName() !== 'Image') continue;
    const f = d.get('Filter'); const filter = f.isNull() ? '' : f.isArray() ? 'multi' : f.asName();
    if (filter === 'FlateDecode') { // lossless: same pixels, stronger deflate
      const w = d.get('Width').asNumber(), h = d.get('Height').asNumber(), bpc = d.get('BitsPerComponent').asNumber();
      const raw = ref.readRawStream().getLength(), px = ref.readStream().asUint8Array(), bpp = px.length / (w * h);
      if (bpc !== 8 || !Number.isInteger(bpp)) continue;
      const plain = deflateSync(px, {level: 9}), pred = deflateSync(pngFilter(px, w, h, bpp), {level: 9});
      const usePred = pred.length < plain.length, best = usePred ? pred : plain; if (best.length >= raw) continue;
      ref.writeRawStream(best);
      if (usePred) { const dp = doc.newDictionary(); dp.put('Predictor', 15); dp.put('Colors', bpp); dp.put('BitsPerComponent', 8); dp.put('Columns', w); ref.put('DecodeParms', dp); } else ref.delete('DecodeParms');
      lossless++; continue;
    }
    if (filter !== 'DCTDecode') continue;
    const smasked = !d.get('SMask').isNull();
    if (smasked && !(opts.maskedPhotos && d.get('Width').asNumber() * d.get('Height').asNumber() >= 500000)) continue;
    if (!d.get('Mask').isNull() || !d.get('Decode').isNull() || (!d.get('ImageMask').isNull() && d.get('ImageMask').asBoolean())) continue;
    if (d.get('BitsPerComponent').asNumber() !== 8) continue;
    const use = imgUse.get(i); if (!use) continue;
    const w = d.get('Width').asNumber(), h = d.get('Height').asNumber();
    const scale = Math.min(1, P.photoPpi / use.minPpi), nw = Math.max(1, Math.round(w * scale)), nh = Math.max(1, Math.round(h * scale));
    const bytes = ref.readRawStream().asUint8Array();
    const cs = d.get('ColorSpace'); const cmyk = !cs.isNull() && cs.toString().includes('CMYK');
    const out = execFileSync('magick', ['jpeg:-', ...(scale < 1 ? ['-filter', 'Lanczos', '-resize', `${nw}x${nh}!`] : []), '-strip', '-sampling-factor', P.chroma, '-quality', String(P.jpegQ), ...(cmyk ? ['-colorspace', 'CMYK'] : []), 'jpeg:-'], {input: bytes, maxBuffer: 80e6});
    photos++;
    if (out.length >= bytes.length && scale === 1) continue;
    const cspace = (b: Uint8Array) => execFileSync('magick', ['identify', '-format', '%[colorspace]', 'jpeg:-'], {input: b, encoding: 'utf8'}).trim();
    if (cspace(out) !== cspace(bytes)) continue; // colour model must match the PDF colour space
    ref.writeRawStream(out); ref.put('Width', nw); ref.put('Height', nh); if (scale < 1) photosResized++; changedDims.add(`${nw}x${nh}`);
  }
  const buf = Buffer.from(doc.saveToBuffer('garbage=4,compress=yes').asUint8Array());
  const tmp = join(tmpdir(), `figma2pptx-opt-${process.pid}-${Date.now()}.pdf`);
  try {
    writeFileSync(tmp, buf);
    if (imageCount(tmp) !== imageCount(rawPdf)) throw Error(`image placements changed in the ${preset} PDF`);
    const gate = renderGate(rawPdf, tmp, photoZones(buf, changedDims));
    if (!gate.pass) throw Error(`render gate failed for the ${preset} PDF: pages ${gate.failed.join(', ')} differ beyond JPEG noise`);
    if (textOf(buf) !== rawText) throw Error(`text changed in the ${preset} PDF`);
    if (pagesOf(buf) !== slides) throw Error(`page count changed in the ${preset} PDF`);
    writeFileSync(outPath, buf);
    return {preset, pages: slides, rawMB: MB(rawBytes.length), optimizedMB: MB(buf.length), photos, photosResized, losslessRecompressed: lossless, gate, output: outPath};
  } finally { rmSync(tmp, {force: true}); }
}
