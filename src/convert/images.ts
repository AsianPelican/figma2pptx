// Pixel work on Figma's renders, through ImageMagick (`magick`). Kept behind an interface so tests can run
// without it.
import {execFileSync} from 'node:child_process';
import {readFileSync} from 'node:fs';

export type Crop = {w: number, h: number, x: number, y: number}; // device pixels
export type Layer = {file: string, x: number, y: number}; // device pixels, top-left of the plate

export interface ImageOps {
  size(file: string): [number, number];
  // Opaque apart from anti-aliased edge pixels.
  isOpaque(file: string): boolean;
  // Write `file` to `out`, optionally cropped; opaque images are flattened on white and written as JPEG q85.
  slideCopy(file: string, out: string, crop: Crop | null, opaque: boolean): void;
  // Background-blur plate: `layers` composited over `bg`, Gaussian blur, then cropped to the panel.
  blurPlate(out: string, w: number, h: number, bg: string, layers: Layer[], sigma: number, crop: Crop): void;
}

// Width and height from a PNG or JPEG header.
export function imageSize(file: string): [number, number] {
  const b = readFileSync(file);
  if (b[0] === 0x89 && b.toString('latin1', 1, 4) === 'PNG') return [b.readUInt32BE(16), b.readUInt32BE(20)];
  if (b[0] === 0xff && b[1] === 0xd8) {
    let o = 2;
    while (o < b.length) {
      if (b[o] !== 0xff) { o++; continue; }
      const mk = b[o + 1];
      if (mk >= 0xc0 && mk <= 0xcf && mk !== 0xc4 && mk !== 0xc8 && mk !== 0xcc) return [b.readUInt16BE(o + 7), b.readUInt16BE(o + 5)];
      o += 2 + b.readUInt16BE(o + 2);
    }
  }
  throw Error(`cannot read image size of ${file}`);
}

export const magickOps: ImageOps = {
  size: imageSize,
  isOpaque: file => +execFileSync('magick', [file, '-alpha', 'extract', '-shave', '3x3', '-format', '%[fx:minima]', 'info:'], {encoding: 'utf8'}).trim() >= 0.999,
  slideCopy(file, out, crop, opaque) {
    execFileSync('magick', [file, ...(crop ? ['-crop', `${crop.w}x${crop.h}+${crop.x}+${crop.y}`, '+repage'] : []), ...(opaque ? ['-background', 'white', '-flatten', '-quality', '85'] : []), out]);
  },
  blurPlate(out, w, h, bg, layers, sigma, crop) {
    const cmd = ['-size', `${w}x${h}`, `xc:${bg}`];
    for (const l of layers) cmd.push('(', l.file, ')', '-geometry', `+${l.x}+${l.y}`, '-composite');
    cmd.push('-blur', `0x${sigma}`, '-crop', `${crop.w}x${crop.h}+${crop.x}+${crop.y}`, '+repage', out);
    execFileSync('magick', cmd);
  },
};
