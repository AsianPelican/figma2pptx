// Deterministic stand-ins for the machine-dependent parts (installed fonts, ImageMagick), and the fixture deck.
import {cpSync, mkdtempSync} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {mapFaceIn, type FontResolver, type FontRow} from '../src/convert/fonts';
import {imageSize, type ImageOps} from '../src/convert/images';

export const FIXTURE_CACHE = join(import.meta.dir, 'fixtures/figma');
export const FIXTURE_KEY = 'SyntheticDeck0001';

// A fixed font table: Arial static, plus a family installed only as a variable font.
export const FONT_ROWS: FontRow[] = [
  {fams: ['Arial'], styles: ['Regular'], ps: 'ArialMT', weight: 80},
  {fams: ['Arial'], styles: ['Bold'], ps: 'Arial-BoldMT', weight: 200},
  {fams: ['Arial'], styles: ['Italic'], ps: 'Arial-ItalicMT', weight: 80},
  {fams: ['Varia'], styles: ['Regular'], ps: 'Varia-Regular', weight: 80, variable: true, file: '/fonts/Varia[wght].ttf'},
  {fams: ['Varia'], styles: ['SemiBold'], ps: 'Varia-SemiBold', weight: 180, variable: true, file: '/fonts/Varia[wght].ttf'},
];
export const SPACE_EM: Record<string, number> = {'ArialMT': 0.27783, 'Arial-BoldMT': 0.27783};

export const tableFonts = (rows = FONT_ROWS, missing: number[] = []): FontResolver => ({
  mapFace: (ps, family, weight, italic) => mapFaceIn(rows, ps, family, weight, italic),
  spaceEm: ps => SPACE_EM[ps] ?? 0.25,
  hasGlyph: (_face, codePoint) => !missing.includes(codePoint),
});

// Sizes come from the file header; nothing is opaque, so no pixels are ever rewritten.
export const fakeImages: ImageOps = {
  size: imageSize,
  isOpaque: () => false,
  slideCopy: () => { throw Error('fixture rasters need no slide copy'); },
  blurPlate: () => { throw Error('fixture has no blur panels'); },
};

// A private copy of the fixture cache, so a run never writes into the repository.
export function fixtureCacheCopy(): string {
  const dir = mkdtempSync(join(tmpdir(), 'figma2pptx-test-'));
  cpSync(FIXTURE_CACHE, dir, {recursive: true});
  return dir;
}
