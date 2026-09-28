// Byte-stable golden output: the synthetic deck, converted offline from fixed Figma responses with a fixed font
// table, must give exactly the committed PPTX. Regenerate deliberately with UPDATE_GOLDEN=1 bun test.
import {test, expect} from 'bun:test';
import {readFileSync, writeFileSync, existsSync, rmSync} from 'node:fs';
import {join} from 'node:path';
import {unzipSync, strFromU8} from 'fflate';
import {convertFigma} from '../src/pipeline';
import {FIXTURE_KEY, tableFonts, fakeImages, fixtureCacheCopy} from './helpers';

const GOLDEN = join(import.meta.dir, 'fixtures/golden');

async function convert() {
  const cache = fixtureCacheCopy();
  try {
    const r = await convertFigma({target: FIXTURE_KEY, page: 'Slides', out: join(cache, 'deck.pptx'), passes: 1, offline: true, cacheDir: cache, fonts: tableFonts(), images: fakeImages, allowFontFallback: false});
    const report = JSON.parse(readFileSync(r.report, 'utf8'));
    for (const k of ['cache', 'seconds', 'timings']) delete report[k];
    return {pptx: readFileSync(r.pptx), report};
  } finally { rmSync(cache, {recursive: true, force: true}); }
}

test('synthetic deck converts to the golden PPTX, byte for byte', async () => {
  const {pptx, report} = await convert();
  if (process.env.UPDATE_GOLDEN) {
    writeFileSync(join(GOLDEN, 'synthetic.pptx'), pptx);
    writeFileSync(join(GOLDEN, 'synthetic.report.json'), JSON.stringify(report, null, 1) + '\n');
  }
  expect(existsSync(join(GOLDEN, 'synthetic.pptx'))).toBe(true);
  const want = readFileSync(join(GOLDEN, 'synthetic.pptx'));
  if (!pptx.equals(want)) {
    // Name the parts that differ, so a failure says where to look.
    const a = unzipSync(pptx), b = unzipSync(want);
    const parts = [...new Set([...Object.keys(a), ...Object.keys(b)])].filter(k => !a[k] || !b[k] || strFromU8(a[k]) !== strFromU8(b[k]));
    expect(parts).toEqual([]);
  }
  expect(pptx.equals(want)).toBe(true);
  expect(report).toEqual(JSON.parse(readFileSync(join(GOLDEN, 'synthetic.report.json'), 'utf8')));
});

test('two runs give identical bytes', async () => {
  const [a, b] = [await convert(), await convert()];
  expect(a.pptx.equals(b.pptx)).toBe(true);
});
