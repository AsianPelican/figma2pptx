// The command line front end: parses arguments, renders progress events, prints the summary.
import {readFileSync} from 'node:fs';
import {join, resolve} from 'node:path';
import {convertFigma, pdfFromDeck, fileSize, type PdfPreset, type StageTiming} from '../pipeline';
import {PRESETS} from '../pdf/optimize';
import {createProgress, type Progress} from './progress';

const ROOT = resolve(import.meta.dir, '../..');
export const VERSION: string = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')).version;

export const USAGE = `figma2pptx ${VERSION}: Figma frames to an editable PowerPoint deck, 1:1.

Usage:
  figma2pptx <figma-url | file-key> [frame ...] [-o deck.pptx] [options]
  figma2pptx pdf <deck.pptx | deck.pdf> [-o deck.pdf] [--pdf-preset screen|standard|print]

Frames are ids ("12:34" or "12-34") or exact frame names, converted in the order given. Without frames, a URL's
node-id is used (a page id means every frame on that page), or pass --page.

Options:
  -o, --out <path>          output .pptx (default: ./<Figma file name>.pptx)
  --page <name | id>        every visible frame of this page (or: look frame names up on it only)
  --pdf                     also write <out>.pdf: PowerPoint's export, size-optimized and render-gated
  --pdf-preset <preset>     screen (default), standard, print, or raw (PowerPoint's export untouched)
  --single-pass             skip the measured second pass (no PowerPoint needed unless --pdf)
  --no-embed-fonts          do not embed the fonts the deck uses
  --allow-font-fallback     convert even when a font is missing, variable-only or substituted
  --offline                 use cached Figma data only; no network
  --cache <dir>             Figma cache (default: <repo>/cache, or $FIGMA2PPTX_CACHE)
  --scale <n>               render scale for pictures (default 2)
  --corr <file>             start from saved per-box text offsets
  --timings                 print the time each stage took
  -h, --help                this help
  -v, --version             the version

The Figma token is read from the FIGMA_TOKEN environment variable and is never logged or cached.
`;

export type Args = {positional: string[], out?: string, page?: string, pdf: boolean, pdfPreset: PdfPreset, singlePass: boolean, embedFonts: boolean, allowFontFallback: boolean, offline: boolean, cache: string, scale: number, kern: string, corr?: string, timings: boolean, help: boolean, version: boolean};

export function parseArgs(argv: string[]): Args {
  const a: Args = {positional: [], pdf: false, pdfPreset: 'screen', singlePass: false, embedFonts: true, allowFontFallback: false, offline: false, cache: process.env.FIGMA2PPTX_CACHE || join(ROOT, 'cache'), scale: 2, kern: '100', timings: false, help: false, version: false};
  const value = (i: number, k: string) => { const v = argv[i + 1]; if (v === undefined || v.startsWith('--')) throw Error(`${k} needs a value`); return v; };
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i];
    switch (k) {
      case '-o': case '--out': a.out = value(i++, k); break;
      case '--page': a.page = value(i++, k); break;
      case '--pdf': a.pdf = true; break;
      case '--pdf-preset': a.pdfPreset = value(i++, k) as PdfPreset; a.pdf = true; break;
      case '--single-pass': a.singlePass = true; break;
      case '--no-embed-fonts': a.embedFonts = false; break;
      case '--allow-font-fallback': a.allowFontFallback = true; break;
      case '--offline': a.offline = true; break;
      case '--cache': a.cache = resolve(value(i++, k)); break;
      case '--scale': a.scale = +value(i++, k); if (!(a.scale > 0)) throw Error('--scale must be a positive number'); break;
      case '--kern': a.kern = value(i++, k); break;
      case '--corr': a.corr = value(i++, k); break;
      case '--timings': a.timings = true; break;
      case '-h': case '--help': a.help = true; break;
      case '-v': case '--version': a.version = true; break;
      default:
        if (k.startsWith('-') && k.length > 1) throw Error(`unknown option ${k}`);
        a.positional.push(k);
    }
  }
  if (a.pdfPreset !== 'raw' && !PRESETS[a.pdfPreset]) throw Error(`unknown --pdf-preset ${a.pdfPreset} (screen, standard, print, raw)`);
  return a;
}

export function timingTable(timings: StageTiming[], total: number): string[] {
  const w = Math.max(...timings.map(t => t.stage.length), 5);
  return ['  timings:', ...timings.map(t => `    ${t.stage.padEnd(w)}  ${t.seconds.toFixed(2).padStart(7)}s`), `    ${'total'.padEnd(w)}  ${total.toFixed(2).padStart(7)}s`];
}

async function run(a: Args, p: Progress): Promise<string> {
  const onProgress = (e: {type: string, text: string}) => e.type === 'stage' ? p.stage(e.text) : e.type === 'step' ? p.update(e.text) : p.note(e.text);
  if (a.positional[0] === 'pdf') {
    if (!a.positional[1]) throw Error('usage: figma2pptx pdf <deck.pptx | deck.pdf> [-o out.pdf] [--pdf-preset screen|standard|print]');
    const t0 = Date.now();
    const r = await pdfFromDeck(a.positional[1], {out: a.out, preset: a.pdfPreset, onProgress});
    if (r.result) p.note(`  ${r.result.rawMB} MB -> ${r.result.optimizedMB} MB, ${r.result.photosResized}/${r.result.photos} photos resampled, render gate passed on ${r.result.pages} pages`);
    return `done: ${r.pdf} (${fileSize(r.pdf)}) in ${((Date.now() - t0) / 1000).toFixed(1)}s`;
  }
  if (!a.positional.length) throw Error('missing the Figma URL or file key (see --help)');
  const r = await convertFigma({
    target: a.positional[0], frames: a.positional.slice(1), page: a.page, out: a.out,
    pdf: a.pdf ? a.pdfPreset : false, passes: a.singlePass ? 1 : 2, embedFonts: a.embedFonts, allowFontFallback: a.allowFontFallback,
    offline: a.offline, cacheDir: a.cache, scale: a.scale, kern: a.kern, corr: a.corr ? JSON.parse(readFileSync(a.corr, 'utf8')) : undefined,
    onProgress,
  });
  if (r.placement) {
    p.note(`  text placement after pass 2: ${r.placement.within1px}/${r.placement.measuredLines} measurable lines within 1 px of Figma (worst ${r.placement.maxOffsetPx} px; ${r.placement.unmeasurableLines} unmeasurable)`);
    for (const f of r.placement.failures) p.note(`  placement ${f.offsetPx == null ? 'unmeasurable' : `${f.offsetPx} px`}: ${f.frame ?? 'unknown slide'} ${f.node ?? 'unknown node'} "${f.text}" — ${f.cause}`);
  }
  if (a.timings) for (const line of timingTable(r.timings, r.seconds)) p.note(line);
  const outputs = [r.pptx, ...(r.pdf ? [r.pdf] : [])].map(f => `${f} (${fileSize(f)})`);
  return `done: ${r.slides} slide${r.slides > 1 ? 's' : ''} in ${r.seconds.toFixed(1)}s -> ${outputs.join(', ')}; report ${r.report}`;
}

export async function main(argv: string[], opts: {tty?: boolean} = {}): Promise<number> {
  let a: Args;
  try { a = parseArgs(argv); } catch (e: any) { process.stderr.write(`figma2pptx: ${e.message}\n\n${USAGE}`); return 2; }
  if (a.help) { process.stdout.write(USAGE); return 0; }
  if (a.version) { process.stdout.write(VERSION + '\n'); return 0; }
  const p = createProgress({tty: opts.tty});
  try {
    await p.end(true, await run(a, p));
    return 0;
  } catch (e: any) {
    await p.end(false, `figma2pptx: ${e.message}`);
    return 1;
  }
}
