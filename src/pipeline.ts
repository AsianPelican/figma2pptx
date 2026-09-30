// The conversion as a library call: a Figma file and frames in, a PPTX (and optionally a PDF) out, with
// progress events. The CLI renders the events as a spinner; any other front end (a local server behind a Figma
// plugin, say) can consume the same events.
//
//   build pass 1 -> font preflight -> PowerPoint export -> measure every text line against Figma's baselines
//   -> build pass 2 with one offset per text box -> PowerPoint export -> check -> optional PDF optimization
import {writeFileSync, readFileSync, mkdirSync, rmSync, existsSync, statSync} from 'node:fs';
import {join, resolve, dirname} from 'node:path';
import {tmpdir} from 'node:os';
import {FigmaClient, readToken, parseTarget} from './figma/api';
import {FigmaFile, resolveFrames} from './figma/source';
import {buildDeck, type BuildReport, type Corrections} from './convert/build';
import {fontconfigResolver, type FontResolver} from './convert/fonts';
import {magickOps, type ImageOps} from './convert/images';
import {exportPdf} from './powerpoint/export';
import {measure, type Measurement} from './measure/corrections';
import {optimizePdf, type OptimizeResult} from './pdf/optimize';
import type {TextTransformOptions} from './convert/text-transform';

// stage: a new step (with its own duration); step: progress inside the current stage; note: a line to keep.
export type ProgressEvent = {type: 'stage' | 'step' | 'note', text: string};
export type PdfPreset = 'screen' | 'standard' | 'print' | 'raw';

export type ConvertOptions = {
  target: string; // Figma URL or file key
  frames?: string[]; // frame ids or exact names; default: the URL's node-id or --page
  page?: string; // page name or id
  out?: string; // .pptx path; default ./<Figma file name>.pptx
  pdf?: PdfPreset | false;
  passes?: 1 | 2; // 2 (default): measured second pass through PowerPoint
  textTransform?: TextTransformOptions; // host-supplied policy; omitted by the public CLI
  allowFontFallback?: boolean; // default false: missing, variable-only or substituted fonts are an error
  offline?: boolean; // cached Figma data only
  token?: string; // default: readToken()
  cacheDir: string;
  scale?: number; // picture render scale, default 2
  kern?: string; // PowerPoint kern threshold (pt x 100), default "100"
  corr?: Corrections; // starting per-box offsets
  onProgress?: (e: ProgressEvent) => void;
  // injectable for tests
  fonts?: FontResolver;
  images?: ImageOps;
};

export type StageTiming = {stage: string, seconds: number};
export type PlacementReport = {
  totalLines: number;
  measuredLines: number;
  within1px: number;
  unmeasurableLines: number;
  maxOffsetPx: number;
  failures: {frame?: string, node?: string, text: string, dx?: number, dy?: number, offsetPx?: number, cause: string}[];
};

export type ConvertResult = {
  timings: StageTiming[]; // wall time per stage, in order
  pptx: string; pdf?: string; report: string;
  slides: number; seconds: number; fileName: string;
  build: BuildReport;
  measurement?: Pick<Measurement, 'measured' | 'unmeasurable' | 'medianDy' | 'spreadMedian' | 'spreadMax'>;
  placement?: PlacementReport; // after pass 2
  transformedPlacement?: PlacementReport; // substituted live-text lines only
  transformedPlacementByFrame?: (PlacementReport & {frame: string, name: string})[];
  pdfResult?: OptimizeResult;
};

export class FontPreflightError extends Error {
  constructor(readonly faces: BuildReport['faces']) {
    super(`${faces.length} font${faces.length > 1 ? 's' : ''} PowerPoint cannot render as designed:\n` + faces.map(f => `  ${f.figma}: ${f.note}`).join('\n') +
      `\nInstall the static faces (PowerPoint for Mac cannot use variable fonts by weight), then run again; or allow the fallbacks above (--allow-font-fallback).`);
  }
}

const safeName = (s: string) => s.replace(/[\/\\:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim() || 'deck';
// Every face the deck uses and what PowerPoint will use for it. Fonts always remain external.
export function fontTable(report: BuildReport): string[] {
  const mark = {exact: 'ok   ', substitute: 'SUBST', 'variable-only': 'VAR  ', missing: 'MISS '} as const;
  return report.faces.map(f => {
    const target = `${f.typeface}${f.b ? ' Bold' : ''}${f.i ? ' Italic' : ''}`;
    return `  ${mark[f.status]} ${f.figma} -> ${target}${f.note && f.status !== 'exact' ? ` (${f.note})` : ''}`;
  });
}

export async function convertFigma(o: ConvertOptions): Promise<ConvertResult> {
  const t0 = Date.now();
  const timings: StageTiming[] = [];
  let open: {stage: string, at: number} | null = null;
  const closeStage = () => { if (open) timings.push({stage: open.stage, seconds: +((Date.now() - open.at) / 1000).toFixed(2)}); open = null; };
  const emit = (type: ProgressEvent['type'], text: string) => {
    if (type === 'stage') { closeStage(); open = {stage: text, at: Date.now()}; }
    o.onProgress?.({type, text});
  };
  const target = parseTarget(o.target);
  const client = o.offline ? null : new FigmaClient(o.token ?? readToken());
  const passes = o.passes ?? 2, scale = o.scale ?? 2, kern = o.kern ?? '100';

  emit('stage', o.offline ? `reading cached Figma file ${target.fileKey}` : `asking Figma for file ${target.fileKey}`);
  const file = await FigmaFile.open(target.fileKey, o.cacheDir, client);
  const frames = resolveFrames(file.meta, o.frames ?? [], {page: o.page, nodeId: target.nodeId});
  if (!frames.length) throw Error('no frames to convert');
  const out = resolve(o.out ? (/\.pptx$/i.test(o.out) ? o.out : o.out + '.pptx') : `${safeName(file.meta.name)}.pptx`);
  const stem = out.replace(/\.pptx$/i, '');
  mkdirSync(dirname(out), {recursive: true});

  emit('stage', `fetching ${frames.length} frame${frames.length > 1 ? 's' : ''} from Figma ("${file.meta.name}")`);
  await file.fetchFrames(frames);

  const env = {fonts: o.fonts ?? fontconfigResolver(), images: o.images ?? magickOps};
  const initial = o.corr ?? {};
  const build = (pass: number, corr: Corrections) => {
    const label = passes === 1 ? 'building' : `pass ${pass}/2: building`;
    emit('stage', label);
    return buildDeck(file, frames, {scale, kern, corr, textTransform: o.textTransform, progress: s => emit('step', `${label}: ${s}`)}, env);
  };
  let {pptx, report} = await build(1, initial);

  // Font preflight: every face must be one PowerPoint can actually render at the weight Figma uses.
  emit('stage', 'checking fonts');
  for (const line of fontTable(report)) emit('note', line);
  const bad = report.faces.filter(f => f.status !== 'exact');
  if (bad.length && !o.allowFontFallback) throw new FontPreflightError(bad);

  const work = join(tmpdir(), `figma2pptx-${process.pid}-${Date.now()}`);
  mkdirSync(work, {recursive: true});
  let m1: Measurement | undefined, m2: Measurement | undefined, pdfResult: OptimizeResult | undefined, pdfPath: string | undefined;
  try {
    if (passes === 2) {
      const p1 = join(work, 'pass1.pptx'), p1pdf = join(work, 'pass1.pdf');
      writeFileSync(p1, pptx);
      emit('stage', 'pass 1/2: PowerPoint export');
      exportPdf(p1, p1pdf);
      emit('stage', 'pass 1/2: measuring text placement against Figma');
      m1 = measure(readFileSync(p1pdf), report, initial);
      ({pptx, report} = await build(2, m1.corr));
    }
    writeFileSync(out, pptx);
    let rawPdf: string | undefined;
    if (passes === 2 || o.pdf) {
      rawPdf = join(work, 'final.pdf');
      emit('stage', passes === 2 ? 'pass 2/2: PowerPoint export' : 'PowerPoint export');
      exportPdf(out, rawPdf);
      if (passes === 2) {
        emit('stage', 'pass 2/2: checking text placement');
        m2 = measure(readFileSync(rawPdf), report, m1!.corr);
      }
    }
    if (o.pdf) {
      pdfPath = stem + '.pdf';
      if (o.pdf === 'raw') writeFileSync(pdfPath, readFileSync(rawPdf!));
      else { emit('stage', `optimizing the PDF (${o.pdf})`); pdfResult = optimizePdf(rawPdf!, pdfPath, o.pdf); }
    }
  } finally {
    rmSync(work, {recursive: true, force: true});
  }

  closeStage();
  // After pass 2: every extractable line must be within one Figma pixel. PowerPoint rasterizes some transparent
  // and rotated text in its PDF; those lines stay explicit and named instead of being counted as passes.
  const placementOf = (lines: Measurement['lines']) => {
    const measured = lines.filter(l => l.status === 'placed');
    const failures = lines.filter(l => l.status !== 'placed' || l.offsetPx! > 1).map(l => ({
      frame: report.textLines.find(x => x.node === l.node && x.t === l.t)?.frame,
      node: l.node,
      text: l.t.trim().slice(0, 100),
      dx: l.dx,
      dy: l.dy,
      offsetPx: l.offsetPx,
      cause: l.cause ?? `${l.offsetPx} px exceeds 1 px`,
    }));
    return {
      totalLines: lines.length,
      measuredLines: measured.length,
      within1px: measured.filter(x => x.offsetPx! <= 1).length,
      unmeasurableLines: lines.length - measured.length,
      maxOffsetPx: +Math.max(0, ...measured.map(x => x.offsetPx!)).toFixed(2),
      failures,
    };
  };
  const placement = m2 && placementOf(m2.lines);
  const transformedKeys = new Set(report.textLines.filter(l => l.substituted).map(l => `${l.node}\0${l.t}`));
  const transformedPlacement = m2 && o.textTransform ? placementOf(m2.lines.filter(l => transformedKeys.has(`${l.node}\0${l.t}`))) : undefined;
  const transformedPlacementByFrame = m2 && o.textTransform ? report.frames.map(frame => {
    const keys = new Set(report.textLines.filter(l => l.frame === frame.id && l.substituted).map(l => `${l.node}\0${l.t}`));
    return {frame: frame.id, name: frame.name, ...placementOf(m2!.lines.filter(l => keys.has(`${l.node}\0${l.t}`)))};
  }) : undefined;
  const summary = (m?: Measurement) => m && {measured: m.measured, unmeasurable: m.unmeasurable, medianDy: m.medianDy, spreadMedian: m.spreadMedian, spreadMax: m.spreadMax};
  const seconds = +((Date.now() - t0) / 1000).toFixed(1);
  const reportPath = stem + '.report.json';
  writeFileSync(reportPath, JSON.stringify({
    fileKey: file.meta.fileKey, fileName: file.meta.name, version: file.meta.version, frameIds: frames, cache: file.dir, seconds, timings, passes,
    measurement: summary(m1), corrections: m1?.corr, placement, transformedPlacement, transformedPlacementByFrame, pdf: pdfResult ?? (o.pdf ? {preset: o.pdf} : undefined),
    ...report,
  }, null, 1));
  for (const w of report.warnings) emit('note', `  warning: ${w}`);
  return {pptx: out, pdf: pdfPath, report: reportPath, slides: frames.length, seconds, timings, fileName: file.meta.name, build: report, measurement: summary(m1), placement, transformedPlacement, transformedPlacementByFrame, pdfResult};
}

// PowerPoint's PDF of an existing deck (or an existing PDF), size-optimized.
export async function pdfFromDeck(src: string, o: {out?: string, preset?: PdfPreset, onProgress?: (e: ProgressEvent) => void} = {}): Promise<{pdf: string, result?: OptimizeResult}> {
  if (!existsSync(src)) throw Error(`no such file: ${src}`);
  const preset = o.preset ?? 'screen';
  const out = resolve(o.out ?? src.replace(/\.(pptx|pdf)$/i, '') + (/\.pdf$/i.test(src) ? `_${preset}.pdf` : '.pdf'));
  const work = join(tmpdir(), `figma2pptx-${process.pid}-${Date.now()}`);
  mkdirSync(work, {recursive: true});
  try {
    let raw = resolve(src);
    if (/\.pptx$/i.test(src)) { o.onProgress?.({type: 'stage', text: 'PowerPoint export'}); raw = join(work, 'raw.pdf'); exportPdf(src, raw); }
    if (preset === 'raw') { writeFileSync(out, readFileSync(raw)); return {pdf: out}; }
    o.onProgress?.({type: 'stage', text: `optimizing the PDF (${preset})`});
    return {pdf: out, result: optimizePdf(raw, out, preset)};
  } finally { rmSync(work, {recursive: true, force: true}); }
}

export const fileSize = (f: string) => { const n = statSync(f).size; return n < 1e6 ? `${Math.max(1, Math.round(n / 1e3))} KB` : `${(n / 1e6).toFixed(1)} MB`; };
