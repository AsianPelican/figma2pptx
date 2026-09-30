import {readFileSync} from 'node:fs';
import {convertFigma} from '../../src/index.ts';
import type {ConversionJob} from './runner.ts';

const send = (value: unknown) => process.stdout.write(JSON.stringify(value) + '\n');

export async function subprocessMain(path: string | undefined): Promise<number> {
  if (!path) { send({type: 'error', error: 'bridge worker request path is missing'}); return 2; }
  try {
    const data = JSON.parse(readFileSync(path, 'utf8')) as ConversionJob;
    const result = await convertFigma({
      target: data.fileKey,
      frames: data.frames,
      page: data.page,
      out: data.out,
      pdf: data.pdf,
      passes: data.passes,
      allowFontFallback: data.allowFontFallback,
      offline: data.offline,
      cacheDir: data.cacheDir,
      scale: data.scale,
      kern: data.kern,
      corr: data.corr,
      onProgress: event => send({type: 'progress', event}),
    });
    send({type: 'done', result: {pptx: result.pptx, pdf: result.pdf, report: result.report, slides: result.slides, seconds: result.seconds}});
    return 0;
  } catch (error) {
    send({type: 'error', error: error instanceof Error ? error.message : String(error)});
    return 1;
  }
}

if (import.meta.main) process.exit(await subprocessMain(process.argv[2]));
