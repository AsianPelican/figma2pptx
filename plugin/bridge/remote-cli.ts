import {createHash} from 'node:crypto';
import {existsSync, mkdirSync, readFileSync, renameSync, rmSync} from 'node:fs';
import {homedir} from 'node:os';
import {basename, dirname, join, resolve} from 'node:path';
import {createProgress} from '../../src/cli/progress.ts';
import {BRIDGE_ORIGIN, parseBridgeOrigin, type BridgeEvent, type DownloadOutput, type ExportRequest, type PdfPreset} from './protocol.ts';

const VERSION: string = JSON.parse(readFileSync(join(import.meta.dir, '..', '..', 'package.json'), 'utf8')).version;

export const REMOTE_USAGE = `figma2pptx ${VERSION}: send Figma conversion to the bridge Mac and download the result.

Usage:
  figma2pptx <figma-url | file-key> [frame ...] [-o deck.pptx] [options]

Frames are ids ("12:34" or "12-34") or exact frame names. Without frames, a URL's node-id is used, or pass --page.

Options:
  -o, --out <path>          local output .pptx (default: ~/Downloads/<remote name>)
  --page <name | id>        every visible frame of this page
  --pdf                     also download a PDF
  --pdf-preset <preset>     screen (default), standard, print, or raw
  --single-pass             skip PowerPoint measurement unless --pdf needs export
  --allow-font-fallback     continue when a face is missing/substituted
  --offline                 use the bridge Mac's cache only
  --scale <n>               picture render scale (default 2)
  --kern <n>                PowerPoint kern threshold (default 100)
  --corr <file>             upload saved per-box text offsets with the job
  --timings                 retain timing progress from the remote converter
  -h, --help                this help
  -v, --version             the version

The placeholder bridge URL is ${BRIDGE_ORIGIN}; configure the real private-network origin with FIGMA2PPTX_BRIDGE_ORIGIN.
The secret file defaults to ~/.config/figma2pptx/bridge-secret and can be overridden with FIGMA2PPTX_BRIDGE_SECRET_FILE. Secrets are sent only in the Authorization header.
`;

export type RemoteCommand = {
  help: boolean;
  version: boolean;
  out?: string;
  timings: boolean;
  request?: ExportRequest;
};

function parseTarget(value: string): {fileKey: string; nodeId?: string} {
  if (/^[A-Za-z0-9]{10,128}$/.test(value)) return {fileKey: value};
  let url: URL;
  try { url = new URL(value); } catch { throw Error(`not a Figma URL or file key: ${value}`); }
  if (!/(^|\.)figma\.com$/.test(url.hostname)) throw Error(`not a figma.com URL: ${value}`);
  const parts = url.pathname.split('/').filter(Boolean);
  const marker = parts.findIndex(part => /^(file|design|proto|board|slides|deck)$/.test(part));
  if (marker < 0 || !parts[marker + 1]) throw Error(`no file key in ${value}`);
  const fileKey = parts[marker + 2] === 'branch' && parts[marker + 3] ? parts[marker + 3] : parts[marker + 1];
  const node = url.searchParams.get('node-id');
  return {fileKey, ...(node ? {nodeId: node.replace(/^(\d+)-(\d+)$/, '$1:$2')} : {})};
}

export function parseRemoteArgs(argv: string[]): RemoteCommand {
  const positional: string[] = [];
  let out: string | undefined, page: string | undefined, corr: ExportRequest['corr'];
  let pdf: PdfPreset | false = false, passes: 1 | 2 = 2, allowFontFallback = false, offline = false;
  let scale = 2, kern = '100', timings = false, help = false, version = false;
  const value = (i: number, key: string) => {
    const next = argv[i + 1];
    if (next === undefined || next.startsWith('--')) throw Error(`${key} needs a value`);
    return next;
  };
  for (let i = 0; i < argv.length; i++) {
    const key = argv[i];
    switch (key) {
      case '-o': case '--out': out = resolve(value(i++, key)); break;
      case '--page': page = value(i++, key); break;
      case '--pdf': pdf = 'screen'; break;
      case '--pdf-preset': pdf = value(i++, key) as PdfPreset; break;
      case '--single-pass': passes = 1; break;
      case '--no-embed-fonts': break;
      case '--allow-font-fallback': allowFontFallback = true; break;
      case '--offline': offline = true; break;
      case '--scale': scale = Number(value(i++, key)); break;
      case '--kern': kern = value(i++, key); break;
      case '--corr': corr = JSON.parse(readFileSync(resolve(value(i++, key)), 'utf8')); break;
      case '--timings': timings = true; break;
      case '--cache': throw Error('--cache is not accepted by the remote client; cache lives on the bridge Mac');
      case '-h': case '--help': help = true; break;
      case '-v': case '--version': version = true; break;
      default:
        if (key.startsWith('-') && key.length > 1) throw Error(`unknown option ${key}`);
        positional.push(key);
    }
  }
  if (help || version) return {help, version, out, timings};
  if (!positional.length) throw Error('missing the Figma URL or file key (see --help)');
  if (positional[0] === 'pdf') throw Error('the remote client accepts Figma conversions only; pass --pdf to receive a PDF with the PPTX');
  if (pdf !== false && !['screen', 'standard', 'print', 'raw'].includes(pdf)) throw Error(`unknown --pdf-preset ${pdf} (screen, standard, print, raw)`);
  if (!(scale > 0 && scale <= 8)) throw Error('--scale must be between 0 and 8');
  if (!/^\d{1,6}$/.test(kern)) throw Error('--kern must be digits');

  const target = parseTarget(positional[0]);
  const frames = positional.slice(1);
  if (!frames.length && target.nodeId) frames.push(target.nodeId);
  if (!frames.length && !page) throw Error('name frames, use --page, or pass a Figma URL with node-id');
  const outputName = out ? basename(/\.pptx$/i.test(out) ? out : `${out}.pptx`) : undefined;
  return {
    help, version, out, timings,
    request: {
      fileKey: target.fileKey,
      frames,
      frameNames: frames,
      ...(page ? {page} : {}),
      ...(outputName ? {outputName} : {}),
      pdf,
      passes,
      embedFonts: false,
      allowFontFallback,
      offline,
      scale,
      kern,
      ...(corr ? {corr} : {}),
    },
  };
}

export function bridgeConfig(env: Record<string, string | undefined> = process.env): {origin: string; secret: string} {
  const origin = parseBridgeOrigin(env.FIGMA2PPTX_BRIDGE_ORIGIN || BRIDGE_ORIGIN);
  if (origin === BRIDGE_ORIGIN) throw Error('FIGMA2PPTX_BRIDGE_ORIGIN is not configured');
  const secretPath = env.FIGMA2PPTX_BRIDGE_SECRET_FILE || join(homedir(), '.config', 'figma2pptx', 'bridge-secret');
  if (!existsSync(secretPath)) throw Error(`bridge secret not found at ${secretPath}; copy it securely from the bridge Mac`);
  const secret = readFileSync(secretPath, 'utf8').trim();
  if (!/^[a-f0-9]{64}$/.test(secret)) throw Error(`invalid bridge secret at ${secretPath}`);
  return {origin, secret};
}

async function download(origin: string, secret: string, remote: DownloadOutput, localPath: string): Promise<void> {
  const response = await fetch(origin + remote.downloadPath, {headers: {Authorization: `Bearer ${secret}`}});
  if (!response.ok) throw Error(`download ${remote.name}: ${response.status} ${(await response.text()).slice(0, 200)}`);
  const bytes = new Uint8Array(await response.arrayBuffer());
  const hash = createHash('sha256').update(bytes).digest('hex');
  if (bytes.byteLength !== remote.size || hash !== remote.sha256) throw Error(`download verification failed for ${remote.name}`);
  mkdirSync(dirname(localPath), {recursive: true});
  const partial = `${localPath}.part-${process.pid}`;
  try {
    await Bun.write(partial, bytes);
    renameSync(partial, localPath);
  } finally {
    if (existsSync(partial)) rmSync(partial);
  }
}

export async function remoteMain(argv: string[], opts: {tty?: boolean} = {}): Promise<number> {
  let command: RemoteCommand;
  try { command = parseRemoteArgs(argv); }
  catch (error) { process.stderr.write(`figma2pptx: ${error instanceof Error ? error.message : String(error)}\n\n${REMOTE_USAGE}`); return 2; }
  if (command.help) { process.stdout.write(REMOTE_USAGE); return 0; }
  if (command.version) { process.stdout.write(VERSION + '\n'); return 0; }

  const progress = createProgress({tty: opts.tty});
  try {
    const {origin, secret} = bridgeConfig();
    const response = await fetch(`${origin}/v1/exports`, {
      method: 'POST',
      headers: {Authorization: `Bearer ${secret}`, 'Content-Type': 'application/json'},
      body: JSON.stringify(command.request),
    });
    if (!response.ok) {
      const body = await response.json().catch(() => ({})) as {error?: string};
      throw Error(body.error || `bridge returned ${response.status}`);
    }
    if (!response.body) throw Error('bridge returned no progress stream');

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let pending = '', doneEvent: Extract<BridgeEvent, {type: 'done'}> | undefined;
    const handle = (event: BridgeEvent) => {
      if (event.type === 'accepted') progress.stage(event.stage);
      if (event.type === 'progress') event.step === event.stage ? progress.stage(event.stage) : progress.update(event.step);
      if (event.type === 'error') throw Error(event.error);
      if (event.type === 'done') doneEvent = event;
    };
    while (true) {
      const chunk = await reader.read();
      pending += decoder.decode(chunk.value, {stream: !chunk.done});
      const lines = pending.split('\n');
      pending = lines.pop() || '';
      for (const line of lines) if (line.trim()) handle(JSON.parse(line));
      if (chunk.done) break;
    }
    if (pending.trim()) handle(JSON.parse(pending));
    if (!doneEvent) throw Error('bridge ended without a completed export');

    const pptx = resolve(command.out ? (/\.pptx$/i.test(command.out) ? command.out : `${command.out}.pptx`) : join(homedir(), 'Downloads', doneEvent.outputs.pptx.name));
    const stem = pptx.replace(/\.pptx$/i, '');
    await download(origin, secret, doneEvent.outputs.pptx, pptx);
    const downloaded = [pptx];
    if (doneEvent.outputs.pdf) {
      const pdf = `${stem}.pdf`;
      await download(origin, secret, doneEvent.outputs.pdf, pdf);
      downloaded.push(pdf);
    }
    const report = `${stem}.report.json`;
    await download(origin, secret, doneEvent.outputs.report, report);
    const summary = `done: ${doneEvent.slides} slide${doneEvent.slides === 1 ? '' : 's'} from bridge -> ${downloaded.join(', ')}; report ${report}`;
    await progress.end(true, summary);
    return 0;
  } catch (error) {
    await progress.end(false, `figma2pptx: ${error instanceof Error ? error.message : String(error)}`);
    return 1;
  }
}

if (import.meta.main) process.exit(await remoteMain(process.argv.slice(2)));
