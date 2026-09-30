import {chmodSync, existsSync, mkdirSync, readFileSync, renameSync, writeFileSync} from 'node:fs';
import {basename, join, resolve, sep} from 'node:path';
import {createHash, randomBytes, randomUUID, timingSafeEqual} from 'node:crypto';
import type {ProgressEvent} from '../../src/index.ts';
import {
  BRIDGE_BIND_HOST,
  BRIDGE_BIND_PORT,
  BRIDGE_PORT,
  parseExportRequest,
  slideFromStep,
  type BridgeEvent,
  type DownloadOutput,
  type DoctorResult,
  type ExportRequest,
  type OutputKind,
} from './protocol.ts';
import {processRunner, type ConversionRunner} from './runner.ts';
import {processDoctor, type DoctorRunner} from './doctor-runner.ts';
import {runtimePaths} from './runtime.ts';

const encoder = new TextEncoder();
const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Authorization, Content-Type',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Private-Network': 'true',
  'Cache-Control': 'no-store',
};

const json = (value: unknown, status = 200) => Response.json(value, {status, headers: corsHeaders});

function authorized(request: Request, secret: string): boolean {
  const value = request.headers.get('authorization') ?? '';
  const expected = `Bearer ${secret}`;
  const a = Buffer.from(value), b = Buffer.from(expected);
  return a.length === b.length && timingSafeEqual(a, b);
}

function safeFilePart(value: string): string {
  return value.replace(/[\/\\:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim().slice(0, 80) || 'Figma selection';
}

function timestamp(date: Date): string {
  const digits = date.toISOString().replace(/\D/g, '').slice(0, 17);
  return `${digits.slice(0, 8)}-${digits.slice(8)}`;
}

function safeProgressLog(text: string): string {
  return text
    .replace(/\s+for file\s+.+$/i, '')
    .replace(/\s+\(".*"\)$/, '');
}

function output(basePath: string, jobId: string, kind: OutputKind): DownloadOutput {
  const bytes = readFileSync(basePath);
  return {
    name: basename(basePath),
    downloadPath: `/v1/exports/${jobId}/files/${kind}`,
    remotePath: basePath,
    size: bytes.byteLength,
    sha256: createHash('sha256').update(bytes).digest('hex'),
  };
}

type JobIndex = {version: 1; jobs: Record<string, BridgeEvent>};

class JobStore {
  private readonly path: string;
  private readonly jobs: Record<string, BridgeEvent>;

  constructor(private readonly outputDir: string, private readonly log: (line: string) => void) {
    this.path = join(outputDir, '.bridge-jobs.json');
    mkdirSync(outputDir, {recursive: true, mode: 0o700});
    this.jobs = this.load();
    let recovered = false;
    for (const [jobId, event] of Object.entries(this.jobs)) {
      if (event.type === 'accepted' || event.type === 'progress') {
        this.jobs[jobId] = {type: 'error', jobId, error: 'The bridge restarted before this export completed. Start a new export.', elapsed: event.elapsed};
        recovered = true;
      }
    }
    if (recovered) this.persist();
  }

  get(jobId: string): BridgeEvent | undefined { return this.jobs[jobId]; }

  set(event: BridgeEvent): void {
    this.jobs[event.jobId] = event;
    this.persist();
  }

  output(jobId: string, kind: OutputKind): string | undefined {
    const event = this.jobs[jobId];
    if (event?.type !== 'done') return undefined;
    const path = event.outputs[kind]?.remotePath;
    if (!path) return undefined;
    const root = resolve(this.outputDir) + sep;
    const resolved = resolve(path);
    return resolved.startsWith(root) ? resolved : undefined;
  }

  private load(): Record<string, BridgeEvent> {
    if (!existsSync(this.path)) return {};
    try {
      const parsed = JSON.parse(readFileSync(this.path, 'utf8')) as JobIndex;
      if (parsed.version !== 1 || !parsed.jobs || typeof parsed.jobs !== 'object') throw Error('unsupported format');
      return parsed.jobs;
    } catch (error) {
      this.log(`job index unreadable; starting empty: ${error instanceof Error ? error.message : String(error)}`);
      return {};
    }
  }

  private persist(): void {
    const temp = `${this.path}.tmp-${process.pid}`;
    writeFileSync(temp, JSON.stringify({version: 1, jobs: this.jobs}, null, 2) + '\n', {mode: 0o600});
    renameSync(temp, this.path);
    chmodSync(this.path, 0o600);
  }
}

export type BridgeOptions = {
  secret: string;
  outputDir: string;
  cacheDir: string;
  runner?: ConversionRunner;
  now?: () => Date;
  heartbeatMs?: number;
  log?: (line: string) => void;
  doctor?: DoctorRunner;
};

export function createBridgeHandler(options: BridgeOptions): (request: Request) => Promise<Response> {
  const runner = options.runner ?? processRunner;
  const now = options.now ?? (() => new Date());
  const log = options.log ?? (() => {});
  const doctor = options.doctor ?? processDoctor;
  const jobs = new JobStore(options.outputDir, log);
  let active = false;

  return async request => {
    const url = new URL(request.url);
    if (request.method === 'OPTIONS') return new Response(null, {status: 204, headers: corsHeaders});
    if (!authorized(request, options.secret)) return json({error: 'Bridge secret is incorrect.'}, 401);
    log(`request ${request.method} ${url.pathname}`);

    if (url.pathname === '/v1/health' && request.method === 'GET') return json({ok: true, busy: active});

    if (url.pathname === '/v1/doctor' && request.method === 'POST') {
      if (active) return json({error: 'Another export is already running. Doctor requires an idle bridge.'}, 409);
      active = true;
      log('doctor started');
      try {
        const result: DoctorResult = await doctor(options.outputDir);
        log('doctor passed');
        return json(result);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        log(`doctor failed: ${message}`);
        return json({error: message}, 503);
      } finally {
        active = false;
      }
    }

    const status = /^\/v1\/exports\/([a-f0-9-]{36})$/.exec(url.pathname);
    if (status && request.method === 'GET') {
      const event = jobs.get(status[1]);
      if (!event) return json({error: `Export job ${status[1]} was not found.`}, 404);
      return json(event);
    }

    const download = /^\/v1\/exports\/([a-f0-9-]{36})\/files\/(pptx|pdf|report)$/.exec(url.pathname);
    if (download && request.method === 'GET') {
      const path = jobs.output(download[1], download[2] as OutputKind);
      if (!path || !existsSync(path)) return json({error: `The ${download[2]} output for job ${download[1]} is not available on the bridge.`}, 404);
      log(`job ${download[1]} download ${download[2]}`);
      return new Response(Bun.file(path), {
        headers: {
          ...corsHeaders,
          'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(basename(path))}`,
        },
      });
    }

    if (url.pathname !== '/v1/exports' || request.method !== 'POST') return json({error: 'Not found.'}, 404);
    if (active) return json({error: 'Another export is already running. Wait for it to finish, then retry.'}, 409);

    let body: ExportRequest;
    try { body = parseExportRequest(await request.json()); }
    catch (error) { return json({error: error instanceof Error ? error.message : String(error)}, 400); }

    active = true;
    const jobId = randomUUID();
    const started = Date.now();
    const elapsed = () => +((Date.now() - started) / 1000).toFixed(1);
    let stage = 'Starting export';
    let step = `Preparing ${body.frames.length || 'page'}${body.frames.length === 1 ? ' slide' : ' slides'}`;
    let slide = 0;
    let totalSlides = body.frames.length;
    const requestedName = body.outputName?.replace(/\.pptx$/i, '');
    const firstLabel = body.frameNames[0] ?? body.page ?? 'Figma export';
    const baseName = requestedName || `${safeFilePart(firstLabel)}${body.frameNames.length > 1 ? ` +${body.frameNames.length - 1}` : ''} - Figma export - ${timestamp(now())}`;
    const jobDir = join(options.outputDir, now().toISOString().slice(0, 10), `${safeFilePart(firstLabel)} - ${timestamp(now())}`);
    mkdirSync(jobDir, {recursive: true, mode: 0o700});
    const out = join(jobDir, `${safeFilePart(baseName)}.pptx`);
    const accepted: BridgeEvent = {type: 'accepted', jobId, stage, step, slide, totalSlides, elapsed: elapsed()};
    jobs.set(accepted);
    log(`job ${jobId} accepted (${body.frames.length || 'page'} slides requested)`);

    let disconnect = () => {};
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        const aborter = new AbortController();
        let closed = false;
        const send = (event: BridgeEvent) => { if (!closed) controller.enqueue(encoder.encode(JSON.stringify(event) + '\n')); };
        send(accepted);
        const current = (): BridgeEvent => ({type: 'progress', jobId, stage, step, slide, totalSlides, elapsed: elapsed()});
        const heartbeat = setInterval(() => {
          const event = current();
          jobs.set(event);
          send(event);
        }, options.heartbeatMs ?? 1000);
        disconnect = () => {
          closed = true;
          clearInterval(heartbeat);
          log(`job ${jobId} progress stream disconnected; conversion continues`);
        };
        const progress = (event: ProgressEvent) => {
          if (event.type === 'stage') stage = event.text;
          step = event.text;
          const count = slideFromStep(event.text);
          if (count) { slide = count.slide; totalSlides = count.totalSlides; }
          const update = current();
          jobs.set(update);
          send(update);
          if (event.type === 'stage' || count) log(`job ${jobId} progress: ${safeProgressLog(event.text)}`);
        };
        void runner({...body, out, cacheDir: options.cacheDir}, progress, aborter.signal)
          .then(result => {
            const files = {pptx: result.pptx, report: result.report, ...(result.pdf ? {pdf: result.pdf} : {})};
            for (const path of Object.values(files)) if (!existsSync(path)) throw Error(`converter reported an output that does not exist: ${basename(path)}`);
            clearInterval(heartbeat);
            const done: BridgeEvent = {
              type: 'done', jobId,
              outputs: {
                pptx: output(result.pptx, jobId, 'pptx'),
                ...(result.pdf ? {pdf: output(result.pdf, jobId, 'pdf')} : {}),
                report: output(result.report, jobId, 'report'),
              },
              slides: result.slides,
              elapsed: elapsed(),
            };
            jobs.set(done);
            log(`job ${jobId} done (${result.slides} slides, ${done.elapsed}s)`);
            send(done);
          })
          .catch(error => {
            clearInterval(heartbeat);
            const failed: BridgeEvent = {type: 'error', jobId, error: error instanceof Error ? error.message : String(error), elapsed: elapsed()};
            jobs.set(failed);
            log(`job ${jobId} error: ${failed.error}`);
            send(failed);
          })
          .finally(() => { active = false; if (!closed) { closed = true; controller.close(); } });
      },
      cancel() { disconnect(); },
    });
    return new Response(stream, {headers: {...corsHeaders, 'Content-Type': 'application/x-ndjson; charset=utf-8'}});
  };
}

export function loadOrCreateSecret(configDir = runtimePaths().config): {secret: string; path: string} {
  mkdirSync(configDir, {recursive: true, mode: 0o700});
  const path = join(configDir, 'bridge-secret');
  if (!existsSync(path)) writeFileSync(path, randomBytes(32).toString('hex') + '\n', {mode: 0o600});
  chmodSync(path, 0o600);
  const secret = readFileSync(path, 'utf8').trim();
  if (!/^[a-f0-9]{64}$/.test(secret)) throw Error(`invalid bridge secret at ${path}; replace it with 32 random bytes encoded as hex`);
  return {secret, path};
}

export function startBridge(): ReturnType<typeof Bun.serve> {
  if (!process.env.FIGMA_TOKEN?.trim()) throw Error('FIGMA_TOKEN is not set. Put it in the owner-only bridge environment file.');
  const host = process.env.FIGMA2PPTX_BRIDGE_HOST?.trim() || BRIDGE_BIND_HOST;
  if (host !== BRIDGE_BIND_HOST) throw Error(`bridge listener must stay on ${BRIDGE_BIND_HOST}; the private-network proxy owns port ${BRIDGE_PORT}`);
  const paths = runtimePaths();
  const config = loadOrCreateSecret(paths.config);
  for (const path of [paths.cache, paths.jobs, paths.logs, paths.temp]) mkdirSync(path, {recursive: true, mode: 0o700});
  process.env.TMPDIR = paths.temp;
  const server = Bun.serve({
    hostname: host,
    port: BRIDGE_BIND_PORT,
    fetch: createBridgeHandler({
      secret: config.secret,
      outputDir: paths.jobs,
      cacheDir: paths.cache,
      log: line => console.log(`[bridge] ${new Date().toISOString()} ${line}`),
    }),
  });
  console.log(`figma2pptx bridge listening on http://${host}:${server.port}`);
  console.log(`Bridge secret: ${config.path} (not printed)`);
  console.log(`Runtime: ${paths.root}`);
  return server;
}

if (import.meta.main) {
  try { startBridge(); }
  catch (error) { console.error(`figma2pptx bridge: ${error instanceof Error ? error.message : String(error)}`); process.exit(1); }
}
