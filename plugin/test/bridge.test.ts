import {afterEach, expect, test} from 'bun:test';
import {chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {createBridgeHandler, loadOrCreateSecret} from '../bridge/server.ts';
import {BRIDGE_BIND_HOST, BRIDGE_BIND_PORT, BRIDGE_ORIGIN, parseBridgeOrigin, parseExportRequest} from '../bridge/protocol.ts';
import type {ConversionJob, ConversionRunner} from '../bridge/runner.ts';
import {createProcessRunner} from '../bridge/runner.ts';
import {loadBridgeEnvironment} from '../bridge/launch.ts';
import {assertBridgeDependencies, findExecutable} from '../bridge/preflight.ts';

const dirs: string[] = [];
afterEach(() => { while (dirs.length) rmSync(dirs.pop()!, {recursive: true, force: true}); });
const temp = () => { const dir = mkdtempSync(join(tmpdir(), 'figma2pptx-plugin-test-')); dirs.push(dir); return dir; };
const body = {fileKey: 'AbCdEfGhIjKl', frameIds: ['12:34', '12:35'], frameNames: ['Opening / frame', 'Evidence']};
const auth = (secret: string) => ({Authorization: `Bearer ${secret}`, 'Content-Type': 'application/json'});

test('validates the export request boundary', () => {
  expect(parseExportRequest(body)).toMatchObject({fileKey: body.fileKey, frames: body.frameIds, frameNames: body.frameNames, pdf: 'screen', passes: 2, embedFonts: false});
  expect(() => parseExportRequest({...body, embedFonts: true})).toThrow('font embedding is disabled');
  expect(() => parseExportRequest({...body, fileKey: '../secret'})).toThrow('fileKey is invalid');
  expect(() => parseExportRequest({...body, frameIds: ['not-an-id']})).toThrow('invalid node ID');
  expect(() => parseExportRequest({...body, frameNames: []})).toThrow('frameNames must match');
});

test('requires the per-install secret before invoking the converter', async () => {
  let calls = 0;
  const runner: ConversionRunner = async () => { calls++; throw Error('must not run'); };
  const handler = createBridgeHandler({secret: 'a'.repeat(64), outputDir: temp(), cacheDir: temp(), runner});
  const response = await handler(new Request('http://127.0.0.1/v1/exports', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(body)}));
  expect(response.status).toBe(401);
  expect(calls).toBe(0);
});

test('allows a Figma secure-context private-network preflight', async () => {
  const handler = createBridgeHandler({secret: 'a'.repeat(64), outputDir: temp(), cacheDir: temp()});
  const response = await handler(new Request('http://127.0.0.1/v1/health', {
    method: 'OPTIONS',
    headers: {
      Origin: 'null',
      'Access-Control-Request-Headers': 'authorization,content-type',
      'Access-Control-Request-Method': 'GET',
      'Access-Control-Request-Private-Network': 'true',
    },
  }));
  expect(response.status).toBe(204);
  expect(response.headers.get('Access-Control-Allow-Origin')).toBe('*');
  expect(response.headers.get('Access-Control-Allow-Headers')).toContain('Authorization');
  expect(response.headers.get('Access-Control-Allow-Private-Network')).toBe('true');
});

test('doctor runs in the authenticated bridge context and reports its checks', async () => {
  const root = temp();
  const handler = createBridgeHandler({
    secret: 'f'.repeat(64), outputDir: root, cacheDir: temp(),
    doctor: async outputDir => ({ok: true, checks: ['Automation works', 'No dialogs'], jobDir: join(outputDir, 'doctor', 'proof')}),
  });
  const denied = await handler(new Request('http://127.0.0.1/v1/doctor', {method: 'POST'}));
  expect(denied.status).toBe(401);
  const response = await handler(new Request('http://127.0.0.1/v1/doctor', {method: 'POST', headers: auth('f'.repeat(64))}));
  expect(await response.json()).toEqual({ok: true, checks: ['Automation works', 'No dialogs'], jobDir: join(root, 'doctor', 'proof')});
});

test('streams library progress, slide count, elapsed time, and output paths', async () => {
  const root = temp();
  let received: ConversionJob | undefined;
  const runner: ConversionRunner = async (job, progress) => {
    received = job;
    progress({type: 'stage', text: 'fetching 2 frames from Figma'});
    progress({type: 'step', text: 'pass 1/2: building: slide 1/2'});
    await Bun.sleep(5);
    progress({type: 'step', text: 'pass 1/2: building: slide 2/2'});
    const pdf = job.out.replace(/\.pptx$/, '.pdf'), report = job.out.replace(/\.pptx$/, '.report.json');
    writeFileSync(job.out, 'pptx bytes');
    writeFileSync(pdf, 'pdf bytes');
    writeFileSync(report, '{"ok":true}');
    return {pptx: job.out, pdf, report, slides: 2, seconds: 0.01};
  };
  const handler = createBridgeHandler({
    secret: 'b'.repeat(64), outputDir: root, cacheDir: join(root, 'cache'), runner,
    now: () => new Date('2026-09-28T08:09:10Z'), heartbeatMs: 2,
  });
  const response = await handler(new Request('http://127.0.0.1/v1/exports', {method: 'POST', headers: auth('b'.repeat(64)), body: JSON.stringify(body)}));
  expect(response.status).toBe(200);
  const events = (await response.text()).trim().split('\n').map(line => JSON.parse(line));
  expect(events[0]).toMatchObject({type: 'accepted', jobId: expect.stringMatching(/^[a-f0-9-]{36}$/), slide: 0, totalSlides: 2});
  expect(events).toContainEqual(expect.objectContaining({type: 'progress', slide: 1, totalSlides: 2}));
  expect(events).toContainEqual(expect.objectContaining({type: 'progress', slide: 2, totalSlides: 2}));
  const done = events.at(-1);
  const downloadPath = done.outputs.pptx.downloadPath;
  expect(done.type).toBe('done');
  expect(done.slides).toBe(2);
  const jobDir = join(root, '2026-09-28', 'Opening - frame - 20260928-080910000');
  expect(done.outputs.pptx).toMatchObject({name: 'Opening - frame +1 - Figma export - 20260928-080910000.pptx', remotePath: join(jobDir, 'Opening - frame +1 - Figma export - 20260928-080910000.pptx')});
  expect(done.outputs.pdf).toMatchObject({name: 'Opening - frame +1 - Figma export - 20260928-080910000.pdf', remotePath: join(jobDir, 'Opening - frame +1 - Figma export - 20260928-080910000.pdf')});
  expect(received).toMatchObject({fileKey: body.fileKey, frames: body.frameIds, frameNames: body.frameNames, cacheDir: join(root, 'cache')});

  expect(done.outputs.pptx.downloadPath).toMatch(/^\/v1\/exports\/[a-f0-9-]{36}\/files\/pptx$/);
  const unauthenticated = await handler(new Request(`http://127.0.0.1${downloadPath}`));
  expect(unauthenticated.status).toBe(401);
  const downloaded = await handler(new Request(`http://127.0.0.1${downloadPath}`, {headers: auth('b'.repeat(64))}));
  expect({status: downloaded.status, body: await downloaded.text()}).toEqual({status: 200, body: 'pptx bytes'});

  const restartedHandler = createBridgeHandler({secret: 'b'.repeat(64), outputDir: root, cacheDir: join(root, 'cache'), runner});
  const persistedStatus = await restartedHandler(new Request(`http://127.0.0.1/v1/exports/${done.jobId}`, {headers: auth('b'.repeat(64))}));
  expect(await persistedStatus.json()).toEqual(done);
  const persistedDownload = await restartedHandler(new Request(`http://127.0.0.1${downloadPath}`, {headers: auth('b'.repeat(64))}));
  expect({status: persistedDownload.status, body: await persistedDownload.text()}).toEqual({status: 200, body: 'pptx bytes'});
  expect(statSync(join(root, '.bridge-jobs.json')).mode & 0o777).toBe(0o600);
});

test('job logs expose lifecycle without file keys or design names', async () => {
  const root = temp(), logs: string[] = [];
  const runner: ConversionRunner = async (job, progress) => {
    progress({type: 'stage', text: `asking Figma for file ${job.fileKey}`});
    progress({type: 'stage', text: 'fetching 2 frames from Figma ("Private design name")'});
    const report = job.out.replace(/\.pptx$/, '.report.json');
    writeFileSync(job.out, 'pptx');
    writeFileSync(report, '{}');
    return {pptx: job.out, report, slides: 2, seconds: 0.01};
  };
  const handler = createBridgeHandler({secret: 'e'.repeat(64), outputDir: root, cacheDir: temp(), runner, log: line => logs.push(line)});
  const response = await handler(new Request('http://127.0.0.1/v1/exports', {method: 'POST', headers: auth('e'.repeat(64)), body: JSON.stringify(body)}));
  await response.text();
  expect(logs.join('\n')).toContain('accepted (2 slides requested)');
  expect(logs.join('\n')).not.toContain(body.fileKey);
  expect(logs.join('\n')).not.toContain('Private design name');
});

test('returns PowerPoint busy errors through the progress stream', async () => {
  const runner: ConversionRunner = async () => { throw Error('PowerPoint already has 1 presentation open; close it and retry (figma2pptx will not touch unrelated decks)'); };
  const handler = createBridgeHandler({secret: 'c'.repeat(64), outputDir: temp(), cacheDir: temp(), runner});
  const response = await handler(new Request('http://127.0.0.1/v1/exports', {method: 'POST', headers: auth('c'.repeat(64)), body: JSON.stringify(body)}));
  const events = (await response.text()).trim().split('\n').map(line => JSON.parse(line));
  expect(events.at(-1)).toEqual(expect.objectContaining({type: 'error', error: expect.stringContaining('will not touch unrelated decks')}));
});

test('disconnecting the client stream keeps conversion running and exposes its durable status', async () => {
  const root = temp();
  let aborted = false;
  let finish!: () => void;
  const runner: ConversionRunner = (job, _progress, signal) => new Promise((resolve, reject) => {
    signal?.addEventListener('abort', () => { aborted = true; reject(Error('canceled')); }, {once: true});
    finish = () => {
      const report = job.out.replace(/\.pptx$/, '.report.json');
      writeFileSync(job.out, 'pptx after disconnect');
      writeFileSync(report, '{}');
      resolve({pptx: job.out, report, slides: 2, seconds: 0.01});
    };
  });
  const logs: string[] = [];
  const handler = createBridgeHandler({secret: 'd'.repeat(64), outputDir: root, cacheDir: temp(), runner, log: line => logs.push(line)});
  const response = await handler(new Request('http://127.0.0.1/v1/exports', {method: 'POST', headers: auth('d'.repeat(64)), body: JSON.stringify(body)}));
  const reader = response.body!.getReader();
  const accepted = JSON.parse(new TextDecoder().decode((await reader.read()).value));
  await reader.cancel();
  expect(aborted).toBe(false);
  finish();
  await Bun.sleep(5);
  const status = await handler(new Request(`http://127.0.0.1/v1/exports/${accepted.jobId}`, {headers: auth('d'.repeat(64))}));
  expect(await status.json()).toMatchObject({type: 'done', jobId: accepted.jobId});
  expect(logs).toContainEqual(expect.stringContaining('progress stream disconnected; conversion continues'));
});

test('bounds a conversion process that never starts and names the macOS permission gate', async () => {
  const root = temp();
  let output!: ReadableStreamDefaultController<Uint8Array>;
  let exit!: (code: number) => void;
  const child = {
    stdout: new ReadableStream<Uint8Array>({start(controller) { output = controller; }}),
    stderr: new Blob([]).stream(),
    exited: new Promise<number>(resolve => { exit = resolve; }),
    exitCode: null as number | null,
    kill() { this.exitCode = 143; output.close(); exit(143); },
  };
  const runner = createProcessRunner((() => child) as unknown as typeof Bun.spawn, 5);
  await expect(runner({...parseExportRequest(body), out: join(root, 'out.pptx'), cacheDir: join(root, 'cache')}, () => {}))
    .rejects.toThrow('macOS Automation permission');
  expect(child.exitCode).toBe(143);
});

test('reads progress and completion from the CLI subprocess protocol', async () => {
  const root = temp();
  const result = {pptx: join(root, 'out.pptx'), pdf: join(root, 'out.pdf'), report: join(root, 'out.report.json'), slides: 2, seconds: 1};
  const lines = [
    {type: 'progress', event: {type: 'stage', text: 'asking Figma'}},
    {type: 'done', result},
  ].map(value => JSON.stringify(value)).join('\n') + '\n';
  const child = {
    stdout: new Blob([lines]).stream(), stderr: new Blob([]).stream(), exited: Promise.resolve(0), exitCode: 0,
    kill() {},
  };
  const progress: string[] = [];
  const runner = createProcessRunner((() => child) as unknown as typeof Bun.spawn, 50);
  await expect(runner({...parseExportRequest(body), out: result.pptx, cacheDir: join(root, 'cache')}, event => progress.push(event.text)))
    .resolves.toEqual(result);
  expect(progress).toEqual(['asking Figma']);
  expect(statSync(join(root, 'request.json')).mode & 0o777).toBe(0o600);
});

test('creates one persistent owner-only secret and separates the loopback listener from the tailnet origin', () => {
  const dir = temp();
  const first = loadOrCreateSecret(dir);
  const second = loadOrCreateSecret(dir);
  expect(first.secret).toMatch(/^[a-f0-9]{64}$/);
  expect(second).toEqual(first);
  expect(statSync(first.path).mode & 0o777).toBe(0o600);
  expect(readFileSync(first.path, 'utf8').trim()).toBe(first.secret);
  expect(BRIDGE_ORIGIN).toBe('https://bridge-host.invalid:38458');
  expect(parseBridgeOrigin('https://example.internal:38458/')).toBe('https://example.internal:38458');
  expect(() => parseBridgeOrigin('http://example.internal:38458')).toThrow('HTTPS origin');
  expect(() => parseBridgeOrigin('https://private-bridge:38458')).toThrow('fully qualified');
  expect(BRIDGE_BIND_HOST).toBe('127.0.0.1');
  expect(BRIDGE_BIND_PORT).toBe(38456);
});

test('loads only an owner-only bridge environment file', () => {
  const dir = temp(), path = join(dir, 'bridge.env');
  writeFileSync(path, 'FIGMA_TOKEN=test-token\nFIGMA2PPTX_BRIDGE_HOST=127.0.0.1\n', {mode: 0o600});
  expect(loadBridgeEnvironment(path)).toEqual({FIGMA_TOKEN: 'test-token', FIGMA2PPTX_BRIDGE_HOST: '127.0.0.1'});
  chmodSync(path, 0o644);
  expect(() => loadBridgeEnvironment(path)).toThrow('owner-only');
});

test('preflights every conversion executable against the service PATH', () => {
  const root = temp(), bin = join(root, 'bin');
  mkdirSync(bin);
  for (const name of ['magick', 'fc-list', 'osascript', 'pdfimages', 'pdftoppm']) {
    const executable = join(bin, name);
    writeFileSync(executable, '#!/bin/sh\n', {mode: 0o700});
  }
  expect(findExecutable('magick', bin)).toBe(join(bin, 'magick'));
  expect(assertBridgeDependencies(bin)).toEqual({
    magick: join(bin, 'magick'),
    'fc-list': join(bin, 'fc-list'),
    osascript: join(bin, 'osascript'),
    pdfimages: join(bin, 'pdfimages'),
    pdftoppm: join(bin, 'pdftoppm'),
  });
  rmSync(join(bin, 'magick'));
  expect(() => assertBridgeDependencies(bin)).toThrow('missing required executable: magick');
});
