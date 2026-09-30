import {chmodSync, writeFileSync} from 'node:fs';
import {dirname, resolve} from 'node:path';
import type {ProgressEvent} from '../../src/index.ts';
import type {ExportRequest, ExportResult} from './protocol.ts';

export type ConversionJob = ExportRequest & {out: string; cacheDir: string};
export type ConversionRunner = (job: ConversionJob, onProgress: (event: ProgressEvent) => void, signal?: AbortSignal) => Promise<ExportResult>;

type ChildMessage =
  | {type: 'progress'; event: ProgressEvent}
  | {type: 'done'; result: ExportResult}
  | {type: 'error'; error: string};

export const PROCESS_START_TIMEOUT_MS = 15000;
const WORKER = resolve(import.meta.dir, 'subprocess-worker.ts');

export function createProcessRunner(spawn: typeof Bun.spawn = Bun.spawn, startupTimeoutMs = PROCESS_START_TIMEOUT_MS): ConversionRunner {
  return async (job, onProgress, signal) => {
    const requestPath = resolve(dirname(job.out), 'request.json');
    writeFileSync(requestPath, JSON.stringify(job), {mode: 0o600});
    chmodSync(requestPath, 0o600);
    const child = spawn([process.execPath, WORKER, requestPath], {
      cwd: resolve(import.meta.dir, '..', '..'),
      env: process.env,
      stdout: 'pipe',
      stderr: 'pipe',
    });
    let startupError: Error | undefined;
    let started = false;
    let result: ExportResult | undefined;
    let childError: string | undefined;
    const startup = setTimeout(() => {
      startupError = Error(`Conversion process did not start within ${Math.round(startupTimeoutMs / 1000)} seconds. It may be waiting for macOS Automation permission on the bridge Mac; allow Bun to control System Events and Microsoft PowerPoint, then retry.`);
      child.kill();
    }, startupTimeoutMs);
    const abort = () => child.kill();
    signal?.addEventListener('abort', abort, {once: true});
    const stderrPromise = new Response(child.stderr).text();
    try {
      const reader = child.stdout.pipeThrough(new TextDecoderStream()).getReader();
      let pending = '';
      while (true) {
        const {done, value} = await reader.read();
        pending += value ?? '';
        const lines = pending.split('\n');
        pending = lines.pop() ?? '';
        for (const line of lines) {
          if (!line.trim()) continue;
          const message = JSON.parse(line) as ChildMessage;
          if (!started) { started = true; clearTimeout(startup); }
          if (message.type === 'progress') onProgress(message.event);
          if (message.type === 'done') result = message.result;
          if (message.type === 'error') childError = message.error;
        }
        if (done) break;
      }
      const code = await child.exited;
      const stderr = (await stderrPromise).trim();
      if (signal?.aborted) throw Error('Export canceled by the client.');
      if (startupError) throw startupError;
      if (childError) throw Error(childError);
      if (code !== 0 || !result) throw Error(`conversion process failed${stderr ? `: ${stderr.split('\n').pop()}` : ''}`);
      return result;
    } finally {
      clearTimeout(startup);
      signal?.removeEventListener('abort', abort);
      if (child.exitCode === null) child.kill();
    }
  };
}

export const processRunner = createProcessRunner();
