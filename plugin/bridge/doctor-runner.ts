import {resolve} from 'node:path';
import type {DoctorResult} from './protocol.ts';

export type DoctorRunner = (outputDir: string) => Promise<DoctorResult>;

const WORKER = resolve(import.meta.dir, 'doctor-worker.ts');

export const processDoctor: DoctorRunner = async outputDir => {
  const child = Bun.spawn([process.execPath, WORKER, outputDir], {
    cwd: resolve(import.meta.dir, '..', '..'),
    env: process.env,
    stdout: 'pipe',
    stderr: 'pipe',
  });
  const [stdout, stderr, code] = await Promise.all([
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
    child.exited,
  ]);
  if (code !== 0) throw Error(stderr.trim().split('\n').pop() || 'doctor worker failed');
  const result = JSON.parse(stdout) as DoctorResult;
  if (!result.ok || !Array.isArray(result.checks)) throw Error('doctor worker returned an invalid result');
  return result;
};
