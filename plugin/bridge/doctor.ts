import {closeSync, openSync, readFileSync, readSync} from 'node:fs';
import {homedir} from 'node:os';
import {join} from 'node:path';
import {BRIDGE_BIND_HOST, BRIDGE_BIND_PORT, type DoctorResult} from './protocol.ts';
import {runtimePaths} from './runtime.ts';

export function sshContextHasFullDiskAccess(
  tcc = join(homedir(), 'Library', 'Application Support', 'com.apple.TCC', 'TCC.db'),
): boolean {
  let fd: number | undefined;
  try {
    fd = openSync(tcc, 'r');
    readSync(fd, Buffer.alloc(1), 0, 1, 0);
    return true;
  } catch {
    return false;
  } finally {
    if (fd !== undefined) closeSync(fd);
  }
}

export async function doctorMain(): Promise<number> {
  try {
    const sshHasFullDiskAccess = sshContextHasFullDiskAccess();
    const secretPath = join(runtimePaths().config, 'bridge-secret');
    const secret = readFileSync(secretPath, 'utf8').trim();
    if (!/^[a-f0-9]{64}$/.test(secret)) throw Error(`Bridge secret is invalid at ${secretPath}`);
    const response = await fetch(`http://${BRIDGE_BIND_HOST}:${BRIDGE_BIND_PORT}/v1/doctor`, {
      method: 'POST',
      headers: {Authorization: `Bearer ${secret}`},
    });
    const body = await response.json() as DoctorResult | {error?: string};
    if (!response.ok || !('ok' in body)) throw Error('error' in body && body.error ? body.error : `doctor returned ${response.status}`);
    process.stdout.write('figma2pptx doctor: launcher permission checks passed\n');
    process.stdout.write(sshHasFullDiskAccess
      ? '  info  SSH context can read protected app data (not required for exports)\n'
      : '  info  SSH context cannot read protected app data (expected with Tailscale SSH; exports run through the launcher)\n');
    for (const check of body.checks) process.stdout.write(`  ok  ${check}\n`);
    process.stdout.write(`  evidence  ${body.jobDir}\n`);
    return 0;
  } catch (error) {
    process.stderr.write(`figma2pptx doctor: ${error instanceof Error ? error.message : String(error)}\n`);
    return 1;
  }
}

if (import.meta.main) process.exit(await doctorMain());
