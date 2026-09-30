import {expect, test} from 'bun:test';
import {mkdtempSync, readFileSync, rmSync, writeFileSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {refreshDesktop} from '../refresh-desktop.ts';
import {pluginFixture} from './plugin-fixture.ts';

const sourceRoot = join(import.meta.dir, '..');

test('authenticated live origin is the only bundle allowed to refresh Desktop', async () => {
  const pluginRoot = pluginFixture(sourceRoot);
  const destination = mkdtempSync(join(tmpdir(), 'figma2pptx-desktop-test-'));
  const secretPath = join(destination, 'secret');
  const secret = 'a'.repeat(64);
  writeFileSync(secretPath, secret);
  let authorization = '';
  let healthUrl = '';
  try {
    await expect(refreshDesktop({
      pluginRoot,
      destination,
      origin: 'https://bridge.example:38458',
      secretPath,
      fetcher: async () => { throw Error('test origin must be rejected before health'); },
    })).rejects.toThrow('configured live bridge origin');
    expect(() => readFileSync(join(destination, 'manifest.json'))).toThrow();

    await expect(refreshDesktop({
      pluginRoot,
      destination,
      origin: 'https://bridge.private.example.net:38458',
      secretPath,
      fetcher: async () => new Response('unauthorized', {status: 401}),
    })).rejects.toThrow('health check failed with 401');
    expect(() => readFileSync(join(destination, 'manifest.json'))).toThrow();

    await refreshDesktop({
      pluginRoot,
      destination,
      origin: 'https://bridge.private.example.net:38458',
      secretPath,
      fetcher: async (url, init) => {
        healthUrl = String(url);
        authorization = new Headers(init?.headers).get('Authorization') ?? '';
        return Response.json({ok: true, busy: false});
      },
    });
    const manifest = JSON.parse(readFileSync(join(destination, 'manifest.json'), 'utf8'));
    const ui = readFileSync(join(destination, 'dist', 'ui.html'), 'utf8');
    expect(manifest.networkAccess.devAllowedDomains).toEqual(['https://bridge.private.example.net:38458']);
    expect(ui).toContain("const BRIDGE = 'https://bridge.private.example.net:38458'");
    expect(ui).not.toContain('bridge.example');
    expect(healthUrl).toBe('https://bridge.private.example.net:38458/v1/health');
    expect(authorization).toBe(`Bearer ${secret}`);
  } finally {
    rmSync(pluginRoot, {recursive: true, force: true});
    rmSync(destination, {recursive: true, force: true});
  }
});
