// Figma access without the network: URL parsing, the token, retries, the versioned cache, frame selection.
import {test, expect, describe} from 'bun:test';
import {readFileSync, existsSync, readdirSync, mkdirSync, rmSync, mkdtempSync} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {parseTarget, readToken, FigmaClient, type Fetch} from '../src/figma/api';
import {FigmaFile, resolveFrames} from '../src/figma/source';
import {FIXTURE_CACHE, FIXTURE_KEY} from './helpers';

describe('parseTarget', () => {
  test('file keys and every Figma URL shape', () => {
    expect(parseTarget('AbCdEf0123456789xyzABC')).toEqual({fileKey: 'AbCdEf0123456789xyzABC'});
    expect(parseTarget('https://www.figma.com/design/AbCdEf0123/My-Deck?node-id=12-345&t=x')).toEqual({fileKey: 'AbCdEf0123', nodeId: '12:345'});
    expect(parseTarget('https://www.figma.com/file/AbCdEf0123/My-Deck')).toEqual({fileKey: 'AbCdEf0123', nodeId: undefined});
    expect(parseTarget('https://figma.com/design/AbCdEf0123/branch/BrAnCh4567/x?node-id=0-1')).toEqual({fileKey: 'BrAnCh4567', nodeId: '0:1'});
    expect(() => parseTarget('https://example.com/design/AbCdEf0123')).toThrow('not a figma.com URL');
    expect(() => parseTarget('short')).toThrow('not a Figma URL or file key');
  });
});

test('readToken prefers FIGMA_TOKEN from the environment', () => {
  expect(readToken({FIGMA_TOKEN: '  TEST_TOKEN_VALUE  '})).toBe('TEST_TOKEN_VALUE');
  expect(() => readToken({})).toThrow('set FIGMA_TOKEN in the environment');
});

describe('FigmaClient', () => {
  const secret = 'TEST_TOKEN_DO_NOT_PRINT';
  test('sends the token as a header, retries 429 after Retry-After', async () => {
    const seen: string[] = []; let calls = 0;
    const fetch: Fetch = async (url, init) => {
      seen.push((init?.headers as any)['X-Figma-Token']);
      return ++calls === 1 ? new Response('slow down', {status: 429, headers: {'retry-after': '0.01'}}) : Response.json({ok: true});
    };
    expect(await new FigmaClient(secret, fetch).get('files/K')).toEqual({ok: true});
    expect(calls).toBe(2);
    expect(seen).toEqual([secret, secret]);
  });
  test('errors name the endpoint and status, never the token or the query', async () => {
    const fetch: Fetch = async () => new Response('nope', {status: 403});
    const err = await new FigmaClient(secret, fetch).get('files/K/nodes?ids=1:2').catch(e => e);
    expect(err.message).toContain('files/K/nodes: 403');
    expect(err.message).not.toContain(secret);
    expect(err.message).not.toContain('ids=');
  });
});

// A fake Figma API serving the fixture deck, counting requests.
function fakeFigma(version: string) {
  const dir = join(FIXTURE_CACHE, FIXTURE_KEY, '1001');
  const meta = JSON.parse(readFileSync(join(dir, 'file.json'), 'utf8'));
  const requests: string[] = [];
  const fetch: Fetch = async url => {
    requests.push(url.replace('https://api.figma.com/v1/', ''));
    const u = new URL(url);
    if (u.hostname === 'renders.test') return new Response(readFileSync(join(dir, u.pathname.slice(1))));
    if (/\/files\/[^/]+$/.test(u.pathname)) return Response.json({name: meta.name, version, lastModified: meta.lastModified, document: {children: meta.pages.map((p: any) => ({id: p.id, name: p.name, type: 'CANVAS', children: p.frames.map((f: any) => ({...f, absoluteBoundingBox: {width: f.width, height: f.height}}))}))}});
    if (u.pathname.endsWith('/nodes')) return Response.json({nodes: Object.fromEntries(u.searchParams.get('ids')!.split(',').map(id => [id, {document: JSON.parse(readFileSync(join(dir, 'nodes', id.replace(':', '-') + '.json'), 'utf8'))}]))});
    if (u.pathname.startsWith('/v1/images/')) {
      const fmt = u.searchParams.get('format')!;
      return Response.json({images: Object.fromEntries(u.searchParams.get('ids')!.split(',').map(id => [id, `https://renders.test/${fmt === 'svg' ? 'svg' : 'raster'}/${id.replace(':', '-')}${fmt === 'svg' ? '.svg' : '@2.png'}`]))});
    }
    return new Response('not found', {status: 404});
  };
  return {fetch, requests};
}

describe('FigmaFile cache', () => {
  test('first run fetches, a warm run of the same version asks only for the version', async () => {
    const root = mkdtempSync(join(tmpdir(), 'figma2pptx-cache-'));
    try {
      const api = fakeFigma('777');
      const f = await FigmaFile.open(FIXTURE_KEY, root, new FigmaClient('t', api.fetch));
      await f.fetchFrames(['1:2', '1:40']);
      await f.ensureRasters(['1:10'], 2);
      expect(existsSync(join(root, FIXTURE_KEY, '777', 'svg', '1-40.svg'))).toBe(true);
      expect(f.document('1:2').name).toBe('Cover');
      const cold = api.requests.length;
      const warm = await FigmaFile.open(FIXTURE_KEY, root, new FigmaClient('t', api.fetch));
      await warm.fetchFrames(['1:2', '1:40']);
      await warm.ensureRasters(['1:10'], 2);
      expect(api.requests.length - cold).toBe(1);
      expect(api.requests.at(-1)).toBe(`files/${FIXTURE_KEY}?depth=2`);
    } finally { rmSync(root, {recursive: true, force: true}); }
  });

  test('a new version replaces the cached one; offline uses what is cached', async () => {
    const root = mkdtempSync(join(tmpdir(), 'figma2pptx-cache-'));
    try {
      await FigmaFile.open(FIXTURE_KEY, root, new FigmaClient('t', fakeFigma('1').fetch));
      await FigmaFile.open(FIXTURE_KEY, root, new FigmaClient('t', fakeFigma('2').fetch));
      expect(readdirSync(join(root, FIXTURE_KEY))).toEqual(['2']);
      const off = await FigmaFile.open(FIXTURE_KEY, root, null);
      expect(off.meta.version).toBe('2');
      await expect(off.fetchFrames(['1:2'])).rejects.toThrow('offline');
      await expect(FigmaFile.open('NothingCached000', root, null)).rejects.toThrow('nothing cached');
    } finally { rmSync(root, {recursive: true, force: true}); }
  });
});

describe('resolveFrames', () => {
  const meta = JSON.parse(readFileSync(join(FIXTURE_CACHE, FIXTURE_KEY, '1001', 'file.json'), 'utf8'));
  test('ids in either form, names, pages, a URL node-id', () => {
    expect(resolveFrames(meta, ['1-40', '1:2'], {})).toEqual(['1:40', '1:2']);
    expect(resolveFrames(meta, ['Cover'], {})).toEqual(['1:2']);
    expect(resolveFrames(meta, [], {page: 'Slides'})).toEqual(['1:2', '1:40']);
    expect(resolveFrames(meta, ['Details'], {page: 'Scratch'})).toEqual(['2:1']);
    expect(resolveFrames(meta, [], {nodeId: '0:1'})).toEqual(['1:2', '1:40']);
    expect(resolveFrames(meta, [], {nodeId: '1:2'})).toEqual(['1:2']);
  });
  test('ambiguous and unknown names fail with the choices', () => {
    expect(() => resolveFrames(meta, ['Details'], {})).toThrow('ambiguous: 1:40 on "Slides", 2:1 on "Scratch"');
    expect(() => resolveFrames(meta, ['Nope'], {})).toThrow('no frame named "Nope"');
    expect(() => resolveFrames(meta, [], {})).toThrow('page "Slides" (0:1): 2 frames');
  });
});
