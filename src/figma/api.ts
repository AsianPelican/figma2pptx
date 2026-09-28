// Read-only Figma REST client. The token comes from the environment and is never printed or written anywhere.
// FIGMA_TOKEN comes only from the environment. The tool never opens a secrets file, prints the token, or writes
// it into its cache or report.
export function readToken(env: Record<string, string | undefined> = process.env): string {
  if (env.FIGMA_TOKEN?.trim()) return env.FIGMA_TOKEN.trim();
  throw Error('no Figma token: set FIGMA_TOKEN in the environment (a personal access token with file read access)');
}

export type Target = {fileKey: string, nodeId?: string};

// A Figma URL (file, design, proto, board; branch URLs use the branch key) or a bare file key.
export function parseTarget(s: string): Target {
  if (/^[A-Za-z0-9]{10,}$/.test(s)) return {fileKey: s};
  let u: URL;
  try { u = new URL(s); } catch { throw Error(`not a Figma URL or file key: ${s}`); }
  if (!/(^|\.)figma\.com$/.test(u.hostname)) throw Error(`not a figma.com URL: ${s}`);
  const parts = u.pathname.split('/').filter(Boolean);
  const k = parts.findIndex(p => /^(file|design|proto|board|slides|deck)$/.test(p));
  if (k < 0 || !parts[k + 1]) throw Error(`no file key in ${s}`);
  const fileKey = parts[k + 2] === 'branch' && parts[k + 3] ? parts[k + 3] : parts[k + 1];
  const node = u.searchParams.get('node-id');
  return {fileKey, nodeId: node ? normalizeId(node) : undefined};
}

// Figma URLs write node ids with '-', the API with ':'.
export const normalizeId = (id: string) => id.replace(/^(\d+)-(\d+)$/, '$1:$2');
export const isNodeId = (s: string) => /^\d+[:-]\d+$/.test(s);

export type Fetch = (url: string, init?: RequestInit) => Promise<Response>;

export class FigmaClient {
  constructor(private token: string, private fetchImpl: Fetch = fetch) {}

  async get(path: string): Promise<any> {
    for (let attempt = 0; attempt < 6; attempt++) {
      const r = await this.fetchImpl('https://api.figma.com/v1/' + path, {headers: {'X-Figma-Token': this.token}});
      if (r.status === 429) {
        const wait = Math.min(60, +(r.headers.get('retry-after') || 0) || 5 * (attempt + 1));
        console.error(`  Figma rate limit, waiting ${wait}s`);
        await Bun.sleep(wait * 1000); continue;
      }
      if (r.status === 403) throw Error(`Figma ${path.split('?')[0]}: 403 forbidden (token lacks access to this file, or has expired)`);
      if (!r.ok) throw Error(`Figma ${path.split('?')[0]}: ${r.status} ${(await r.text()).slice(0, 300)}`);
      return r.json();
    }
    throw Error(`Figma ${path.split('?')[0]}: still rate limited after 6 attempts`);
  }

  // Download a rendered image from the URL /v1/images returned (a signed S3 URL; no token needed).
  async download(url: string): Promise<Uint8Array> {
    for (let attempt = 0; ; attempt++) {
      const r = await this.fetchImpl(url);
      if (r.ok) return new Uint8Array(await r.arrayBuffer());
      if (attempt >= 3) throw Error(`render download failed: ${r.status}`);
      await Bun.sleep(2000 * (attempt + 1));
    }
  }
}
