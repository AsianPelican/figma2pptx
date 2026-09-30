// A Figma file as the converter reads it, through an on-disk cache.
//
// Cache layout: <root>/<fileKey>/<version>/
//   file.json          pages and their top-level frames (sections expanded), with the file's version
//   nodes/<id>.json    one frame's node tree
//   svg/<id>.svg       one frame's SVG export (embedded bitmaps stripped: they are never read)
//   outlined-text/<id>.svg and <id>@<s>.png  tight outlined text plus its PNG fallback
//   raster/<id>@<s>.png  renders of nodes that become pictures, plus derived slide copies and blur plates
//   figpdf/<id>.pdf    Figma's own PDF of a frame (benchmark reference only)
// Every run asks Figma for the file's current version (one small request), so a changed file is never served
// stale; older versions of the same file are deleted, which keeps the cache to one copy per file.
import {readFileSync, writeFileSync, existsSync, mkdirSync, readdirSync, rmSync, statSync} from 'node:fs';
import {join} from 'node:path';
import type {FigmaClient} from './api';
import {isNodeId, normalizeId} from './api';
import type {FrameSource} from '../convert/build';

export type FrameInfo = {id: string, name: string, type: string, visible: boolean, width: number, height: number};
export type PageInfo = {id: string, name: string, frames: FrameInfo[]};
export type FileMeta = {fileKey: string, name: string, version: string, lastModified: string, pages: PageInfo[]};

const SLIDE_TYPES = /^(FRAME|COMPONENT|COMPONENT_SET|INSTANCE|GROUP)$/;
const fileName = (id: string) => id.replace(/:/g, '-').replace(/;/g, '_');
const chunks = <T>(a: T[], n: number) => Array.from({length: Math.ceil(a.length / n)}, (_, i) => a.slice(i * n, i * n + n));

export class FigmaFile implements FrameSource {
  private docs = new Map<string, any>();
  private constructor(readonly meta: FileMeta, readonly dir: string, private client: FigmaClient | null) {}

  // Resolve the file's current version and open its cache. Offline: use the newest cached version, no network.
  static async open(fileKey: string, root: string, client: FigmaClient | null): Promise<FigmaFile> {
    const base = join(root, fileKey);
    if (!client) {
      const versions = existsSync(base) ? readdirSync(base).filter(v => existsSync(join(base, v, 'file.json'))) : [];
      if (!versions.length) throw Error(`offline: nothing cached for ${fileKey} under ${root}`);
      versions.sort((a, b) => statSync(join(base, b, 'file.json')).mtimeMs - statSync(join(base, a, 'file.json')).mtimeMs);
      const dir = join(base, versions[0]);
      return new FigmaFile(JSON.parse(readFileSync(join(dir, 'file.json'), 'utf8')), dir, null);
    }
    const f = await client.get(`files/${fileKey}?depth=2`);
    const dir = join(base, String(f.version));
    let meta: FileMeta;
    if (existsSync(join(dir, 'file.json'))) meta = JSON.parse(readFileSync(join(dir, 'file.json'), 'utf8'));
    else {
      const frame = (n: any): FrameInfo => ({id: n.id, name: n.name, type: n.type, visible: n.visible !== false, width: n.absoluteBoundingBox?.width ?? 0, height: n.absoluteBoundingBox?.height ?? 0});
      // Sections are containers on the canvas: their frames count as the page's frames, in canvas order.
      const sectionIds: string[] = f.document.children.flatMap((p: any) => (p.children || []).filter((c: any) => c.type === 'SECTION').map((c: any) => c.id));
      const sections = new Map<string, any>();
      for (const ids of chunks(sectionIds, 50)) { const j = await client.get(`files/${fileKey}/nodes?ids=${ids.join(',')}&depth=1`); for (const id of ids) if (j.nodes[id]) sections.set(id, j.nodes[id].document); }
      const flatten = (c: any): FrameInfo[] => c.type === 'SECTION' ? (sections.get(c.id)?.children || []).flatMap(flatten) : SLIDE_TYPES.test(c.type) ? [frame(c)] : [];
      meta = {fileKey, name: f.name, version: String(f.version), lastModified: f.lastModified, pages: f.document.children.map((p: any) => ({id: p.id, name: p.name, frames: (p.children || []).flatMap(flatten)}))};
      mkdirSync(dir, {recursive: true});
      writeFileSync(join(dir, 'file.json'), JSON.stringify(meta));
    }
    if (existsSync(base)) for (const v of readdirSync(base)) if (v !== String(f.version)) rmSync(join(base, v), {recursive: true, force: true});
    return new FigmaFile(meta, dir, client);
  }

  private need(): FigmaClient {
    if (!this.client) throw Error('offline: this needs Figma data that is not cached; run once without --offline');
    return this.client;
  }

  // Node trees and SVG exports of the frames, fetched once per version.
  async fetchFrames(ids: string[]) {
    const key = this.meta.fileKey;
    const nodesDir = join(this.dir, 'nodes'), svgDir = join(this.dir, 'svg');
    mkdirSync(nodesDir, {recursive: true}); mkdirSync(svgDir, {recursive: true});
    const missingNodes = ids.filter(id => !existsSync(join(nodesDir, fileName(id) + '.json')));
    for (const chunk of chunks(missingNodes, 20)) {
      const j = await this.need().get(`files/${key}/nodes?ids=${chunk.join(',')}`);
      for (const id of chunk) {
        if (!j.nodes[id]) throw Error(`node ${id} not found in file ${key}`);
        writeFileSync(join(nodesDir, fileName(id) + '.json'), JSON.stringify(j.nodes[id].document));
      }
    }
    const missingSvg = ids.filter(id => !existsSync(join(svgDir, fileName(id) + '.svg')));
    for (const chunk of chunks(missingSvg, 4)) {
      const j = await this.need().get(`images/${key}?ids=${chunk.join(',')}&format=svg&svg_outline_text=false&svg_include_node_id=true&svg_simplify_stroke=false`);
      await Promise.all(chunk.map(async id => {
        if (!j.images[id]) throw Error(`Figma could not export ${id} as SVG`);
        const s = new TextDecoder().decode(await this.need().download(j.images[id]));
        writeFileSync(join(svgDir, fileName(id) + '.svg'), s.replace(/xlink:href="data:[^"]{200,}"/g, 'xlink:href="data:stripped"').replace(/href="data:[^"]{200,}"/g, 'href="data:stripped"'));
      }));
    }
  }

  document(id: string) {
    if (!this.docs.has(id)) this.docs.set(id, JSON.parse(readFileSync(join(this.dir, 'nodes', fileName(id) + '.json'), 'utf8')));
    return this.docs.get(id);
  }
  svg(id: string) { return readFileSync(join(this.dir, 'svg', fileName(id) + '.svg'), 'utf8'); }

  rasterFile(id: string, scale: number) { const p = join(this.dir, 'raster', `${fileName(id)}@${scale}`); return existsSync(p + '.jpg') ? p + '.jpg' : p + '.png'; }
  outlinedTextFile(id: string) { return join(this.dir, 'outlined-text', fileName(id) + '.svg'); }
  outlinedTextFallbackFile(id: string, scale: number) { return join(this.dir, 'outlined-text', `${fileName(id)}@${scale}.png`); }
  platePath(id: string, scale: number) { return join(this.dir, 'raster', `plate-${fileName(id)}@${scale}.png`); }

  async ensureOutlinedText(ids: string[], scale: number) {
    if (!ids.length) return;
    const dir = join(this.dir, 'outlined-text'); mkdirSync(dir, {recursive: true});
    const need = ids.filter(id => !existsSync(this.outlinedTextFile(id)));
    for (const chunk of chunks(need, 4)) {
      const j = await this.need().get(`images/${this.meta.fileKey}?ids=${chunk.join(',')}&format=svg&svg_outline_text=true&svg_include_node_id=true&svg_simplify_stroke=false`);
      await Promise.all(chunk.map(async id => {
        if (!j.images[id]) throw Error(`Figma could not outline text node ${id}`);
        const s = new TextDecoder().decode(await this.need().download(j.images[id]));
        writeFileSync(this.outlinedTextFile(id), s.replace(/xlink:href="data:[^"]{200,}"/g, 'xlink:href="data:stripped"').replace(/href="data:[^"]{200,}"/g, 'href="data:stripped"'));
      }));
    }
    const needPng = ids.filter(id => !existsSync(this.outlinedTextFallbackFile(id, scale)));
    for (const chunk of chunks(needPng, 10)) {
      const j = await this.need().get(`images/${this.meta.fileKey}?ids=${chunk.join(',')}&format=png&scale=${scale}`);
      await Promise.all(chunk.map(async id => {
        if (!j.images[id]) throw Error(`Figma could not render outlined text fallback ${id}`);
        writeFileSync(this.outlinedTextFallbackFile(id, scale), await this.need().download(j.images[id]));
      }));
    }
  }

  // Always PNG: an image fill can carry alpha; opaque renders become JPEG when placed.
  async ensureRasters(ids: string[], scale: number) {
    mkdirSync(join(this.dir, 'raster'), {recursive: true});
    const need = ids.filter(id => !existsSync(join(this.dir, 'raster', `${fileName(id)}@${scale}.png`)) && !existsSync(join(this.dir, 'raster', `${fileName(id)}@${scale}.jpg`)));
    for (const chunk of chunks(need, 10)) {
      const j = await this.need().get(`images/${this.meta.fileKey}?ids=${chunk.join(',')}&format=png&scale=${scale}&use_absolute_bounds=true`);
      await Promise.all(chunk.map(async id => {
        if (!j.images[id]) throw Error(`Figma could not render ${id}`);
        writeFileSync(join(this.dir, 'raster', `${fileName(id)}@${scale}.png`), await this.need().download(j.images[id]));
      }));
    }
  }

  // Figma's own PDF of each frame: the benchmark's reference.
  async framePdfs(ids: string[]): Promise<string[]> {
    const dir = join(this.dir, 'figpdf'); mkdirSync(dir, {recursive: true});
    const path = (id: string) => join(dir, fileName(id) + '.pdf');
    for (const chunk of chunks(ids.filter(id => !existsSync(path(id))), 4)) {
      const j = await this.need().get(`images/${this.meta.fileKey}?ids=${chunk.join(',')}&format=pdf`);
      await Promise.all(chunk.map(async id => { if (!j.images[id]) throw Error(`Figma could not export ${id} as PDF`); writeFileSync(path(id), await this.need().download(j.images[id])); }));
    }
    return ids.map(path);
  }
}

export function describePages(meta: FileMeta): string {
  return meta.pages.map(p => `  page "${p.name}" (${p.id}): ${p.frames.length} frames${p.frames.length ? ': ' + p.frames.slice(0, 8).map(f => `"${f.name}" ${f.id}`).join(', ') + (p.frames.length > 8 ? ', ...' : '') : ''}`).join('\n');
}

// Which frames to convert. Selectors are frame ids ("12:34" or "12-34") or exact frame names; --page picks all
// visible frames of a page (or restricts name lookups to it); with neither, a URL's node-id is used (a page
// id means all its frames).
export function resolveFrames(meta: FileMeta, selectors: string[], opts: {page?: string, nodeId?: string}): string[] {
  let pages = meta.pages;
  if (opts.page) {
    const p = meta.pages.filter(p => p.name === opts.page || p.id === normalizeId(opts.page!));
    if (!p.length) throw Error(`no page named "${opts.page}" in "${meta.name}". Pages:\n${describePages(meta)}`);
    pages = p;
  }
  const all = pages.flatMap(p => p.frames.map(f => ({...f, page: p.name})));
  if (selectors.length) return selectors.map(s => {
    if (isNodeId(s)) return normalizeId(s);
    const hits = all.filter(f => f.name === s);
    if (hits.length === 1) return hits[0].id;
    if (!hits.length) throw Error(`no frame named "${s}"${opts.page ? ` on page "${opts.page}"` : ''}. Pages:\n${describePages(meta)}`);
    throw Error(`frame name "${s}" is ambiguous: ${hits.map(h => `${h.id} on "${h.page}"`).join(', ')}; use the id or --page`);
  });
  if (opts.page) return all.filter(f => f.visible).map(f => f.id);
  if (opts.nodeId) {
    const page = meta.pages.find(p => p.id === opts.nodeId);
    return page ? page.frames.filter(f => f.visible).map(f => f.id) : [opts.nodeId];
  }
  throw Error(`name the frames (ids or names), use --page, or pass a URL with a node-id. Pages in "${meta.name}":\n${describePages(meta)}`);
}
