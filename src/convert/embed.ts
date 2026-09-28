// Embed the faces a deck uses, so it renders the same on a machine without them: each face's font file,
// wrapped as EOT, in the slot PowerPoint selects it by (typeface + bold/italic flags).
import {readFileSync} from 'node:fs';
import {readSfnt, toEot, embeddingAllowed} from './sfnt';
import type {Face} from './fonts';
import type {EmbeddedFont} from './pptx';

export type EmbedResult = {typeface: string, slot: string, file?: string, embedded: boolean, reason?: string};

const slotOf = (f: Face) => f.b ? (f.i ? 'boldItalic' : 'bold') : (f.i ? 'italic' : 'regular');

export function embedFaces(faces: Face[]): {fonts: EmbeddedFont[], results: EmbedResult[]} {
  const byTypeface = new Map<string, EmbeddedFont>(), results: EmbedResult[] = [];
  const seen = new Set<string>();
  for (const f of faces) {
    const slot = slotOf(f), key = `${f.typeface}|${slot}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const res: EmbedResult = {typeface: f.typeface, slot, file: f.file, embedded: false};
    results.push(res);
    if (f.status === 'variable-only' || f.status === 'missing') { res.reason = 'not installed as a static face'; continue; }
    if (!f.file) { res.reason = 'font file unknown'; continue; }
    try {
      const info = readSfnt(new Uint8Array(readFileSync(f.file)), f.index ?? 0);
      if (!embeddingAllowed(info.fsType)) { res.reason = `the font's licence flags forbid embedding (fsType 0x${info.fsType.toString(16)})`; continue; }
      if (info.variable) { res.reason = 'variable font'; continue; }
      if (!byTypeface.has(f.typeface)) byTypeface.set(f.typeface, {typeface: f.typeface});
      byTypeface.get(f.typeface)![slot] = toEot(info);
      res.embedded = true;
    } catch (e: any) { res.reason = `unreadable font file (${e.message})`; }
  }
  return {fonts: [...byTypeface.values()], results};
}
