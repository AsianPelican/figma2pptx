// Text lines of one PDF page, with every character's origin, through mupdf's structured text.
import * as m from 'mupdf';

export type PdfChar = {c: string, x: number, y: number, x1: number};
export type PdfLine = {t: string, x0: number, x1: number, base: number, size: number, font: string, chars: PdfChar[]};

// `scale` converts PDF points to the caller's unit (1 / 0.75 for Figma px).
export function pdfLines(pdf: Uint8Array, page: number, scale: number): PdfLine[] {
  const d = m.Document.openDocument(Buffer.from(pdf), 'application/pdf');
  const p = d.loadPage(page);
  const out: PdfLine[] = [];
  let cur: any;
  p.toStructuredText('preserve-whitespace').walk({
    beginLine() { cur = {t: '', chars: [], size: 0, font: ''}; },
    onChar(c: string, o: number[], f: any, s: number, q: number[]) {
      cur.t += c; cur.chars.push({c, x: o[0] * scale, y: o[1] * scale, x1: Math.max(q[2], q[6]) * scale});
      if (c.trim()) { cur.size = Math.max(cur.size, s * scale); cur.font = f.getName(); }
    },
    endLine() {
      const cs = cur.chars.filter((c: PdfChar) => c.c.trim());
      if (!cs.length) return;
      out.push({t: cur.t, x0: cs[0].x, x1: Math.max(...cs.map((c: PdfChar) => c.x1)), base: cs[0].y, size: cur.size, font: cur.font, chars: cur.chars});
    },
  } as any);
  return out;
}

export const pageCount = (pdf: Uint8Array) => m.Document.openDocument(Buffer.from(pdf), 'application/pdf').countPages();
