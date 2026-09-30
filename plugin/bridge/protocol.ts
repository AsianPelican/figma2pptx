export const BRIDGE_PORT = 38458;
export const BRIDGE_ORIGIN = `https://bridge-host.invalid:${BRIDGE_PORT}`;
export const BRIDGE_BIND_HOST = '127.0.0.1';
export const BRIDGE_BIND_PORT = 38456;

export function parseBridgeOrigin(value = BRIDGE_ORIGIN): string {
  const origin = value.trim().replace(/\/$/, '');
  const url = new URL(origin);
  if (url.protocol !== 'https:' || url.port !== String(BRIDGE_PORT) || url.pathname !== '/' || url.username || url.password || url.search || url.hash) {
    throw Error(`bridge origin must be an HTTPS origin on port ${BRIDGE_PORT}`);
  }
  if (!url.hostname.includes('.')) throw Error('bridge origin hostname must be fully qualified and contain a dot for Figma');
  return origin;
}

export type PdfPreset = 'screen' | 'standard' | 'print' | 'raw';

export type ExportRequest = {
  fileKey: string;
  frames: string[];
  frameNames: string[];
  page?: string;
  outputName?: string;
  pdf: PdfPreset | false;
  passes: 1 | 2;
  embedFonts: boolean;
  allowFontFallback: boolean;
  offline: boolean;
  scale: number;
  kern: string;
  corr?: Record<string, {dx: number; dy: number}>;
};

export type ExportResult = {
  pptx: string;
  pdf?: string;
  report: string;
  slides: number;
  seconds: number;
};

export type OutputKind = 'pptx' | 'pdf' | 'report';

export type DownloadOutput = {
  name: string;
  downloadPath: string;
  remotePath: string;
  size: number;
  sha256: string;
};

export type BridgeEvent =
  | {type: 'accepted'; jobId: string; stage: string; step: string; slide: number; totalSlides: number; elapsed: number}
  | {type: 'progress'; jobId: string; stage: string; step: string; slide: number; totalSlides: number; elapsed: number}
  | {type: 'done'; jobId: string; outputs: {pptx: DownloadOutput; pdf?: DownloadOutput; report: DownloadOutput}; slides: number; elapsed: number}
  | {type: 'error'; jobId: string; error: string; elapsed: number};

export type DoctorResult = {
  ok: true;
  checks: string[];
  jobDir: string;
};

const isStrings = (value: unknown): value is string[] => Array.isArray(value) && value.every(item => typeof item === 'string');
const isOptionalString = (value: unknown): value is string | undefined => value === undefined || typeof value === 'string';

function optionalBoolean(body: Record<string, unknown>, key: string, fallback: boolean): boolean {
  const value = body[key];
  if (value === undefined) return fallback;
  if (typeof value !== 'boolean') throw Error(`${key} must be a boolean`);
  return value;
}

export function parseExportRequest(value: unknown): ExportRequest {
  if (!value || typeof value !== 'object') throw Error('request body must be a JSON object');
  const body = value as Record<string, unknown>;
  if (typeof body.fileKey !== 'string' || !/^[A-Za-z0-9]{10,128}$/.test(body.fileKey)) throw Error('fileKey is invalid');

  const pluginIds = body.frameIds;
  const selectors = body.frames ?? pluginIds ?? [];
  if (!isStrings(selectors) || selectors.length > 500) throw Error('frames must contain at most 500 ids or exact names');
  if (pluginIds !== undefined && (!isStrings(pluginIds) || pluginIds.some(id => !/^\d+[:\-]\d+$/.test(id)))) throw Error('frameIds contains an invalid node ID');
  if (selectors.some(item => !item.trim() || item.length > 300)) throw Error('frames contains an invalid selector');

  const names = body.frameNames ?? selectors;
  if (!isStrings(names) || names.length !== selectors.length) throw Error('frameNames must match frames');
  if (names.some(name => !name.trim() || name.length > 300)) throw Error('frameNames contains an invalid name');
  if (!isOptionalString(body.page) || (body.page !== undefined && (!body.page.trim() || body.page.length > 300))) throw Error('page is invalid');
  if (!selectors.length && body.page === undefined) throw Error('name frames or provide a page');

  const outputName = body.outputName;
  if (!isOptionalString(outputName) || (outputName !== undefined && !/^[^/\\]{1,180}\.pptx$/i.test(outputName))) throw Error('outputName must be a plain .pptx file name');
  const pdf = body.pdf ?? 'screen';
  if (pdf !== false && !['screen', 'standard', 'print', 'raw'].includes(String(pdf))) throw Error('pdf must be false, screen, standard, print, or raw');
  const passes = body.passes ?? 2;
  if (passes !== 1 && passes !== 2) throw Error('passes must be 1 or 2');
  const scale = body.scale ?? 2;
  if (typeof scale !== 'number' || !Number.isFinite(scale) || scale <= 0 || scale > 8) throw Error('scale must be between 0 and 8');
  const kern = body.kern ?? '100';
  if (typeof kern !== 'string' || !/^\d{1,6}$/.test(kern)) throw Error('kern is invalid');
  const corr = body.corr;
  if (corr !== undefined && (!corr || typeof corr !== 'object' || Array.isArray(corr))) throw Error('corr must be an object');
  const embedFonts = optionalBoolean(body, 'embedFonts', false);
  if (embedFonts) throw Error('font embedding is disabled; exported PowerPoint files never embed fonts');

  return {
    fileKey: body.fileKey,
    frames: [...selectors],
    frameNames: [...names],
    ...(body.page === undefined ? {} : {page: body.page}),
    ...(outputName === undefined ? {} : {outputName}),
    pdf: pdf as PdfPreset | false,
    passes,
    embedFonts: false,
    allowFontFallback: optionalBoolean(body, 'allowFontFallback', false),
    offline: optionalBoolean(body, 'offline', false),
    scale,
    kern,
    ...(corr === undefined ? {} : {corr: corr as ExportRequest['corr']}),
  };
}

export function slideFromStep(text: string): {slide: number; totalSlides: number} | undefined {
  const match = /\bslide\s+(\d+)\/(\d+)\b/i.exec(text);
  return match ? {slide: Number(match[1]), totalSlides: Number(match[2])} : undefined;
}
