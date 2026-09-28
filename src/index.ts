// figma2pptx as a library. The CLI (src/cli.ts) is one front end over convertFigma; a local server behind a
// Figma plugin can be another, consuming the same progress events.
export {convertFigma, pdfFromDeck, fontTable, FontPreflightError, type ConvertOptions, type ConvertResult, type ProgressEvent, type PdfPreset, type StageTiming} from './pipeline';
export {buildDeck, rasterReason, type FrameSource, type BuildOptions, type BuildEnv, type BuildReport, type Corrections} from './convert/build';
export {fontconfigResolver, mapFaceIn, type FontResolver, type Face, type FontRow} from './convert/fonts';
export {magickOps, type ImageOps} from './convert/images';
export {FigmaClient, readToken, parseTarget, normalizeId} from './figma/api';
export {FigmaFile, resolveFrames, type FileMeta} from './figma/source';
export {measure, type Measurement} from './measure/corrections';
export {exportPdf} from './powerpoint/export';
export {optimizePdf, PRESETS, type OptimizeResult} from './pdf/optimize';
