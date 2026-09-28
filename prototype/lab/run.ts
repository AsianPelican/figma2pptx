// One command, Figma frames -> measured 1:1 PPTX (+ PowerPoint PDF).
//   bun run.ts <fileKey> <frameIds> <outStem>
// build -> PowerPoint export -> measure every text line against Figma's baselines -> rebuild with per-box offsets
// -> export again.  Writes <outStem>.pptx, <outStem>.pdf, <outStem>.report.json, <outStem>.corr.json.
import {execFileSync} from 'node:child_process';
const [key, ids, stem] = process.argv.slice(2);
const t0 = Date.now(), lap = (s: string) => console.log(`[${((Date.now() - t0) / 1000).toFixed(1)}s] ${s}`);
const run = (...a: string[]) => execFileSync('bun', a, {stdio: 'inherit'});
run('figma2pptx.ts', key, ids, `${stem}.pptx`); lap('pass 1 built');
run('ppexport.ts', `${stem}.pptx`, `${stem}.pdf`); lap('pass 1 exported');
run('corr.ts', `${stem}.pdf`, `${stem}.report.json`, '-', `${stem}.corr.json`); lap('measured');
run('figma2pptx.ts', key, ids, `${stem}.pptx`, '--corr', `${stem}.corr.json`); lap('pass 2 built');
run('ppexport.ts', `${stem}.pptx`, `${stem}.pdf`); lap('pass 2 exported: done');
