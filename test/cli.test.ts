// The command line: plain line-per-stage progress when stderr is not a terminal, a spinner when it is, and a
// one-line summary at the end.
import {test, expect} from 'bun:test';
import {rmSync, existsSync} from 'node:fs';
import {join} from 'node:path';
import {convertFigma, FontPreflightError} from '../src/pipeline';
import {FIXTURE_KEY, fixtureCacheCopy, fakeImages, tableFonts} from './helpers';

const CLI = join(import.meta.dir, '../src/cli.ts');
const ROOT = join(import.meta.dir, '..');
const run = async (args: string[], env: Record<string, string> = {}) => {
  const p = Bun.spawn(['bun', CLI, ...args], {cwd: ROOT, stdout: 'pipe', stderr: 'pipe', env: {PATH: process.env.PATH!, HOME: process.env.HOME!, ...env}});
  const [out, err, code] = [await new Response(p.stdout).text(), await new Response(p.stderr).text(), await p.exited];
  return {out, err, code};
};

test('non-TTY: a timestamped line per stage and step, the font table, timings, a final summary; no escape codes', async () => {
  const cache = fixtureCacheCopy();
  try {
    const out = join(cache, 'out', 'deck.pptx');
    const r = await run([FIXTURE_KEY, '--page', 'Slides', '--offline', '--cache', cache, '--single-pass', '--no-embed-fonts', '--allow-font-fallback', '--timings', '-o', out]);
    expect(r.code).toBe(0);
    expect(r.err).not.toContain('\x1b');
    const lines = r.err.trimEnd().split('\n');
    const stamped = lines.filter(l => /^\[\s*\d+\.\ds\] /.test(l)).map(l => l.replace(/^\[\s*\d+\.\ds\] /, ''));
    expect(stamped).toEqual([
      `reading cached Figma file ${FIXTURE_KEY}`,
      'fetching 2 frames from Figma ("Synthetic deck")',
      'building',
      'building: renders from Figma (1 nodes)',
      'building: slide 1/2',
      'building: slide 2/2',
      'checking fonts', // no 'embedding fonts' step: --no-embed-fonts
    ]);
    expect(lines.some(l => /^\s+(ok|SUBST|MISS)\s+ArialMT -> Arial/.test(l))).toBe(true);
    expect(lines.some(l => /^\s+timings:$/.test(l))).toBe(true);
    expect(lines.some(l => /^\s+total\s+\d+\.\d\ds$/.test(l))).toBe(true);
    expect(lines.at(-1)).toMatch(new RegExp(`^done: 2 slides in \\d+\\.\\ds -> ${out.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')} \\(\\d+ KB\\); report .*deck\\.report\\.json$`));
    expect(existsSync(out)).toBe(true);
  } finally { rmSync(cache, {recursive: true, force: true}); }
});

test('errors end with one clear line and exit code 1', async () => {
  const cache = fixtureCacheCopy();
  try {
    const r = await run([FIXTURE_KEY, 'Details', '--offline', '--cache', cache, '--single-pass']);
    expect(r.code).toBe(1);
    expect(r.err.trimEnd().split('\n').at(-1)).toBe('figma2pptx: frame name "Details" is ambiguous: 1:40 on "Slides", 2:1 on "Scratch"; use the id or --page');
  } finally { rmSync(cache, {recursive: true, force: true}); }
});

test('usage errors exit 2; --help and --version print to stdout', async () => {
  expect((await run(['--bogus'])).code).toBe(2);
  const h = await run(['--help']);
  expect(h.code).toBe(0);
  expect(h.out).toContain('figma2pptx <figma-url | file-key> [frame ...]');
  expect((await run(['--version'])).out.trim()).toMatch(/^\d+\.\d+\.\d+$/);
});

const progressScript = (tty: boolean) => `
import {createProgress} from ${JSON.stringify(join(ROOT, 'src/cli/progress.ts'))};
const p = createProgress({tty: ${tty}, heartbeatMs: 150});
p.stage('exporting with PowerPoint');
Bun.sleepSync(700); // the main thread is blocked, as during an export
p.update('exporting: slide 2/2');
p.note('  a note');
await p.end(true, 'done: summary');
`;

test('non-TTY progress keeps talking while the main thread is blocked', async () => {
  const p = Bun.spawn(['bun', '-e', progressScript(false)], {stderr: 'pipe'});
  const err = await new Response(p.stderr).text();
  expect(await p.exited).toBe(0);
  const lines = err.trimEnd().split('\n');
  expect(lines[0]).toMatch(/^\[\s*0\.\ds\] exporting with PowerPoint$/);
  expect(lines.filter(l => /\] still exporting with PowerPoint \(\d+\.\ds\)$/.test(l)).length).toBeGreaterThanOrEqual(2);
  expect(lines.slice(-3)).toEqual([expect.stringMatching(/\] exporting: slide 2\/2$/), '  a note', 'done: summary']);
});

test('TTY progress: a spinner that redraws in place, then a checkmark line with the stage duration', async () => {
  const p = Bun.spawn(['bun', '-e', progressScript(true)], {stderr: 'pipe'});
  const err = await new Response(p.stderr).text();
  expect(await p.exited).toBe(0);
  const frames = err.split('\r\x1b[2K').filter(s => /^[⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏] exporting with PowerPoint/.test(s));
  expect(frames.length).toBeGreaterThan(4); // it kept spinning through the blocked 0.7 s
  expect(err).toMatch(/✓\x1b\[0m exporting: slide 2\/2  \x1b\[2m0\.\ds\x1b\[0m\ndone: summary\n/);
  expect(err.endsWith('done: summary\n\x1b[?25h') || err.includes('\x1b[?25h')).toBe(true);
});

test('font preflight: a face PowerPoint cannot render stops the conversion and names it', async () => {
  const cache = fixtureCacheCopy();
  try {
    const err = await convertFigma({target: FIXTURE_KEY, page: 'Slides', out: join(cache, 'x.pptx'), passes: 1, offline: true, cacheDir: cache, fonts: tableFonts([]), images: fakeImages}).catch(e => e);
    expect(err).toBeInstanceOf(FontPreflightError);
    expect(err.faces.map((f: any) => [f.figma, f.status])).toEqual([['Arial-BoldMT', 'missing'], ['ArialMT', 'missing']]);
    expect(existsSync(join(cache, 'x.pptx'))).toBe(false);
  } finally { rmSync(cache, {recursive: true, force: true}); }
});
