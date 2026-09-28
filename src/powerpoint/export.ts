// Export a PPTX to PDF with Microsoft PowerPoint for Mac, through AppleScript.
//
// PowerPoint only ever sees a copy, so the source deck is never opened or locked. The copy and the PDF are
// staged inside PowerPoint's own sandbox container: saving anywhere else raises a "Grant File Access" sheet for
// folders PowerPoint has not been granted, which blocks AppleScript.
import {execFileSync} from 'node:child_process';
import {copyFileSync, mkdirSync, writeFileSync, rmSync, existsSync, readFileSync} from 'node:fs';
import {join, resolve, basename} from 'node:path';
import {homedir, tmpdir} from 'node:os';

const SCRIPT = `on run argv
 set s to POSIX file (item 1 of argv)
 set o to POSIX file (item 2 of argv)
 set nm to name of (info for s)
 tell application "Microsoft PowerPoint"
  open s
  repeat 300 times
   if exists presentation nm then exit repeat
   delay 0.2
  end repeat
  set d to presentation nm
  set c to count of slides of d
  with timeout of 180 seconds
   save d in o as save as PDF
  end timeout
  close d saving no
  return c
 end tell
end run`;

const PROBE_SCRIPT = `tell application "System Events"
 set isRunning to exists process "Microsoft PowerPoint"
end tell
if not isRunning then return "0|0|0|0"
with timeout of 5 seconds
 tell application "Microsoft PowerPoint"
  set n to count of presentations
  set m to count of (windows whose modal is true)
  set s to count of (windows whose sheet is true)
 end tell
end timeout
return "1|" & n & "|" & m & "|" & s`;

export type PowerPointState = {running: boolean, presentations: number, modalWindows: number, sheetWindows: number};
export type ExportOptions = {timeoutMs?: number, probe?: () => PowerPointState};

export function parsePowerPointState(value: string): PowerPointState {
  const [running, presentations, modalWindows, sheetWindows] = value.trim().split('|').map(Number);
  if (![running, presentations, modalWindows, sheetWindows].every(Number.isFinite)) throw Error(`unexpected PowerPoint readiness response: ${value.trim()}`);
  return {running: running === 1, presentations, modalWindows, sheetWindows};
}

export function powerPointBusyReason(s: PowerPointState): string | undefined {
  if (s.presentations) return `PowerPoint already has ${s.presentations} presentation${s.presentations === 1 ? '' : 's'} open; close them and retry (figma2pptx will not touch unrelated decks)`;
  if (s.modalWindows || s.sheetWindows) return 'PowerPoint has a modal dialog or sheet open; dismiss it and retry';
}

export function probePowerPoint(): PowerPointState {
  try {
    return parsePowerPointState(execFileSync('osascript', ['-e', PROBE_SCRIPT], {encoding: 'utf8', timeout: 7000, stdio: ['ignore', 'pipe', 'pipe']}));
  } catch (e: any) {
    if (e.signal === 'SIGTERM' || e.code === 'ETIMEDOUT') throw Error('PowerPoint is busy or has a modal dialog open; dismiss it and retry');
    throw Error('PowerPoint readiness check failed: ' + String(e.stderr || e.message).trim().split('\n').pop());
  }
}

export function powerPointExportFailure(e: {signal?: string, code?: string, stderr?: unknown, message?: string}, timeoutMs: number): Error {
  const msg = String(e.stderr || e.message || 'unknown error');
  if (e.signal === 'SIGTERM' || e.code === 'ETIMEDOUT' || /timed out|-1712/i.test(msg)) return Error(`PowerPoint did not finish the export within ${Math.round(timeoutMs / 1000)} seconds (it may be busy or a repair/permission dialog may be open)`);
  return Error('PowerPoint export failed: ' + msg.trim().split('\n').pop());
}

function stagingDir(): string {
  const container = join(homedir(), 'Library/Containers/com.microsoft.Powerpoint/Data/tmp');
  const dir = join(existsSync(container) ? container : tmpdir(), `figma2pptx-${process.pid}-${Date.now()}`);
  mkdirSync(dir, {recursive: true});
  return dir;
}

// A PowerPoint that is still launching answers AppleScript with "Connection is invalid" (-609) or "isn't
// running" (-600); wait for it and try again.
function runExport(script: string, deck: string, out: string, timeoutMs: number): number {
  for (let attempt = 1; ; attempt++) {
    try {
      return +execFileSync('osascript', [script, deck, out], {encoding: 'utf8', timeout: timeoutMs, stdio: ['ignore', 'pipe', 'pipe']}).trim();
    } catch (e: any) {
      const msg = String(e.stderr || e.message);
      if (attempt < 4 && /-609|-600|Connection is invalid|isn.t running/.test(msg)) { Bun.sleepSync(3000); continue; }
      throw powerPointExportFailure(e, timeoutMs);
    }
  }
}

// On an export error, close only the uniquely named staging copy that figma2pptx opened. A modal PowerPoint
// may refuse the command; the cleanup itself is bounded and never targets another presentation.
function closeStagingCopy(name: string): void {
  const script = `on run argv
 set nm to item 1 of argv
 tell application "Microsoft PowerPoint"
  if exists presentation nm then close presentation nm saving no
 end tell
end run`;
  try { execFileSync('osascript', ['-e', script, name], {timeout: 5000, stdio: 'ignore'}); } catch {}
}

// Returns PowerPoint's slide count; the PDF lands at `pdf`.
export function exportPdf(pptx: string, pdf: string, options: ExportOptions = {}): number {
  const busy = powerPointBusyReason((options.probe ?? probePowerPoint)());
  if (busy) throw Error(busy);
  const dir = stagingDir();
  try {
    const deck = join(dir, `F2P_${Date.now()}_${basename(pptx)}`), staged = join(dir, 'export.pdf');
    copyFileSync(resolve(pptx), deck);
    writeFileSync(join(dir, 'export.applescript'), SCRIPT);
    let n: number;
    try { n = runExport(join(dir, 'export.applescript'), deck, staged, options.timeoutMs ?? 180000); }
    catch (e) { closeStagingCopy(basename(deck)); throw e; }
    if (!existsSync(staged)) throw Error('PowerPoint did not write the PDF');
    writeFileSync(resolve(pdf), readFileSync(staged));
    return n;
  } finally {
    rmSync(dir, {recursive: true, force: true});
  }
}
