import {closeSync, existsSync, mkdirSync, openSync, readSync, writeFileSync} from 'node:fs';
import {homedir} from 'node:os';
import {join, resolve} from 'node:path';
import {buildPptx, slideXml} from '../../src/convert/pptx.ts';
import {exportPdf, powerPointBusyReason, probePowerPoint} from '../../src/powerpoint/export.ts';
import type {DoctorResult} from './protocol.ts';

function assertFullDiskAccess(): void {
  const tcc = join(homedir(), 'Library', 'Application Support', 'com.apple.TCC', 'TCC.db');
  let fd: number | undefined;
  try {
    fd = openSync(tcc, 'r');
    readSync(fd, Buffer.alloc(1), 0, 1, 0);
  } catch {
    throw Error('Full Disk Access is missing for Figma2Pptx Bridge.app. Enable it in System Settings > Privacy & Security > Full Disk Access.');
  } finally {
    if (fd !== undefined) closeSync(fd);
  }
}

export function runDoctorWorker(outputDir: string): DoctorResult {
  assertFullDiskAccess();
  const before = probePowerPoint();
  const busy = powerPointBusyReason(before);
  if (busy) throw Error(`${busy}. Doctor cannot safely continue.`);

  const stamp = new Date().toISOString().replace(/\D/g, '').slice(0, 17);
  const jobDir = resolve(outputDir, 'doctor', stamp);
  mkdirSync(jobDir, {recursive: true, mode: 0o700});
  const pptx = join(jobDir, 'permission-check.pptx');
  const pdf = join(jobDir, 'permission-check.pdf');
  writeFileSync(pptx, buildPptx([{xml: slideXml('', 'Permission check'), rels: []}], [], 1600, 900, 'figma2pptx permission check'));
  try {
    const slides = exportPdf(pptx, pdf, {timeoutMs: 20000});
    if (slides !== 1 || !existsSync(pdf)) throw Error('PowerPoint did not produce the one-slide permission-check PDF');
  } catch (error) {
    let after;
    try { after = probePowerPoint(); } catch {}
    if (after && (after.modalWindows || after.sheetWindows)) {
      throw Error('PowerPoint opened a file-access or other modal dialog. Grant the requested access once, dismiss the dialog, and rerun doctor.');
    }
    throw Error(`Headless PowerPoint export failed: ${error instanceof Error ? error.message : String(error)}`);
  }
  const after = probePowerPoint();
  if (after.presentations || after.modalWindows || after.sheetWindows) {
    throw Error('PowerPoint did not return to an idle, dialog-free state after the permission-check export.');
  }
  return {
    ok: true,
    checks: [
      'Figma2Pptx Bridge.app has Full Disk Access',
      'Automation to Microsoft PowerPoint works',
      'PowerPoint has no open presentation, modal dialog, or file-access sheet',
      'One-slide container-staged headless export completed without a prompt',
    ],
    jobDir,
  };
}

if (import.meta.main) {
  try {
    const outputDir = process.argv[2];
    if (!outputDir) throw Error('doctor output directory is missing');
    process.stdout.write(JSON.stringify(runDoctorWorker(outputDir)) + '\n');
  } catch (error) {
    process.stderr.write((error instanceof Error ? error.message : String(error)) + '\n');
    process.exit(1);
  }
}
