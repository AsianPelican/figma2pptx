// Live progress on stderr: a spinner with the current stage and elapsed time on a terminal, plain
// line-per-stage output otherwise. Rendering runs in a worker (see progress-worker.ts).
export interface Progress {
  stage(text: string): void; // start a stage (closes the previous one)
  update(text: string): void; // change the current stage's text (slide 3/14)
  note(text: string): void; // a permanent line: warnings, the font table
  end(ok: boolean, text: string): Promise<void>; // the final summary or error line; stops the renderer
}

export function createProgress(opts: {tty?: boolean, heartbeatMs?: number} = {}): Progress {
  const tty = opts.tty ?? !!process.stderr.isTTY;
  const worker = new Worker(new URL('./progress-worker.ts', import.meta.url));
  worker.postMessage({t: 'start', tty, heartbeatMs: opts.heartbeatMs ?? 2000, at: Date.now()});
  let ended = false;
  const restoreCursor = () => { if (tty) process.stderr.write('\x1b[?25h'); };
  if (tty) process.stderr.write('\x1b[?25l');
  process.once('exit', restoreCursor);
  return {
    stage: text => worker.postMessage({t: 'stage', text, at: Date.now()}),
    update: text => worker.postMessage({t: 'update', text, at: Date.now()}),
    note: text => worker.postMessage({t: 'note', text, at: Date.now()}),
    end(ok, text) {
      if (ended) return Promise.resolve();
      ended = true;
      return new Promise(resolve => {
        worker.onmessage = () => { worker.terminate(); restoreCursor(); resolve(); };
        worker.postMessage({t: 'end', ok, text, at: Date.now()});
      });
    },
  };
}
