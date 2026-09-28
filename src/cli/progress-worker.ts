// Progress renderer, in its own thread: the conversion blocks the main thread for seconds at a time (PowerPoint
// export, PDF render gate), and the spinner must keep moving meanwhile.
//
// Messages: {t: 'start', tty, heartbeatMs, at} | {t: 'stage' | 'update' | 'note', text, at} | {t: 'end', ok, text, at}
// A stage is a step with its own duration; an update changes the current stage's text (slide 3/14).
// TTY: completed stages stay as "✓ text  1.2s" lines, the current one spins with the total elapsed time.
// Not a TTY: one "[  12.3s] text" line per stage, plus "still ..." lines during long silent stages.
declare const self: Worker;

const FRAMES = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏'];
let tty = false, heartbeatMs = 2000, t0 = 0;
let stage: {text: string, at: number} | null = null;
let lastLine = 0, frame = 0, timer: ReturnType<typeof setInterval> | null = null;

const secs = (ms: number) => (ms / 1000).toFixed(1) + 's';
const stamp = (at: number) => `[${secs(at - t0).padStart(7)}]`;
const out = (s: string) => process.stderr.write(s);

function render() {
  if (!stage) return;
  if (tty) {
    out(`\r\x1b[2K${FRAMES[frame = (frame + 1) % FRAMES.length]} ${stage.text}  \x1b[2m${secs(Date.now() - t0)}\x1b[0m`);
  } else if (Date.now() - lastLine >= heartbeatMs) {
    out(`${stamp(Date.now())} still ${stage.text} (${secs(Date.now() - stage.at)})\n`);
    lastLine = Date.now();
  }
}

function closeStage(at: number) {
  if (stage && tty) out(`\r\x1b[2K\x1b[32m✓\x1b[0m ${stage.text}  \x1b[2m${secs(at - stage.at)}\x1b[0m\n`);
  stage = null;
}

self.onmessage = (e: MessageEvent) => {
  const m = e.data;
  if (m.t === 'start') {
    tty = m.tty; heartbeatMs = m.heartbeatMs; t0 = m.at; lastLine = m.at;
    timer = setInterval(render, tty ? 80 : 250);
  } else if (m.t === 'stage') {
    closeStage(m.at);
    stage = {text: m.text, at: m.at};
    if (!tty) { out(`${stamp(m.at)} ${m.text}\n`); lastLine = Date.now(); }
    render();
  } else if (m.t === 'update') {
    if (!stage) stage = {text: m.text, at: m.at}; else stage.text = m.text;
    if (!tty) { out(`${stamp(m.at)} ${m.text}\n`); lastLine = Date.now(); }
    render();
  } else if (m.t === 'note') {
    if (tty) out(`\r\x1b[2K${m.text}\n`); else { out(`${m.text}\n`); lastLine = Date.now(); }
    render();
  } else if (m.t === 'end') {
    if (m.ok) closeStage(m.at);
    else { if (tty) out('\r\x1b[2K'); stage = null; }
    if (timer) clearInterval(timer);
    out(m.text + '\n');
    postMessage('ended');
  }
};
