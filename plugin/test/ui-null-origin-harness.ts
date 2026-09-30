import {readFileSync} from 'node:fs';
import {join} from 'node:path';

const root = join(import.meta.dir, '..');
const ui = readFileSync(join(root, 'local', 'dist', 'ui.html'), 'utf8');
const mock = `<script>
Object.defineProperty(globalThis, 'crypto', {value: {}, configurable: true});
const mockBytes = new TextEncoder().encode('pptx bytes');
let healthAttempts = 0;
globalThis.fetch = async url => {
  const path = String(url);
  if (path.endsWith('/v1/health')) {
    healthAttempts += 1;
    if (healthAttempts < 3) throw new TypeError('Failed to fetch');
    return Response.json({ok: true, busy: false});
  }
  if (path.endsWith('/v1/exports')) {
    const events = [
      {type: 'accepted', jobId: '11111111-1111-4111-8111-111111111111', stage: 'Starting export', step: 'Preparing', slide: 0, totalSlides: 1, elapsed: 0},
      {type: 'done', jobId: '11111111-1111-4111-8111-111111111111', outputs: {pptx: {name: 'harness.pptx', downloadPath: '/v1/exports/11111111-1111-4111-8111-111111111111/files/pptx', remotePath: '/runtime/harness.pptx', size: 10, sha256: '77db3ca8ad18a50fbc2168e6d9413fc1662eb4255ba840ae19c53c4c29b81bfe'}, report: {name: 'harness.report.json', downloadPath: '/report', remotePath: '/runtime/report', size: 2, sha256: ''}}, slides: 1, elapsed: 0.1},
    ];
    return new Response(events.map(value => JSON.stringify(value)).join('\\n') + '\\n', {headers: {'Content-Type': 'application/x-ndjson'}});
  }
  if (path.includes('/files/pptx')) return new Response(mockBytes);
  return Response.json({error: 'not found'}, {status: 404});
};
</script>`;
const driver = `<script>
const waitFor = async (predicate, label) => {
  for (let i = 0; i < 100; i++) { if (predicate()) return; await new Promise(resolve => setTimeout(resolve, 25)); }
  throw Error('timed out waiting for ' + label);
};
(async () => {
  try {
    window.dispatchEvent(new MessageEvent('message', {data: {pluginMessage: {type: 'state', request: {fileKey: 'AbCdEfGhIjKl', frameIds: ['1:2'], frameNames: ['Harness frame']}, bridgeSecret: 'a'.repeat(64), activeJobId: ''}}}));
    await waitFor(() => document.getElementById('bridgeLabel').textContent === 'Reconnecting…', 'automatic reconnect state');
    await waitFor(() => !document.getElementById('export').disabled, 'configured export button');
    document.getElementById('export').click();
    await waitFor(() => document.getElementById('stage').textContent === 'Export complete', 'export completion');
    document.getElementById('downloadPptx').click();
    await waitFor(() => document.getElementById('downloadPptx').textContent === 'Downloaded', 'verified download');
    parent.postMessage({type: 'harness-result', ok: true, subtle: Boolean(globalThis.crypto && globalThis.crypto.subtle), stage: document.getElementById('stage').textContent, healthAttempts, pairingHidden: document.getElementById('pairing').classList.contains('hidden')}, '*');
  } catch (error) {
    parent.postMessage({type: 'harness-result', ok: false, error: error instanceof Error ? error.message : String(error)}, '*');
  }
})();
</script>`;

const srcdoc = ui.replace('<script>', mock + '<script>').replace('</body>', driver + '</body>');
const serializedSrcdoc = JSON.stringify(srcdoc).replaceAll('</script>', '<\\/script>');
const html = `<!doctype html><meta charset="utf-8"><title>figma2pptx null-origin harness</title>
<h1 id="result">RUNNING</h1><pre id="details"></pre><iframe id="plugin" sandbox="allow-scripts allow-downloads" style="width:440px;height:620px"></iframe>
<script>
const frame = document.getElementById('plugin');
frame.srcdoc = ${serializedSrcdoc};
window.addEventListener('message', event => {
  if (event.data && event.data.type === 'harness-result') {
    const pass = event.data.ok && event.data.subtle === false && event.data.stage === 'Export complete' && event.data.healthAttempts === 3 && event.data.pairingHidden === true;
    document.getElementById('result').textContent = pass ? 'PASS' : 'FAIL';
    document.getElementById('details').textContent = JSON.stringify(event.data, null, 2);
    document.title = pass ? 'PASS - figma2pptx null-origin harness' : 'FAIL - figma2pptx null-origin harness';
  }
});
</script>`;

const server = Bun.serve({hostname: '127.0.0.1', port: 0, fetch: () => new Response(html, {headers: {'Content-Type': 'text/html; charset=utf-8'}})});
console.log(`http://127.0.0.1:${server.port}/`);
