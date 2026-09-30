import {expect, test} from 'bun:test';

type ProbeResult = {ok: boolean; unauthorized?: boolean; busy?: boolean};
type Scheduled = {callback: () => void; delay: number};

declare global {
  var figma2pptxCreateAutoConnector: (options: {
    probe: (secret: string) => Promise<ProbeResult>;
    onConnecting: () => void;
    onReconnecting: () => void;
    onConnected: (result: ProbeResult) => void;
    onUnauthorized: (result: ProbeResult) => void;
    schedule: (callback: () => void, delay: number) => number;
    cancel: (id: number) => void;
  }) => {start(secret: string): Promise<void> | undefined; retry(): void; stop(): void};
}

// Browser-targeted plain JS is injected verbatim into the plugin UI at build time.
// @ts-expect-error no TypeScript declaration is shipped for this injected script
await import('../src/connection.js');

const flush = () => new Promise(resolve => setTimeout(resolve, 0));

test('stored pairing auto-connects on open and retries network failures with backoff', async () => {
  const scheduled: Scheduled[] = [];
  const states: string[] = [];
  let probes = 0;
  const connector = globalThis.figma2pptxCreateAutoConnector({
    probe: async secret => {
      expect(secret).toBe('stored-secret');
      probes += 1;
      if (probes < 3) throw Error('network unavailable');
      return {ok: true, busy: false};
    },
    onConnecting: () => states.push('connecting'),
    onReconnecting: () => states.push('reconnecting'),
    onConnected: () => states.push('connected'),
    onUnauthorized: () => states.push('unauthorized'),
    schedule: (callback, delay) => (scheduled.push({callback, delay}), scheduled.length),
    cancel: () => {},
  });

  await connector.start('stored-secret');
  expect(states).toEqual(['connecting', 'reconnecting']);
  expect(scheduled[0]?.delay).toBe(250);

  scheduled.shift()?.callback();
  await flush();
  expect(states).toEqual(['connecting', 'reconnecting', 'reconnecting']);
  expect(scheduled[0]?.delay).toBe(500);

  scheduled.shift()?.callback();
  await flush();
  expect(states.at(-1)).toBe('connected');
  expect(probes).toBe(3);
});

test('401 stops retries and returns the UI to pairing', async () => {
  const scheduled: Scheduled[] = [];
  const states: string[] = [];
  const connector = globalThis.figma2pptxCreateAutoConnector({
    probe: async () => ({ok: false, unauthorized: true}),
    onConnecting: () => states.push('connecting'),
    onReconnecting: () => states.push('reconnecting'),
    onConnected: () => states.push('connected'),
    onUnauthorized: () => states.push('unauthorized'),
    schedule: (callback, delay) => (scheduled.push({callback, delay}), scheduled.length),
    cancel: () => {},
  });

  await connector.start('rejected-secret');
  expect(states).toEqual(['connecting', 'unauthorized']);
  expect(scheduled).toHaveLength(0);
});
