import {expect, test} from 'bun:test';

declare global {
  var figma2pptxSha256Hex: (input: ArrayBuffer | Uint8Array, subtle?: SubtleCrypto | null) => Promise<string>;
}

// Browser-targeted plain JS is injected verbatim into the plugin UI at build time.
// @ts-expect-error no TypeScript declaration is shipped for this injected script
await import('../src/sha256.js');

test('pure-JS SHA-256 verifies downloads when crypto.subtle is unavailable', async () => {
  const bytes = new TextEncoder().encode('abc');
  expect(await globalThis.figma2pptxSha256Hex(bytes, null)).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad');
});

test('Web Crypto and pure-JS paths return the same digest', async () => {
  const bytes = new TextEncoder().encode('figma2pptx');
  expect(await globalThis.figma2pptxSha256Hex(bytes, null)).toBe(await globalThis.figma2pptxSha256Hex(bytes, crypto.subtle));
});
