# Prototype provenance

Copies of the figma2pptx prototype, taken 2026-09-28 before any refactor.

- `lab/` is the scout prototype of 2026-09-25: the converter (`run.ts`, `figma2pptx.ts` and its modules) and the benchmark (`bench.ts`).
- `tooling/` holds the two helpers it called from a sibling tooling folder: `export-pdf.ts` (size-optimized, render-gated PDF export) and `lines1to1.ts` (the line comparator `bench.ts` imports).

The files are byte-identical to the originals except for three scrubbed literals, so that no client material or local path is published:

- `lab/fetchsvg.ts` read the token file from a client-named folder in the home directory; it now reads the path from `FIGMA_ENV_FILE`.
- `lab/bench.ts` imported `lines1to1.ts` by an absolute home path; it now imports `../tooling/lines1to1.ts`.
- `lab/bench.ts` defaulted `--frames` to a client deck's node ids; the default is now empty.

Left out entirely, because they carry client deck content: the calibration decks and their measurements, the benchmark result files, and every cached Figma response.
