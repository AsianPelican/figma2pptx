# figma2pptx

Deterministic Figma-to-PowerPoint conversion with editable text and vectors.

`figma2pptx` reads Figma through the REST API, uses Figma's SVG layout for exact line breaks and geometry, and writes DrawingML directly. A measured second pass through PowerPoint corrects text placement against Figma's baselines. There is no LLM in the runtime path.

## Requirements

- macOS with Microsoft PowerPoint for the measured second pass and PDF export
- [Bun](https://bun.sh/)
- ImageMagick (`magick`), Poppler (`pdftoppm`), and fontconfig (`fc-list`)
- A Figma personal access token with read access to the source file

Install the pinned dependencies:

```sh
bun install --frozen-lockfile
```

Set the token in the process environment. It is sent only in the `X-Figma-Token` request header and is never logged, cached, or written to the conversion report.

```sh
read -s FIGMA_TOKEN
export FIGMA_TOKEN
```

## Usage

```sh
bun run figma2pptx '<figma-url-or-fileKey>' '12:34' '56:78' -o out/deck.pptx
```

Frames may be node IDs or exact names. With no frame arguments, a URL's `node-id` is used; `--page <name-or-id>` converts all visible frames on that page.

Useful options:

```text
--pdf                     also write a render-gated, size-optimized PDF
--pdf-preset <preset>     screen, standard, print, or raw
--single-pass             skip measurement (PowerPoint is not needed unless --pdf)
--offline                 use only the local Figma cache
--cache <dir>             choose the cache directory
--allow-font-fallback     continue after an explicit font substitution warning
--timings                 print every stage's wall time
```

The CLI renders a live spinner and elapsed time on a terminal. Redirected output becomes timestamped stage lines, with a heartbeat during blocking PowerPoint exports. The final line names the outputs, slide count, report, and total time.

PowerPoint must not have another presentation open. The converter checks first and refuses to touch unrelated documents. Each export also has a bounded timeout with a clear busy/dialog error.

## Library API

The CLI is a thin wrapper over the importable converter. A local bridge or another front end can consume the same progress events.

```ts
import {convertFigma} from 'figma2pptx';

const result = await convertFigma({
  target: 'file-key-or-url',
  frames: ['12:34', '56:78'],
  out: 'out/deck.pptx',
  pdf: 'screen',
  cacheDir: 'cache',
  onProgress: event => console.error(event.type, event.text),
});
```

The result includes output paths, the slide count, per-stage timings, font/build details, and line-placement measurements. `token`, `fonts`, and image operations are injectable for host applications and tests.

## How it works

1. Fetches the file tree, live-text SVG for each frame, and PNGs only for elements PowerPoint cannot draw.
2. Writes editable text with Figma's fixed line breaks, native DrawingML vectors, and raster fallbacks for unsupported effects.
3. Maps each Figma weight/style to a static installed face PowerPoint can select. Missing, substituted, and variable-only faces fail preflight by default.
4. Leaves fonts external in every export; no commercial font files are bundled or embedded.
5. Exports pass 1 with PowerPoint, measures every extractable line, applies per-text-box corrections, and exports pass 2.
6. Optionally optimizes the PDF and validates that resampling did not alter non-photo content.

Figma responses and renders live under `cache/`; conversion outputs belong under `out/`. Both are gitignored because source files may contain confidential content.

## Known limits

- Line breaks are locked for fidelity. Editing does not reflow paragraphs automatically.
- Shadows, masks, photo fills, angular/radial gradients, clipping frames, and dense decorative groups are rasterized.
- Fonts are never embedded. Fonts need a static installed face for PowerPoint's render pass; variable-only installations fail preflight.
- PowerPoint for Mac is required for measured pass 2 and PDF output. `--single-pass` can build the PPTX without it.
- Semi-transparent or rotated live text may be rasterized by PowerPoint's PDF exporter, so the report lists those lines as unmeasurable instead of claiming a placement pass.
- Reflowable text, font installation, and downstream presentation-pipeline integration are follow-ups.

## Benchmarking

The benchmark compares a PowerPoint-exported PDF with one Figma PDF per frame and the converter's report:

```sh
bun run bench out/deck.pdf \
  --report out/deck.report.json \
  --ref path/to/figma-frame-pdfs \
  --label run-name
```

It reports pixel difference, exact line breaks, the prior 0.75 pt comparison, the prototype-compatible one-pixel count, the corrected per-line one-pixel result, worst-case offset, and each unmeasurable or failing line by slide and cause. Add `--diffs out/diffs` for heat maps. Benchmark outputs are gitignored.

The committed test fixture is synthetic and network-free. `bun test` checks SVG-to-DrawingML geometry, line mapping, tracking/centring corrections, font selection, progress behavior, PowerPoint safety decisions, and byte-stable golden PPTX output.

## Figma Desktop plugin

The [plugin setup guide](plugin/README.md) covers the authenticated bridge, local plugin build, and one-time pairing. No bridge secret or Figma token ships with the plugin: each bridge installation generates its own random secret outside the repository.

## Development

```sh
bun test
bun run typecheck
```

The cache and golden package writer are deterministic. Update the golden only for a deliberate output-format change:

```sh
UPDATE_GOLDEN=1 bun test test/golden.test.ts
```

## License

MIT. See [LICENSE](LICENSE).
