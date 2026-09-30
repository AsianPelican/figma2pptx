# Changelog

All notable changes to this project are documented here.

## 0.2.0 — 2026-09-29

- Add the Figma Desktop export plugin and authenticated private-network bridge with per-install secrets, durable jobs, and verified downloads.
- Preserve source text order, lists, blank lines, per-run spacing, kerning, and localized font weights.
- Outline unavailable glyphs without reflow and measure each line against its remaining native text.
- Add host text transforms and safer PowerPoint container staging.
- Remove font embedding; fonts remain external in every export.
- Keep public fixtures synthetic and exclude local credentials, configured plugin builds, runtime data, and working records.

## 0.1.0 — 2026-09-28

- Adopt the deterministic prototype as a Bun/TypeScript package and CLI.
- Add a reusable conversion API, local cache, two-pass PowerPoint measurement, and PDF export.
- Add static-face font preflight and licence-aware font embedding.
- Add live progress, stage timings, PowerPoint busy/timeout safety, and fidelity benchmarking.
- Add synthetic offline fixtures, deterministic golden output, and unit coverage for conversion geometry and text behavior.
