# Changelog

## 0.2.1

- Enabled the development-plugin file-key API and added explicit disabled-export reasons.
- Made bridge jobs durable across stream disconnects and service restarts, with status reattachment and lifecycle logging.
- Disabled PowerPoint font embedding at every plugin and bridge request boundary.
- Fixed configured builds incorrectly reporting that the bridge origin was still a placeholder.
- Added a zero-click macOS permission doctor with a bounded container-staged PowerPoint export.
- Made SSH-context Full Disk Access informational; launcher-context permissions remain the export gate.
- Auto-connect stored bridge pairings and retry transient network failures with capped backoff.
- Guard Desktop refreshes against test origins and require authenticated live health before copying.
- Added a pure-JavaScript SHA-256 fallback for Figma's null-origin UI iframe.
- Removed the unsafe `/tmp` PowerPoint staging fallback when Full Disk Access is missing.

## 0.2.0 — 2026-09-28

- Move the bridge and all conversion work to a configurable private-network Mac.
- Add authenticated, job-scoped PPTX/PDF/report downloads with byte-count and SHA-256 verification.
- Add the client-side thin CLI, owner-only remote environment loader, and bridge launch-agent installer.
- Isolate each export in a cancellable CLI subprocess and keep bridge state in a dedicated sibling runtime tree.
- Configure the Figma development manifest and UI for one private bridge origin and add explicit download buttons.
- Package each configured client build as a self-contained `local/` import folder with a real `manifest.json`.
- Reject dotless bridge hostnames before Figma rejects the generated development manifest.
- Serve the private bridge over TLS and allow Chromium Private Network Access preflights.
- Give the macOS Automation launcher an explicit persistent designated requirement.
- Stop an accidentally imported source manifest before it fetches the placeholder origin.

## 0.1.1 — 2026-09-28

- Correct the local-development manifest: use a local slug ID, deny published network access, and allow only the localhost bridge through `devAllowedDomains`.
- Validate all documented manifest fields, network scope, and referenced files as part of every plugin build.

## 0.1.0 — 2026-09-28

- Add a Figma Desktop selection/export UI with live frame names and conversion progress.
- Add a loopback-only authenticated Bun bridge with a dedicated sibling runtime tree for config, cache, jobs, and logs.
- Add selection mapping and bridge protocol tests plus exact local installation instructions.
