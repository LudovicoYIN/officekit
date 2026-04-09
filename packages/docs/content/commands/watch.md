# watch

`watch` is the live-preview runtime surface for officekit.

## Browser flow

- preview server serves the current HTML document
- SSE delivers replacement events
- file watching triggers re-rendering through caller-provided renderers
- browser auto-open is attempted by default for interactive watch sessions
- `unwatch` is the matching teardown surface at the CLI layer

## Design note

Like OfficeCLI's watch server, the runtime surface should avoid opening or mutating office files directly. Rendering belongs to upstream format handlers; the watch layer is transport and orchestration.

## Current Parity Notes

- Aligned with OfficeCLI on local HTTP preview, SSE refresh, and browser-facing live iteration.
- Browser launch now follows the OfficeCLI-style default of opening the preview URL when possible; `--no-open` keeps headless workflows quiet.
- `officekit` currently treats `watch`/`unwatch` as explicit CLI surfaces rather than hiding teardown behind process management alone.
- Long-tail rendering fidelity still depends on format-specific HTML/SVG renderers, so visual output can differ from OfficeCLI even when the transport layer is aligned.
