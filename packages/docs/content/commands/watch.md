# watch

`watch` is the live-preview runtime surface for officekit.

## Browser flow

- preview server serves the current HTML document
- SSE delivers replacement events
- file watching triggers re-rendering through caller-provided renderers

## Design note

Like OfficeCLI's watch server, the runtime surface should avoid opening or mutating office files directly. Rendering belongs to upstream format handlers; the watch layer is just transport and orchestration.
