# preview

`preview` owns HTML shell generation for developer-visible document inspection.

## Browser flow

- render a self-contained HTML shell
- mount the current document HTML into `#preview-root`
- subscribe to `/events` for live updates

## Verification

Preview verification should confirm the root HTML changes after a publish event and that the browser-visible shell remains stable between updates.

## Current Parity Notes

- OfficeCLI emphasizes watch-driven preview as a developer workflow; `officekit` keeps the same watch/SSE model.
- The preview shell is shared infrastructure, while document rendering still comes from the format packages.
- Current differences versus OfficeCLI are mostly in rendering fidelity and long-tail visual semantics, not in the preview transport itself.
