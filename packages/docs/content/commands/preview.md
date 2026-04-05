# preview

`preview` owns HTML shell generation for developer-visible document inspection.

## Browser flow

- render a self-contained HTML shell
- mount the current document HTML into `#preview-root`
- subscribe to `/events` for live updates

## Verification

Preview verification should confirm the root HTML changes after a publish event and that the browser-visible shell remains stable between updates.
