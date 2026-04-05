export const SUPPORTED_PREVIEW_EXTENSIONS = ['.docx', '.xlsx', '.pptx'];
export const DEFAULT_PREVIEW_PORT = 18080;

export function assertPreviewExtension(extension) {
  if (!SUPPORTED_PREVIEW_EXTENSIONS.includes(extension)) {
    throw new Error(`Unsupported preview format: ${extension}`);
  }
}

export function buildWaitingState({ productName = 'officekit' } = {}) {
  return `<div class="msg"><h2>Waiting for first update...</h2><p>Run an ${productName} command to see the preview.</p></div>`;
}

export function buildEventStreamClient({ eventStreamPath = '/events' } = {}) {
  return `
const source = new EventSource(${JSON.stringify(eventStreamPath)});
source.addEventListener('update', (event) => {
  const payload = JSON.parse(event.data);
  if (payload.title) document.title = payload.title;
  if (payload.html) {
    const mount = document.getElementById('preview-root');
    mount.innerHTML = payload.html;
  }
});
`;
}

export function createPreviewEnvelope({ fileName, extension, bodyHtml = '', title } = {}) {
  assertPreviewExtension(extension);
  return {
    product: 'officekit',
    fileName,
    extension,
    title: title ?? `officekit preview · ${fileName}`,
    html: bodyHtml,
    transport: 'sse'
  };
}

export function renderPreviewDocument({
  title = 'officekit preview',
  bodyHtml = '',
  eventStreamPath = '/events',
  productName = 'officekit',
  waitingState = buildWaitingState({ productName })
} = {}) {
  const mountHtml = bodyHtml || waitingState;
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta name="generator" content="${productName}" />
    <title>${title}</title>
    <style>
      :root { color-scheme: dark; font-family: Inter, system-ui, sans-serif; }
      body { margin: 0; background: #0f172a; color: #e2e8f0; }
      #preview-root { min-height: 100vh; padding: 24px; }
      .msg { display: grid; min-height: calc(100vh - 48px); place-items: center; text-align: center; }
      .msg h2 { margin-bottom: 0.5rem; }
      a { color: #93c5fd; }
    </style>
  </head>
  <body>
    <main id="preview-root">${mountHtml}</main>
    <script type="module">${buildEventStreamClient({ eventStreamPath })}</script>
  </body>
</html>`;
}
