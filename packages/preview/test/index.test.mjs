import test from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_PREVIEW_PORT, SUPPORTED_PREVIEW_EXTENSIONS, buildWaitingState, createPreviewEnvelope, renderPreviewDocument } from '../src/index.mjs';

test('preview shell keeps OfficeCLI watch semantics with officekit branding', () => {
  const html = renderPreviewDocument({ title: 'deck preview', eventStreamPath: '/watch/events' });
  assert.match(html, /EventSource\("\/watch\/events"\)/);
  assert.match(html, /Waiting for first update/);
  assert.match(html, /Run an officekit command to see the preview/);
  assert.match(html, /id="preview-root"/);
  assert.equal(DEFAULT_PREVIEW_PORT, 18080);
});

test('preview envelope accepts supported document extensions only', () => {
  const envelope = createPreviewEnvelope({ fileName: 'demo.pptx', extension: '.pptx', bodyHtml: '<section>ok</section>' });
  assert.equal(envelope.transport, 'sse');
  assert.deepEqual(SUPPORTED_PREVIEW_EXTENSIONS, ['.docx', '.xlsx', '.pptx']);
  assert.match(buildWaitingState(), /officekit command/);
  assert.throws(() => createPreviewEnvelope({ fileName: 'demo.pdf', extension: '.pdf' }), /Unsupported preview format/);
});
