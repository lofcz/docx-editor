// Evaluation only: prove an older built private SDK refuses pending and retained whole-table proposals.
import { strict as assert } from 'node:assert';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  atoms,
  context,
  decodePort,
  png,
  preparedCellDocument,
} from '../test/support/tracked-cell-atoms';
import { readEditorRevision } from '../src/projection/revision-cache';
import { recomposeOwnedProposals } from '../src/review/recompose';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../src/review/native-module';
import { readPropertyReviewStore } from '../src/review/property-record-store';
import { acceptAllReviewGroups } from '../src/review/accept-all';
import { rejectAllReviewGroups } from '../src/review/reject-all';
assert.ok(process.argv[2], 'Pass an older built SDK index.js path.');
const prior = await import(pathToFileURL(resolve(process.argv[2])).href);
let checks = 0;
for (const decision of ['accept', 'reject'] as const) {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: preparedCellDocument(
      '<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>'
    ),
    imageDecodePort: decodePort,
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
  const token = () => readEditorRevision(editor, context.documentId, context.epoch)!;
  const check = async () => {
    const store = readPropertyReviewStore(editor.surface!.session.currentPackage());
    assert.equal(store.status, 'ready');
    if (store.status !== 'ready') return;
    const raw = store.retainedRecords.get('sibling-older-reader')!.raw;
    assert.equal(JSON.parse(raw).schemaVersion, 6);
    const view = await prior.readWorkingView(editor.surface!.session, token());
    assert.equal(view.success, false);
    assert.match(view.error, /unsupported|unknown/i);
    checks++;
  };
  try {
    const staged = await recomposeOwnedProposals(
      editor,
      {
        revision: token(),
        targetMarkdown: `Before\n\nIntro\n\n| A | B |\n| - | - |\n| ${atoms} | Tail |\n\nEnd\n\nAfter`,
        ownership: [],
      },
      context,
      {
        payloads: { presence: { proposalId: 'sibling-older-reader' } },
        images: { decodePort, resolve: async () => ({ bytes: png, mime: 'image/png' }) },
      }
    );
    assert.equal(staged.success, true);
    await check();
    const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
      editor,
      token(),
      context
    );
    assert.equal(decided.success, true);
    await check();
  } finally {
    editor.destroy();
  }
}
console.log(
  `Older built SDK refuses all ${checks} pending and retained whole-table-presence snapshots.`
);
