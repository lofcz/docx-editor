import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../packages/sciobot/test/support/document';
import { withPayloadRecord } from '../../packages/sciobot/src/blocks/payload-opc-codec';
import { recomposeOwnedProposals } from '../../packages/sciobot/src/review/recompose';
import { readWorkingView } from '../../packages/sciobot/src/review/working-view';
import { readCommittedView } from '../../packages/sciobot/src/review/committed-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../../packages/sciobot/src/review/native-module';
import { readEditorRevision } from '../../packages/sciobot/src/projection/revision-cache';
const previous = process.argv[2];
if (!previous) throw Error('Pass the pinned previous adapter index.js path.');
const old = await import(previous);
const loaded = readOoxmlPackage(
  docxFixture(
    '<w:p><w:r><w:t>Lead </w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="scio:1:source"/></w:sdtPr><w:sdtContent><w:r><w:t>Book</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> tail</w:t></w:r></w:p>'
  )
);
if (!loaded.ok) throw Error(loaded.reason);
const stored = withPayloadRecord(
  loaded.package,
  {
    version: 1,
    id: 'source',
    type: 'source',
    attributes: { title: 'Book', url: 'https://source.test', trust: 'verified' },
  },
  'Citation'
);
if (!stored.ok) throw Error(stored.reason);
const editor = createDocxEditor({
  container: document.createElement('div'),
  document: writeOoxmlPackage(stored.pkg),
  modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
});
const context = {
  documentId: 'earlier-reader',
  epoch: 'native',
  author: 'AI',
  date: '2026-10-26T12:00:00Z',
};
const token = () => readEditorRevision(editor, context.documentId, context.epoch)!;
try {
  const initial = await readWorkingView(editor.surface!.session, token());
  if (!initial.success) throw Error(initial.error);
  const target = initial.view.projection.text.replace('title="Book"', 'title="Final"');
  const result = await recomposeOwnedProposals(
    editor,
    { revision: token(), targetMarkdown: target, ownership: [], contentProofs: [] },
    context,
    {
      preparePayloads: (targets) => ({
        atomic: {
          proposalId: 'source-review',
          fallbackTargets: targets.map((t) => ({
            blockId: t.blockId,
            before: t.beforeText,
            after: 'Final',
          })),
        },
      }),
    }
  );
  if (!result.success) throw Error(result.error);
  for (const [before, current, expected] of [
    [old.readWorkingView, readWorkingView, target],
    [old.readCommittedView, readCommittedView, initial.view.projection.text],
  ] as const) {
    const a = await before(editor.surface!.session, token()),
      b = await current(editor.surface!.session, token());
    if (
      !a.success ||
      !b.success ||
      a.view.projection.text !== expected ||
      a.view.projection.text !== b.view.projection.text
    )
      throw Error('Earlier reader changed citation meaning');
  }
  console.log(
    'Pinned previous reader preserves exact original and proposed citation views from the repaired writer.'
  );
} finally {
  editor.destroy();
}
