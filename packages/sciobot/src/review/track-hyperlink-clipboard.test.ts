// The existing test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  findNode,
  paragraphOffsetIndex,
  readOoxmlPart,
  type OoxmlNode,
  type OoxmlParagraphNode,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  document_,
  context,
  plan,
  mount,
  fingerprint,
} from '../../test/support/link-structure-execution';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackHyperlinkPlan } from './track-hyperlinks';
import { applyDirectPlan } from '../ai/apply-direct';
import { runWithTransactionActor } from '@docx-editor.dev/core/store/package/actor-scoped-ids';
const target = 'Alpha [new](https://example.org/new) end\n\nTail';
function fragment(children?: OoxmlParagraphNode['children']): OoxmlParagraphNode {
  const r = readOoxmlPart(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr/><w:t xml:space="preserve">Alpha </w:t></w:r></w:p></w:body></w:document>',
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!r.ok) throw Error(r.reason);
  const body = r.part.root.children.find((n) => n.kind === 'body')!;
  if (body.kind !== 'body') throw Error('body');
  const p = body.children[0]!;
  if (p.kind !== 'paragraph') throw Error('paragraph');
  return children ? { ...p, children } : p;
}
function paste(editor: ReturnType<typeof mount>, block: OoxmlNode) {
  const id = editor.surface!.session.paragraphIds()[0]!,
    p = findNode(editor.surface!.session.part(), id)!;
  if (p.kind !== 'paragraph') throw Error('p');
  return editor.surface!.applyAutomationOps(
    () => [
      {
        op: 'insertFragment',
        paragraphId: id,
        offset: paragraphOffsetIndex(p).length,
        blocks: [block],
      },
    ],
    { kind: 'body' },
    [],
    false,
    true
  );
}
for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const action of ['accept', 'reject'] as const)
      test(`independent copied source ${clientIds}/${receiver}/${action}`, async () => {
        const pair = await createTwoEditorFixture({
          document: document_(),
          documentId: context.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          clientIds,
        });
        const golden = mount();
        try {
          const [alice, bob] = pair.editors;
          const result = trackHyperlinkPlan(alice!, await plan(alice!, target), context);
          if (!result.success) throw Error(result.error);
          await pair.settle();
          const p = findNode(
            bob!.surface!.session.part(),
            bob!.surface!.session.paragraphIds()[0]!
          )!;
          if (p.kind !== 'paragraph') throw Error('p');
          const source = p.children.find((n) => n.kind === 'run')!;
          pair.pause();
          expect(paste(bob!, fragment([source]))).toMatchObject({
            committed: true,
            rejected: false,
          });
          pair.reconnect(receiver);
          await pair.settle();
          const pendingBeforeDecision = fingerprint(alice!);
          if (action === 'accept') {
            const directPlan = await plan(golden, target);
            const direct = runWithTransactionActor('alice', () =>
              applyDirectPlan(golden, directPlan, context)
            );
            if (!direct.success) throw Error(direct.error);
          }
          const beforePaste = fingerprint(golden);
          expect(paste(golden, fragment())).toMatchObject({
            committed: true,
            rejected: false,
          });
          const afterPaste = fingerprint(golden);
          expect(alice!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(fingerprint(editor)).toBe(afterPaste);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(fingerprint(editor)).toBe(beforePaste);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(fingerprint(editor)).toBe(afterPaste);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(fingerprint(editor)).toBe(pendingBeforeDecision);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(fingerprint(editor)).toBe(afterPaste);
        } finally {
          pair.destroy();
          golden.destroy();
        }
      }, 30000);
