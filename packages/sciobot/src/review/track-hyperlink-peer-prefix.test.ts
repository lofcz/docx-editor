// The existing test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  document_,
  context,
  plan,
  projection,
  ast,
} from '../../test/support/link-structure-execution';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackHyperlinkPlan } from './track-hyperlinks';

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const action of ['accept', 'reject'] as const)
      test(`link proposal preserves peer prefix ${clientIds}/${receiver}/${action}`, async () => {
        const pair = await createTwoEditorFixture({
          document: document_(),
          documentId: context.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          clientIds,
        });
        try {
          const [alice, bob] = pair.editors;
          pair.pause();
          const result = trackHyperlinkPlan(
            alice!,
            await plan(alice!, 'Alpha [new](https://example.org/new) end\n\nTail'),
            context
          );
          if (!result.success) throw Error(result.error);
          const paragraphId = bob!.surface!.session.paragraphIds()[0]!;
          expect(
            bob!.surface!.applyAutomationOps(
              () => [{ op: 'insertText', paragraphId, offset: 1, text: 'PEER' }],
              { kind: 'body' },
              [],
              false,
              true
            )
          ).toMatchObject({ committed: true, rejected: false });
          pair.reconnect(receiver);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(ast((await projection(editor, true)).text)).toBe(ast('APEERlpha end\n\nTail'));
            expect(ast((await projection(editor)).text)).toBe(
              ast('APEERlpha [new](https://example.org/new) end\n\nTail')
            );
          }
          expect(alice!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          await pair.settle();
          const wanted =
            action === 'accept'
              ? 'APEERlpha [new](https://example.org/new) end\n\nTail'
              : 'APEERlpha end\n\nTail';
          for (const editor of pair.editors)
            expect(ast((await projection(editor)).text)).toBe(ast(wanted));
        } finally {
          pair.destroy();
        }
      }, 30000);
