import { buildBlockIndex } from '../blocks/block-index';
// The test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { propertyRecordFixture } from '../../test/support/property-record';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { recomposeOwnedProposals } from './recompose';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';

async function views(editor: DocxEditorInstance) {
  const before = await readCommittedView(editor.surface!.session, propertyToken(editor));
  const after = await readWorkingView(editor.surface!.session, propertyToken(editor));
  if (!before.success || !after.success) throw Error('Missing native property views');
  return [before.view.projection.text, after.view.projection.text];
}
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`property replacement waves retain peer input, owned history, decisions, geometry, and saved convergence; ${clients}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(propertyRecordFixture()),
        documentId: propertyContext.documentId,
        ownedAttributes: true,
        clientIds: clients,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const initial = (await views(alice))[1]!;
        const baselinePayload = buildBlockIndex(alice.surface!.session.currentPackage()).byId.get(
          'activity'
        )!.payload;
        const target = (name: string, body: string) =>
          initial
            .replace('name="Old activity"', `name="${name} activity"`)
            .replace('Body', `Body ${body}`);
        const options = (name: string) => ({
          payloads: {
            properties: {
              proposalId: `shared-${name.toLowerCase()}`,
              fallbackTargets: [
                {
                  blockId: 'activity',
                  before: 'Old activity — 10 minutes',
                  after: `${name} activity — 10 minutes`,
                },
              ],
            },
          },
        });
        const first = await recomposeOwnedProposals(
          alice,
          {
            revision: propertyToken(alice),
            targetMarkdown: target('First', 'first'),
            ownership: [],
            contentProofs: [],
          },
          propertyContext,
          options('First')
        );
        if (!first.success) throw Error(first.error);
        await pair.settle();
        for (const editor of pair.editors)
          expect(
            buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('activity')!.payload
          ).toEqual(baselinePayload);
        pair.pause();
        const paragraphId = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Outside')!;
        expect(
          bob.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId, offset: 7, text: ' peer' },
          ]).rejected
        ).toBe(false);
        const second = await recomposeOwnedProposals(
          alice,
          {
            revision: propertyToken(alice),
            targetMarkdown: target('Second', 'second'),
            ownership: first.ownership,
            contentProofs: first.contentProofs,
          },
          propertyContext,
          options('Second')
        );
        if (!second.success) throw Error(second.error);
        pair.reconnect(receiver);
        await pair.settle();
        const original = initial.replace('Outside', 'Outside peer');
        const proposed = target('Second', 'second').replace('Outside', 'Outside peer');
        const check = async (expected: readonly string[]) => {
          for (const editor of pair.editors) expect(await views(editor)).toEqual([...expected]);
        };
        await check([original, proposed]);
        bob.surface!.setSelection({
          anchor: { paragraphId, offset: 12 },
          head: { paragraphId, offset: 12 },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (value) => value.actorId === 'bob'
        );
        expect(remote?.head.nodeId).toBe(paragraphId);
        expect(remote?.head.offset).toBe(12);
        const carets = pair.editors.map((editor) =>
          caretAt(editor.surface!.layout(), { paragraphId, offset: 12 })
        );
        expect(carets[0]).not.toBeNull();
        expect(carets[1]).not.toBeNull();
        expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({ x: carets[1]!.x, y: carets[1]!.y });
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check([original, target('First', 'first').replace('Outside', 'Outside peer')]);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check([original, proposed]);
        for (const action of ['accept', 'reject'] as const) {
          const decided = (action === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            bob,
            propertyToken(bob),
            propertyContext
          );
          if (!decided.success) throw Error(decided.error);
          await pair.settle();
          await check(action === 'accept' ? [proposed, proposed] : [original, original]);
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check([original, proposed]);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
        }
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 30000);
