// The test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { atomicFixture, atomicVisual, atomicListening } from '../../test/support/atomic-execution';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import type { JsonObject } from '../data/json';
import { buildBlockIndex } from '../blocks/block-index';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { recomposeOwnedProposals } from './recompose';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';

const editorOptions = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
async function views(editor: DocxEditorInstance) {
  const before = await readCommittedView(editor.surface!.session, propertyToken(editor));
  const after = await readWorkingView(editor.surface!.session, propertyToken(editor));
  if (!before.success || !after.success) throw Error('Missing native atomic views');
  return [before.view.projection.text, after.view.projection.text];
}
const payloads = (editor: DocxEditorInstance) => {
  const index = buildBlockIndex(editor.surface!.session.currentPackage());
  return ['visual', 'audio'].map((id) => index.byId.get(id)!.payload);
};
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`complete atomic targets retain peer input, native ownership, decisions, geometry, and room reopen; ${clients}/${receiver}`, async () => {
      const opaque = {
        ...atomicListening,
        mdast: {
          ...(atomicListening.mdast as JsonObject),
          name: 'NovelReactPanel',
          retainedRootData: { keep: 'custom-codec' },
        },
      };
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(atomicFixture([atomicVisual, opaque])),
        documentId: propertyContext.documentId,
        ownedAttributes: true,
        clientIds: clients,
        editor: editorOptions,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const initial = (await views(alice))[1]!;
        const target = (wave: string) =>
          initial
            .replace('vfs://pristine.png', `vfs://${wave}.png`)
            .replace('Old picture', `${wave} picture`)
            .replace('original-audio', `${wave}-audio`)
            .replace('Good morning!', `${wave} line!`)
            .replace('Phase body', `Phase ${wave}`);
        const options = (wave: string) => ({
          payloads: {
            atomic: {
              proposalId: `shared-target-${wave}`,
              fallbackTargets: ['visual', 'audio'].map((blockId) => ({
                blockId,
                before: `Old ${blockId} fallback`,
                after: `${wave} ${blockId} fallback`,
              })),
            },
          },
        });
        const first = await recomposeOwnedProposals(
          alice,
          {
            revision: propertyToken(alice),
            targetMarkdown: target('first'),
            ownership: [],
            contentProofs: [],
          },
          propertyContext,
          options('first')
        );
        if (!first.success) throw Error(first.error);
        await pair.settle();
        for (const editor of pair.editors) expect(payloads(editor)).toEqual([atomicVisual, opaque]);
        pair.pause();
        const paragraphId = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Keep tail')!;
        expect(
          bob.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId, offset: 9, text: ' peer' },
          ]).rejected
        ).toBe(false);
        const second = await recomposeOwnedProposals(
          alice,
          {
            revision: propertyToken(alice),
            targetMarkdown: target('second'),
            ownership: first.ownership,
            contentProofs: first.contentProofs,
          },
          propertyContext,
          options('second')
        );
        if (!second.success) throw Error(second.error);
        pair.reconnect(receiver);
        await pair.settle();
        const original = initial.replace('Keep tail', 'Keep tail peer');
        const proposed = target('second').replace('Keep tail', 'Keep tail peer');
        const check = async (expected: readonly string[]) => {
          for (const editor of pair.editors) expect(await views(editor)).toEqual([...expected]);
        };
        await check([original, proposed]);
        bob.surface!.setSelection({
          anchor: { paragraphId, offset: 14 },
          head: { paragraphId, offset: 14 },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (value) => value.actorId === 'bob'
        );
        expect(remote?.head.nodeId).toBe(paragraphId);
        expect(remote?.head.offset).toBe(14);
        const carets = pair.editors.map((editor) =>
          caretAt(editor.surface!.layout(), { paragraphId, offset: 14 })
        );
        expect(carets[0]).not.toBeNull();
        expect(carets[1]).not.toBeNull();
        expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({ x: carets[1]!.x, y: carets[1]!.y });
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check([original, target('first').replace('Keep tail', 'Keep tail peer')]);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check([original, proposed]);
        for (const action of ['accept', 'reject'] as const) {
          const result = (action === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            bob,
            propertyToken(bob),
            propertyContext
          );
          if (!result.success) throw Error(result.error);
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
        const reopened = await createTwoEditorFixture({
          document: new Uint8Array(await alice.save()),
          documentId: propertyContext.documentId,
          roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
          ownedAttributes: true,
          editor: editorOptions,
        });
        try {
          for (const editor of reopened.editors) {
            expect(await views(editor)).toEqual([original, proposed]);
            expect(payloads(editor)).toEqual([atomicVisual, opaque]);
          }
        } finally {
          reopened.destroy();
        }
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 30000);
