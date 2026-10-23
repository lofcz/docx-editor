import * as Y from 'yjs';
// The test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  paragraphTextOf,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  readOoxmlPart,
  replaceChildren,
  insertChildren,
  bodyStoryRoot,
  withPart,
} from '@docx-editor.dev/core/store';
import { caretAt } from '@docx-editor.dev/core/layout';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { groupToken } from '../../test/support/review-groups';
import { bornWordBody, paraIds, wordPart } from '../../test/support/born-word-body';
import { planTargetDocument } from '../diff/plan-target';
import { applyDirectPlan } from '../ai/apply-direct';
import { recomposeOwnedProposals } from './recompose';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';

const context = { ...propertyContext, date: '2026-10-23T00:00:00Z' };
const source = '<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>';
async function views(editor: DocxEditorInstance) {
  const original = await readCommittedView(editor.surface!.session, groupToken(editor));
  const proposed = await readWorkingView(editor.surface!.session, groupToken(editor));
  if (!original.success || !proposed.success) throw Error('Missing review views');
  return [original.view.projection.text.trimEnd(), proposed.view.projection.text.trimEnd()];
}
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`peer suffix inside a born table survives replacement, decisions, history, carets, and Word reopen; ${clients}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: docxFixture(source),
        documentId: context.documentId,
        ownedAttributes: true,
        clientIds: clients,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const originalPackage = alice.surface!.session.currentPackage();
        const imported = new Set(paraIds(wordPart(alice.surface!.session.currentPackage()).root));
        const target = (one: string) =>
          `Before\n\n| ${one} | Two |\n| --- | --- |\n| Three | Four |\n\nAfter`;
        const first = await recomposeOwnedProposals(
          alice,
          {
            revision: groupToken(alice),
            targetMarkdown: target('One'),
            ownership: [],
            contentProofs: [],
          },
          context
        );
        if (!first.success) throw Error(first.error);
        await pair.settle();
        pair.pause();
        const paragraphId = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob.surface!.session.part(), id)?.endsWith('One'))!;
        expect(paragraphId).toBeDefined();
        const offset = paragraphTextOf(bob.surface!.session.part(), paragraphId)!.length;
        expect(
          bob.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId, offset, text: ' Peer' },
          ]).rejected
        ).toBe(false);
        const second = await recomposeOwnedProposals(
          alice,
          {
            revision: groupToken(alice),
            targetMarkdown: target('One revised'),
            ownership: first.ownership,
            contentProofs: first.contentProofs,
          },
          context
        );
        if (!second.success) throw Error(second.error);
        pair.reconnect(receiver);
        await pair.settle();
        // Original shows the independent peer text at the same boundary as rejection.
        const original = 'Before\n\n\u00a0Peer\n\nAfter';
        const proposed = target('One Peer revised');
        const check = async (expected: readonly string[]) => {
          for (const editor of pair.editors) expect(await views(editor)).toEqual([...expected]);
        };
        await check([original, proposed]);
        // Wave two replaces the born paragraph. A user selects the new editable owner,
        // rather than its retired predecessor, whose retained address has no live text.
        const liveParagraphId = bob
          .surface!.session.paragraphIds()
          .find((id) =>
            paragraphTextOf(bob.surface!.session.part(), id)?.includes('One Peer revised')
          )!;
        expect(liveParagraphId).toBeDefined();
        const liveText = paragraphTextOf(bob.surface!.session.part(), liveParagraphId)!;
        const liveOffset = liveText.indexOf('One Peer') + 'One Peer'.length;
        bob.surface!.setSelection({
          anchor: { paragraphId: liveParagraphId, offset: liveOffset },
          head: { paragraphId: liveParagraphId, offset: liveOffset },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (value) => value.actorId === 'bob'
        );
        expect(remote?.head.nodeId).toBe(liveParagraphId);
        expect(remote?.head.offset).toBe(liveOffset);
        const carets = pair.editors.map((editor) =>
          caretAt(editor.surface!.layout(), { paragraphId: liveParagraphId, offset: liveOffset })
        );
        expect(carets[0]).not.toBeNull();
        expect(carets[1]).not.toBeNull();
        expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({
          x: carets[1]!.x,
          y: carets[1]!.y,
        });
        const unchanged = pair.editors.map((editor) =>
          canonicalOoxmlFingerprint(editor.surface!.session.part())
        );
        for (const editor of pair.editors) editor.surface!.setRevisionDisplayMode('original');
        await pair.settle();
        const originalCarets = pair.editors.map((editor) =>
          caretAt(editor.surface!.layout(), { paragraphId: liveParagraphId, offset: liveOffset })
        );
        expect(originalCarets[0]).not.toBeNull();
        expect(originalCarets[1]).not.toBeNull();
        expect({ x: originalCarets[0]!.x, y: originalCarets[0]!.y }).toEqual({
          x: originalCarets[1]!.x,
          y: originalCarets[1]!.y,
        });
        for (const [index, editor] of pair.editors.entries()) {
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(unchanged[index]!);
          editor.surface!.setRevisionDisplayMode('all-markup');
        }
        await pair.settle();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check([original, target('One Peer')]);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check([original, proposed]);
        const loaded = readOoxmlPackage(docxFixture(source));
        if (!loaded.ok) throw Error(loaded.reason);
        const direct = propertyEditor(loaded.package);
        const projection = await readWorkingView(direct.surface!.session, groupToken(direct));
        if (!projection.success) throw Error(projection.error);
        const plan = planTargetDocument(
          {
            pkg: direct.surface!.session.currentPackage(),
            revision: groupToken(direct),
            projection: projection.view.projection,
          },
          proposed
        );
        if (!plan.success) throw Error(plan.error);
        const applied = applyDirectPlan(direct, plan.plan, context);
        if (!applied.success) throw Error(applied.error);
        // The AI prefix, peer insertion, and AI suffix retain three independent native runs.
        // Author these boundaries explicitly; do not normalize the produced Word tree.
        const split = readOoxmlPart(
          '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>One</w:t></w:r><w:r><w:t xml:space="preserve"> Peer</w:t></w:r><w:r><w:t xml:space="preserve"> revised</w:t></w:r></w:p>',
          {
            name: '/expected-peer-runs.xml',
            contentType: direct.surface!.session.part().contentType,
          }
        );
        if (!split.ok) throw Error(split.reason);
        const firstCellParagraph = direct
          .surface!.session.paragraphIds()
          .find(
            (id) => paragraphTextOf(direct.surface!.session.part(), id) === 'One Peer revised'
          )!;
        const expectedPart = replaceChildren(
          direct.surface!.session.part(),
          firstCellParagraph,
          split.part.root.children
        );
        if (!expectedPart.ok) throw Error(JSON.stringify(expectedPart.issues));
        const acceptedPackage = withPart(
          direct.surface!.session.currentPackage(),
          expectedPart.part
        );
        const peerSeed = readOoxmlPart(
          '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="ABCDEF01" w14:textId="ABCDEF01"><w:pPr/><w:r><w:t xml:space="preserve"> Peer</w:t></w:r></w:p>',
          {
            name: '/expected-recovered-peer.xml',
            contentType: direct.surface!.session.part().contentType,
          }
        );
        if (!peerSeed.ok) throw Error(peerSeed.reason);
        const originalPart = originalPackage.parts.get(originalPackage.mainDocumentPart)!;
        const originalBody = bodyStoryRoot(originalPart);
        if (!originalBody || originalBody.kind === 'textValue')
          throw Error('Missing original story');
        const rejectedPart = insertChildren(originalPart, originalBody.id, 1, [peerSeed.part.root]);
        if (!rejectedPart.ok) throw Error(JSON.stringify(rejectedPart.issues));
        const rejectedPackage = withPart(originalPackage, rejectedPart.part);
        try {
          for (const action of ['accept', 'reject'] as const) {
            expect(bob.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
            await pair.settle();
            await check(
              action === 'accept'
                ? [proposed, proposed]
                : ['Before\n\n\u00a0Peer\n\nAfter', 'Before\n\n\u00a0Peer\n\nAfter']
            );
            if (action === 'reject') {
              for (const [index, editor] of pair.editors.entries()) {
                expect(paragraphTextOf(editor.surface!.session.part(), liveParagraphId)).toBe(
                  ' Peer'
                );
                const recoveredCaret = caretAt(editor.surface!.layout(), {
                  paragraphId: liveParagraphId,
                  offset: ' Peer'.length,
                });
                expect(recoveredCaret).not.toBeNull();
                expect({ x: recoveredCaret!.x, y: recoveredCaret!.y }).toEqual({
                  x: originalCarets[index]!.x,
                  y: originalCarets[index]!.y,
                });
              }
            }
            for (const editor of pair.editors)
              expect(bornWordBody(editor.surface!.session.currentPackage(), imported)).toBe(
                bornWordBody(action === 'accept' ? acceptedPackage : rejectedPackage, imported)
              );
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
        } finally {
          direct.destroy();
        }
        const saved = readOoxmlPackage(new Uint8Array(await alice.save()));
        if (!saved.ok) throw Error(saved.reason);
        const reopened = propertyEditor(saved.package);
        try {
          expect(await views(reopened)).toEqual([original, proposed]);
        } finally {
          reopened.destroy();
        }
        const room = await createTwoEditorFixture({
          document: new Uint8Array(await alice.save()),
          roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
          documentId: context.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          await room.settle();
          for (const editor of room.editors)
            expect(await views(editor)).toEqual([original, proposed]);
          expect(
            room.editors[1]!.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok
          ).toBe(true);
          await room.settle();
          for (const editor of room.editors) {
            expect(await views(editor)).toEqual([original, original]);
            expect(bornWordBody(editor.surface!.session.currentPackage(), imported)).toBe(
              bornWordBody(rejectedPackage, imported)
            );
          }
          const digests = await room.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          room.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          room.destroy();
        }
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 30000);
