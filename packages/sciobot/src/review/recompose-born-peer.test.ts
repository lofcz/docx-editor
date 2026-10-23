// The test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { caretAt } from '@docx-editor.dev/core/layout';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { groupToken } from '../../test/support/review-groups';
import { rejectedBornPeerBody } from '../../test/support/born-peer-oracle';
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
function wordBody(editor: DocxEditorInstance) {
  const read = readOoxmlPackage(writeOoxmlPackage(editor.surface!.session.currentPackage()));
  if (!read.ok) throw Error(read.reason);
  return canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!);
}
for (const checklist of [false, true])
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(`peer suffix inside a born ${checklist ? 'checkbox' : 'bullet'} survives later waves and both decisions; ${clients}/${receiver}`, async () => {
        const pair = await createTwoEditorFixture({
          document: docxFixture(source),
          documentId: context.documentId,
          ownedAttributes: true,
          clientIds: clients,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          const expectedBody = rejectedBornPeerBody(alice.surface!.session.currentPackage());
          const marker = checklist ? '- [ ] ' : '- ';
          const target = (one: string) => `Before\n\n${marker}${one}\n${marker}Two\n\nAfter`;
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
          const original = 'Before\n\n\u00a0PeerAfter';
          const proposed = `Before\n\n${marker}One Peer revised\n\n${marker}Two\n\nAfter`;
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
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check([original, `Before\n\n${marker}One Peer\n\n${marker}Two\n\nAfter`]);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check([original, proposed]);
          for (const action of ['accept', 'reject'] as const) {
            expect(bob.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
            await pair.settle();
            await check(action === 'accept' ? [proposed, proposed] : [original, original]);
            if (action === 'reject')
              for (const editor of pair.editors) expect(wordBody(editor)).toBe(expectedBody);
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
          const saved = readOoxmlPackage(new Uint8Array(await alice.save()));
          if (!saved.ok) throw Error(saved.reason);
          const reopened = propertyEditor(saved.package);
          try {
            expect(await views(reopened)).toEqual([original, proposed]);
          } finally {
            reopened.destroy();
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 30000);
