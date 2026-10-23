import { caretAt } from '@docx-editor.dev/core/layout';
// Test-only collaboration transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  TreePackageStore,
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import {
  preparedApplicationDocument,
  applicationFragmentTarget as target,
  applicationFragmentOptions as options,
} from '../../test/support/application-fragments';
import { previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { recomposeOwnedProposals } from './recompose';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { readReviewGroups } from './groups';
import { parseTargetMarkdown } from '../projection/parse-target';
import { targetAstKey } from '../diff/target-key';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
const word = (pkg: OoxmlPackage) => {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return canonicalOoxmlFingerprint(read.package.parts.get(pkg.mainDocumentPart)!);
};
const key = (text: string) => {
  const parsed = parseTargetMarkdown(text);
  if (!parsed.success) throw Error(parsed.error);
  return targetAstKey(parsed.target.normalizedAst);
};
// Use the same documented actor scope for independent native resource allocation.
function actorDirect(pkg: OoxmlPackage, plan: Parameters<typeof previewDirectPlan>[1]) {
  const store = new TreePackageStore(pkg, pkg.parts.get(pkg.mainDocumentPart)!);
  const result = store.transact(
    { kind: 'body' },
    (ctx) => {
      if (
        !ctx.applyPackage((working) => {
          const preview = previewDirectPlan(working, plan);
          if (!preview.success) throw Error(preview.error);
          return preview.pkg;
        })
      )
        throw Error('Independent direct transaction refused');
    },
    { recordsHistory: false, actorId: 'alice' }
  );
  if (!result.ok) throw Error(result.reason);
  return store.currentPackage();
}
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`AI application fragments preserve peer text, ownership, both histories, ${decision}, exact Word output, and room reopening; ${clients}/${receiver}`, async () => {
        const fixture = {
          document: preparedApplicationDocument(
            docxFixture('<w:p><w:r><w:t>Outside</w:t></w:r></w:p>')
          ),
          documentId: propertyContext.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        };
        const pair = await createTwoEditorFixture({ ...fixture, clientIds: clients });
        try {
          const [alice, bob] = pair.editors;
          const baseline = alice!.surface!.session.currentPackage();
          const id = alice!.surface!.session.paragraphIds()[0]!;
          const initial = await readWorkingView(alice!.surface!.session, propertyToken(alice!));
          if (!initial.success) throw Error(initial.error);
          const firstDirectPlan = planTargetDocument(
            { pkg: baseline, revision: propertyToken(alice!), projection: initial.view.projection },
            target('First'),
            options('first')
          );
          if (!firstDirectPlan.success) throw Error(firstDirectPlan.error);
          const firstDirect = actorDirect(baseline, firstDirectPlan.plan);
          const first = await recomposeOwnedProposals(
            alice!,
            { revision: propertyToken(alice!), targetMarkdown: target('First'), ownership: [] },
            propertyContext,
            { target: options('first'), payloads: { presence: { proposalId: 'birth-first' } } }
          );
          if (!first.success) throw Error(first.error);
          await pair.settle();
          for (const editor of pair.editors) {
            const groups = readReviewGroups(editor.surface!.session.currentPackage());
            expect(groups.success && groups.groups.every((group) => !group.readOnly)).toBe(true);
          }
          pair.pause();
          expect(
            bob!.surface!.applyAutomationOps(() => [
              { op: 'insertText', paragraphId: id, offset: 7, text: ' Peer' },
            ]).rejected
          ).toBe(false);
          const second = await recomposeOwnedProposals(
            alice!,
            {
              revision: propertyToken(alice!),
              targetMarkdown: target('Second'),
              ownership: first.ownership,
              contentProofs: first.contentProofs,
            },
            propertyContext,
            { target: options('second'), payloads: { presence: { proposalId: 'birth-second' } } }
          );
          if (!second.success) throw Error(second.error);
          pair.reconnect(receiver);
          await pair.settle();
          const check = async (proposed: boolean, peer = true) => {
            for (const editor of pair.editors) {
              const original = await readCommittedView(
                editor.surface!.session,
                propertyToken(editor)
              );
              const working = await readWorkingView(editor.surface!.session, propertyToken(editor));
              if (!original.success || !working.success) throw Error('Missing review view');
              const prefix = peer ? 'Outside Peer' : 'Outside';
              expect(key(original.view.projection.text)).toBe(key(prefix));
              expect(key(working.view.projection.text)).toBe(
                key(proposed ? target('Second').replace('Outside', prefix) : prefix)
              );
            }
          };
          await check(true);
          const born = bob!
            .surface!.session.paragraphIds()
            .find(
              (paragraph) =>
                paragraphTextOf(bob!.surface!.session.part(), paragraph) === 'A rich link.'
            );
          if (!born) throw Error('Missing born editable paragraph');
          bob!.surface!.setSelection({
            anchor: { paragraphId: born, offset: 2 },
            head: { paragraphId: born, offset: 2 },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          );
          expect(remote?.head.nodeId).toBe(born);
          expect(remote?.head.offset).toBe(2);
          const a = caretAt(alice!.surface!.layout(), { paragraphId: born, offset: 2 });
          const b = caretAt(bob!.surface!.layout(), { paragraphId: born, offset: 2 });
          expect(a).not.toBeNull();
          expect(b).not.toBeNull();
          expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
          expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(
            0
          );
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(true, false);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(true);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            const read = await readWorkingView(editor.surface!.session, propertyToken(editor));
            if (!read.success) throw Error(read.error);
            expect(key(read.view.projection.text)).toBe(
              key(target('First').replace('Outside', 'Outside Peer'))
            );
          }
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(true);
          const bytes = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
          let exported: Uint8Array | undefined;
          expect(
            alice!.surface!.applyAutomationOps(
              () => [],
              { kind: 'body' },
              [
                (working) => {
                  exported = writeOoxmlPackage(structuredClone(working));
                  return working;
                },
              ],
              true,
              true,
              'after'
            ).rejected
          ).toBe(false);
          await pair.settle();
          pair.documents.forEach((doc, index) =>
            expect(Y.encodeStateAsUpdate(doc)).toEqual(bytes[index]!)
          );
          const pending = readOoxmlPackage(exported!);
          if (!pending.ok) throw Error(pending.reason);
          const reopenedGroups = readReviewGroups(pending.package);
          expect(reopenedGroups.success && reopenedGroups.groups.length).toBe(3);
          expect(
            reopenedGroups.success && reopenedGroups.groups.every((group) => !group.readOnly)
          ).toBe(true);
          const peer = applyTreeOp(baseline.parts.get(baseline.mainDocumentPart)!, {
            op: 'insertText',
            paragraphId: id,
            offset: 7,
            text: ' Peer',
          });
          if (!peer.ok) throw Error(peer.reason);
          const peerBase = withPart(baseline, peer.part);
          const committed = await readCommittedView(alice!.surface!.session, propertyToken(alice!));
          if (!committed.success) throw Error(committed.error);
          const expectedPlan = planTargetDocument(
            {
              pkg: peerBase,
              revision: propertyToken(alice!),
              projection: committed.view.projection,
            },
            target('Second').replace('Outside', 'Outside Peer'),
            options('second')
          );
          if (!expectedPlan.success) throw Error(expectedPlan.error);
          const direct = actorDirect(withPart(firstDirect, peer.part), expectedPlan.plan);
          const expected = word(decision === 'accept' ? direct : peerBase);
          const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            bob!,
            propertyToken(bob!),
            propertyContext
          );
          if (!decided.success) throw Error(decided.error);
          await pair.settle();
          for (const editor of pair.editors)
            expect(word(editor.surface!.session.currentPackage())).toBe(expected);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(true);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          // Pending DOCX and complete room snapshots must retain actionable ownership.
          for (const shared of [false, true]) {
            const pendingRoom = await createTwoEditorFixture({
              ...fixture,
              document: exported!,
              ...(shared ? { roomSnapshot: bytes[0]! } : {}),
            });
            try {
              await pendingRoom.settle();
              const groups = readReviewGroups(
                pendingRoom.editors[0]!.surface!.session.currentPackage()
              );
              expect(groups.success && groups.groups.length).toBe(3);
              expect(groups.success && groups.groups.every((group) => !group.readOnly)).toBe(true);
              const result = (
                decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups
              )(pendingRoom.editors[1]!, propertyToken(pendingRoom.editors[1]!), propertyContext);
              if (!result.success) throw Error(result.error);
              await pendingRoom.settle();
              for (const editor of pendingRoom.editors)
                expect(word(editor.surface!.session.currentPackage())).toBe(expected);
              pendingRoom.ownership.forEach((handle) => handle.assertHealthy());
            } finally {
              pendingRoom.destroy();
            }
          }
          const room = await createTwoEditorFixture({
            ...fixture,
            document: new Uint8Array(await alice!.save()),
            roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
          });
          try {
            await room.settle();
            for (const editor of room.editors)
              expect(word(editor.surface!.session.currentPackage())).toBe(expected);
            room.ownership.forEach((handle) => handle.assertHealthy());
          } finally {
            room.destroy();
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
          expect(paragraphTextOf(alice!.surface!.session.part(), id)).toBe('Outside Peer');
        } finally {
          pair.destroy();
        }
      }, 30000);
