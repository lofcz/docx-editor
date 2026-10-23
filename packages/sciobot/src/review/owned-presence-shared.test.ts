// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  withPart,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  presenceReviewFixture,
  installPresenceOwnership,
} from '../../test/support/presence-ownership';
import {
  propertyContext,
  propertyToken,
  propertyEditor,
} from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { planParagraphText } from '../diff/paragraph-text';
import { previewDirectPlan } from '../ai/apply-direct';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { recomposeOwnedProposals } from './recompose';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { readReviewGroups } from './groups';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { verifyOwnedProposalContentProofs } from './owned-proposal-content';

const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
const word = (pkg: OoxmlPackage) => {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return native(read.package);
};
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const remove of [false, true])
    for (const decision of ['accept', 'reject'] as const)
      test(`owned presence recomposition ${clients}/${remove}/${decision} retains offline edits, histories, geometry, exact decisions, and room reopening`, async () => {
        const pair = await createTwoEditorFixture({
          document: writeOoxmlPackage(presenceReviewFixture()),
          documentId: propertyContext.documentId,
          ownedAttributes: true,
          clientIds: clients,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors;
          const owned = installPresenceOwnership(alice!, remove);
          await pair.settle();
          const outside = bob!
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'Outside')!;
          const original = await readCommittedView(alice!.surface!.session, propertyToken(alice!));
          if (!original.success) throw Error(original.error);
          const markdown = original.view.projection.text.replace('Outside', 'Rewritten');
          pair.pause();
          expect(
            bob!.surface!.applyAutomationOps(() => [
              { op: 'insertText', paragraphId: outside, offset: 7, text: ' Peer' },
            ]).rejected
          ).toBe(false);
          const replaced = await recomposeOwnedProposals(
            alice!,
            {
              revision: propertyToken(alice!),
              targetMarkdown: markdown,
              ownership: owned.ownership,
              contentProofs: owned.contentProofs,
            },
            propertyContext
          );
          if (!replaced.success) throw Error(replaced.error);
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const proposed = markdown.replace('Rewritten', 'Rewritten Peer');
          const check = async () => {
            for (const editor of pair.editors) {
              const working = await readWorkingView(editor.surface!.session, propertyToken(editor));
              if (!working.success) throw Error(working.error);
              expect(working.view.projection.text.trimEnd()).toBe(proposed.trimEnd());
              const rows = readReviewGroups(editor.surface!.session.currentPackage());
              if (!rows.success) throw Error(rows.error);
              expect(rows.groups.length).toBeGreaterThan(0);
              expect(
                rows.groups.every((group) => group.properties.length === 0 && !group.readOnly)
              ).toBe(true);
            }
          };
          await check();
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(
              verifyOwnedProposalContentProofs(
                editor.surface!.session.currentPackage(),
                propertyToken(editor),
                owned.ownership,
                owned.contentProofs
              )
            ).toBeNull();
            const rows = readReviewGroups(editor.surface!.session.currentPackage());
            expect(rows.success && rows.groups.map((group) => group.id)).toEqual(
              owned.ownership.map((group) => group.groupId)
            );
            expect(paragraphTextOf(editor.surface!.session.part(), outside)).toBe('Outside Peer');
          }
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check();
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            const read = await readWorkingView(editor.surface!.session, propertyToken(editor));
            expect(read.success && read.view.projection.text.trimEnd()).toBe(markdown.trimEnd());
          }
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check();
          bob!.surface!.setSelection({
            anchor: { paragraphId: outside, offset: 14 },
            head: { paragraphId: outside, offset: 14 },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          );
          expect(remote?.head.nodeId).toBe(outside);
          expect(remote?.head.offset).toBe(14);
          const geometry = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId: outside, offset: 14 })
          );
          expect(geometry[0]).not.toBeNull();
          expect(geometry[1]).not.toBeNull();
          expect({ x: geometry[0]!.x, y: geometry[0]!.y }).toEqual({
            x: geometry[1]!.x,
            y: geometry[1]!.y,
          });
          const base = owned.baseline;
          const peer = applyTreeOp(base.parts.get(base.mainDocumentPart)!, {
            op: 'insertText',
            paragraphId: outside,
            offset: 7,
            text: ' Peer',
          });
          if (!peer.ok) throw Error(peer.reason);
          const root = buildBlockIndex(base).byId.get('outer')!.controlId;
          const expectedOriginal = withPart(
            base,
            remove
              ? peer.part
              : {
                  ...peer.part,
                  root: {
                    ...peer.part.root,
                    children: peer.part.root.children.map((node) =>
                      node.kind === 'body'
                        ? ({
                            ...node,
                            children: node.children.filter((child) => child.id !== root),
                          } as OoxmlElement)
                        : node
                    ),
                  } as OoxmlElement,
                }
          );
          const directPlan = planParagraphText(
            {
              revision: propertyToken(alice!),
              partName: base.mainDocumentPart,
              paragraphId: outside,
              text: 'Outside Peer',
            },
            'Rewritten Peer'
          );
          if (!directPlan.success) throw Error(directPlan.error);
          const direct = previewDirectPlan(expectedOriginal, directPlan.plan);
          if (!direct.success) throw Error(direct.error);
          const expected = decision === 'accept' ? direct.pkg : expectedOriginal;
          const roomBytes = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
          let pendingWord: Uint8Array | undefined;
          const serialized = alice!.surface!.applyAutomationOps(
            () => [],
            { kind: 'body' },
            [
              (working) => {
                pendingWord = writeOoxmlPackage(working);
                return working;
              },
            ],
            true,
            true,
            'after'
          );
          expect(serialized.rejected).toBe(false);
          await pair.settle();
          pair.documents.forEach((doc, index) =>
            expect(Y.encodeStateAsUpdate(doc)).toEqual(roomBytes[index]!)
          );
          const wordRead = readOoxmlPackage(pendingWord!);
          if (!wordRead.ok) throw Error(wordRead.reason);
          expect(native(wordRead.package)).not.toContain(
            'urn:docx-editor:tracked-source-bindings:1'
          );
          const local = propertyEditor(wordRead.package);
          try {
            const wordDecision = (
              decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups
            )(local, propertyToken(local), propertyContext);
            if (!wordDecision.success) throw Error(wordDecision.error);
            expect(word(local.surface!.session.currentPackage())).toBe(word(expected));
            expect(local.exec({ type: 'undo' }).ok).toBe(true);
            expect(local.exec({ type: 'redo' }).ok).toBe(true);
            expect(word(local.surface!.session.currentPackage())).toBe(word(expected));
          } finally {
            local.destroy();
          }
          const pendingRoom = await createTwoEditorFixture({
            document: pendingWord!,
            documentId: propertyContext.documentId,
            ownedAttributes: true,
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
            roomSnapshot: roomBytes[0]!,
          });
          try {
            await pendingRoom.settle();
            const roomDecision = (
              decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups
            )(pendingRoom.editors[0]!, propertyToken(pendingRoom.editors[0]!), propertyContext);
            if (!roomDecision.success) throw Error(roomDecision.error);
            await pendingRoom.settle();
            for (const editor of pendingRoom.editors)
              expect(word(editor.surface!.session.currentPackage())).toBe(word(expected));
            pendingRoom.ownership.forEach((handle) => handle.assertHealthy());
          } finally {
            pendingRoom.destroy();
          }
          const pending = pair.editors.map((editor) =>
            native(editor.surface!.session.currentPackage())
          );
          const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            alice!,
            propertyToken(alice!),
            propertyContext
          );
          if (!decided.success) throw Error(decided.error);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(word(editor.surface!.session.currentPackage())).toBe(word(expected));
            const payloads = readPayloadStore(editor.surface!.session.currentPackage());
            expect(payloads.status === 'ready' && [...payloads.entries.keys()].sort()).toEqual(
              remove ? ['inner', 'outer'] : []
            );
            const rows = readReviewGroups(editor.surface!.session.currentPackage());
            expect(rows.success && rows.groups.length).toBe(0);
          }
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          pair.editors.forEach((editor, index) =>
            expect(native(editor.surface!.session.currentPackage())).toBe(pending[index]!)
          );
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          const saved = await pair.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
          const reopened = await createTwoEditorFixture({
            document: new Uint8Array(await bob!.save()),
            documentId: propertyContext.documentId,
            ownedAttributes: true,
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
            roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
          });
          try {
            await reopened.settle();
            for (const editor of reopened.editors)
              expect(word(editor.surface!.session.currentPackage())).toBe(word(expected));
            reopened.ownership.forEach((handle) => handle.assertHealthy());
          } finally {
            reopened.destroy();
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 30000);

test('rich paragraph source bindings pass shared field limits and restore every original run after an offline suffix', async () => {
  const { docxFixture } = await import('../../test/support/document');
  const { replaceOwnedProposals } = await import('./replace-owned-proposals');
  const pair = await createTwoEditorFixture({
    document: docxFixture('<w:p>' + '<w:r><w:t>X</w:t></w:r>'.repeat(80) + '</w:p>'),
    documentId: propertyContext.documentId,
    ownedAttributes: true,
    clientIds: [41, 73],
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
  });
  try {
    const [alice, bob] = pair.editors;
    const baseline = alice!.surface!.session.currentPackage();
    const paragraphId = alice!.surface!.session.paragraphIds()[0]!;
    pair.pause();
    expect(
      bob!.surface!.applyAutomationOps(() => [
        { op: 'insertText', paragraphId, offset: 80, text: ' Peer' },
      ]).rejected
    ).toBe(false);
    const plan = planParagraphText(
      {
        revision: propertyToken(alice!),
        partName: baseline.mainDocumentPart,
        paragraphId,
        text: 'X'.repeat(80),
      },
      'Y'
    );
    if (!plan.success) throw Error(plan.error);
    const proposed = replaceOwnedProposals(alice!, [], plan.plan, propertyContext);
    if (!proposed.success) throw Error(proposed.error);
    pair.reconnect(1);
    await pair.settle();
    const boundPart = alice!.surface!.session.part();
    const { findNode } = await import('@docx-editor.dev/core/store');
    const paragraph = findNode(boundPart, paragraphId)!;
    if (paragraph.kind !== 'paragraph') throw Error('Missing rich paragraph');
    const fields = paragraph.attributes.filter(
      (attribute) => attribute.namespaceUri === 'urn:docx-editor:tracked-source-bindings:1'
    );
    expect(fields.length).toBeGreaterThan(1);
    expect(fields.every((attribute) => attribute.value.length <= 2048)).toBe(true);
    const expected = applyTreeOp(baseline.parts.get(baseline.mainDocumentPart)!, {
      op: 'insertText',
      paragraphId,
      offset: 80,
      text: ' Peer',
    });
    if (!expected.ok) throw Error(expected.reason);
    const pending = native(alice!.surface!.session.currentPackage());
    const rejected = rejectAllReviewGroups(alice!, propertyToken(alice!), propertyContext);
    if (!rejected.success) throw Error(rejected.error);
    await pair.settle();
    for (const editor of pair.editors)
      expect(word(editor.surface!.session.currentPackage())).toBe(
        word(withPart(baseline, expected.part))
      );
    expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors)
      expect(native(editor.surface!.session.currentPackage())).toBe(pending);
    expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    const saved = await pair.savedDigests();
    expect(saved[0]).toEqual(saved[1]);
    pair.ownership.forEach((handle) => handle.assertHealthy());
  } finally {
    pair.destroy();
  }
}, 30000);
