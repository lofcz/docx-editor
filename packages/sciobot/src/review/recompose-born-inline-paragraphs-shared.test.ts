// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
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
import { planTargetDocument } from '../diff/plan-target';
import { previewDirectPlan } from '../ai/apply-direct';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { recomposeOwnedProposals } from './recompose';
import { readWorkingView } from './working-view';
import { readReviewGroups } from './groups';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { readPayloadStore, withInitializedPayloadStore } from '../blocks/payload-opc-codec';

const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
const word = (pkg: OoxmlPackage) => {
  const reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!reopened.ok) throw Error(reopened.reason);
  return native(reopened.package);
};
// Native allocation uses the same documented actor scope as the independent editor transaction.
function actorDirect(pkg: OoxmlPackage, plan: Parameters<typeof previewDirectPlan>[1]) {
  const store = new TreePackageStore(pkg, pkg.parts.get(pkg.mainDocumentPart)!);
  const result = store.transact(
    { kind: 'body' },
    (ctx) => {
      if (
        !ctx.applyPackage((working) => {
          const applied = previewDirectPlan(working, plan);
          if (!applied.success) throw Error(applied.error);
          return applied.pkg;
        })
      )
        throw Error('Independent direct transaction refused.');
    },
    { recordsHistory: false, actorId: 'alice' }
  );
  if (!result.ok) throw Error(result.reason);
  return store.currentPackage();
}
const atom = '<scio-mention key="user" value="Ada" />';
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const decision of ['accept', 'reject'] as const)
    test(`born paragraph inline atom ${clients}/${decision} retains offline edits, histories, exact decisions, and room reopening`, async () => {
      const pair = await createTwoEditorFixture({
        document: docxFixture('<w:p><w:r><w:t>Start</w:t></w:r></w:p>'),
        documentId: propertyContext.documentId,
        ownedAttributes: true,
        clientIds: clients,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors;
        const sourceId = alice!.surface!.session.paragraphIds()[0]!;
        const before = alice!.surface!.session.currentPackage();
        const peerOp = {
          op: 'insertText' as const,
          paragraphId: sourceId,
          offset: 5,
          text: ' Peer',
        };
        const peerPart = applyTreeOp(before.parts.get(before.mainDocumentPart)!, peerOp);
        if (!peerPart.ok) throw Error(peerPart.reason);
        const rejectedExpected = withPart(before, peerPart.part);
        const initial = await readWorkingView(alice!.surface!.session, propertyToken(alice!));
        if (!initial.success) throw Error(initial.error);
        const direct = planTargetDocument(
          { pkg: before, revision: propertyToken(alice!), projection: initial.view.projection },
          'Start\n\nBorn ' + atom + ' tail.',
          { allocateBlockId: () => 'born-atom' }
        );
        if (!direct.success) throw Error(direct.error);
        const expectedPlan = actorDirect(before, direct.plan);
        const expectedPeer = applyTreeOp(expectedPlan.parts.get(before.mainDocumentPart)!, peerOp);
        if (!expectedPeer.ok) throw Error(expectedPeer.reason);
        const acceptedExpected = withPart(expectedPlan, expectedPeer.part);
        pair.pause();
        expect(bob!.surface!.applyAutomationOps(() => [peerOp]).rejected).toBe(false);
        const staged = await recomposeOwnedProposals(
          alice!,
          {
            revision: propertyToken(alice!),
            targetMarkdown: 'Start\n\nBorn ' + atom + ' tail.',
            ownership: [],
          },
          propertyContext,
          {
            target: { allocateBlockId: () => 'born-atom' },
            payloads: { presence: { proposalId: 'born-presence' } },
          }
        );
        if (!staged.success) throw Error(staged.error);
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const check = async () => {
          for (const editor of pair.editors) {
            const view = await readWorkingView(editor.surface!.session, propertyToken(editor));
            expect(view.success && view.view.projection.text.trimEnd()).toBe(
              'Start Peer\n\nBorn [Ada](mention:user) tail\\.'
            );
            const groups = readReviewGroups(editor.surface!.session.currentPackage());
            expect(groups.success && groups.groups.length).toBe(1);
            expect(groups.success && groups.groups[0]!.readOnly).toBe(false);
          }
        };
        await check();
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(word(editor.surface!.session.currentPackage())).toBe(word(rejectedExpected));
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check();
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(paragraphTextOf(editor.surface!.session.part(), sourceId)).toBe('Start');
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check();
        bob!.surface!.setSelection({
          anchor: { paragraphId: sourceId, offset: 10 },
          head: { paragraphId: sourceId, offset: 10 },
        });
        await pair.settle();
        const geometry = pair.editors.map((editor) =>
          caretAt(editor.surface!.layout(), { paragraphId: sourceId, offset: 10 })
        );
        expect(geometry[0]).not.toBeNull();
        expect(geometry[0]).toEqual(geometry[1]);
        const pending = pair.editors.map((editor) =>
          native(editor.surface!.session.currentPackage())
        );
        const resolved = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
          alice!,
          propertyToken(alice!),
          propertyContext
        );
        if (!resolved.success) throw Error(resolved.error);
        await pair.settle();
        const expected = decision === 'accept' ? acceptedExpected : rejectedExpected;
        for (const editor of pair.editors) {
          expect(word(editor.surface!.session.currentPackage())).toBe(word(expected));
          const payloads = readPayloadStore(editor.surface!.session.currentPackage());
          expect(payloads.status === 'ready' && [...payloads.entries.keys()]).toEqual(
            decision === 'accept' ? ['born-atom'] : []
          );
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        pair.editors.forEach((editor, index) =>
          expect(native(editor.surface!.session.currentPackage())).toBe(pending[index]!)
        );
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
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
    });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`peer text inside a born atom paragraph ${clients} survives birth undo and native rejection`, async () => {
    const source = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Start</w:t></w:r></w:p>'));
    if (!source.ok) throw Error(source.reason);
    // Shared application hosts initialize payload storage before room bootstrap.
    const prepared = withInitializedPayloadStore(source.package);
    if (!prepared.ok) throw Error(prepared.reason);
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(prepared.pkg),
      documentId: propertyContext.documentId,
      ownedAttributes: true,
      clientIds: clients,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors;
      const staged = await recomposeOwnedProposals(
        alice!,
        {
          revision: propertyToken(alice!),
          targetMarkdown: 'Start\n\nBorn ' + atom + ' tail.',
          ownership: [],
        },
        propertyContext,
        {
          target: { allocateBlockId: () => 'born-atom' },
          payloads: { presence: { proposalId: 'born-presence' } },
        }
      );
      if (!staged.success) throw Error(staged.error);
      await pair.settle();
      const bornId = bob!
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob!.surface!.session.part(), id)!.startsWith('Born '))!;
      expect(
        bob!.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: bornId, offset: 5, text: ' Peer ' },
        ]).rejected
      ).toBe(false);
      await pair.settle();
      const pending = pair.editors.map((editor) =>
        native(editor.surface!.session.currentPackage())
      );
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      const peerOnly = () => {
        for (const editor of pair.editors) {
          const session = editor.surface!.session;
          expect(session.paragraphIds().map((id) => paragraphTextOf(session.part(), id))).toEqual([
            'Start',
            ' Peer ',
          ]);
          const payloads = readPayloadStore(session.currentPackage());
          expect(payloads.status === 'ready' && [...payloads.entries.keys()]).toEqual([]);
        }
      };
      peerOnly();
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      pair.editors.forEach((editor, index) =>
        expect(native(editor.surface!.session.currentPackage())).toBe(pending[index]!)
      );
      const rejected = rejectAllReviewGroups(alice!, propertyToken(alice!), propertyContext);
      if (!rejected.success) throw Error(rejected.error);
      await pair.settle();
      peerOnly();
      const decided = pair.editors.map((editor) =>
        native(editor.surface!.session.currentPackage())
      );
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      pair.editors.forEach((editor, index) =>
        expect(native(editor.surface!.session.currentPackage())).toBe(pending[index]!)
      );
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      peerOnly();
      pair.editors.forEach((editor, index) =>
        expect(native(editor.surface!.session.currentPackage())).toBe(decided[index]!)
      );
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      pair.ownership.forEach((handle) => handle.assertHealthy());
    } finally {
      pair.destroy();
    }
  });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`offline peer typing survives concurrent rejection of a born atom paragraph ${clients}`, async () => {
    const source = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Start</w:t></w:r></w:p>'));
    if (!source.ok) throw Error(source.reason);
    const prepared = withInitializedPayloadStore(source.package);
    if (!prepared.ok) throw Error(prepared.reason);
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(prepared.pkg),
      documentId: propertyContext.documentId,
      ownedAttributes: true,
      clientIds: clients,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors;
      const staged = await recomposeOwnedProposals(
        alice!,
        {
          revision: propertyToken(alice!),
          targetMarkdown: 'Start\n\nBorn ' + atom + ' tail.',
          ownership: [],
        },
        propertyContext,
        {
          target: { allocateBlockId: () => 'born-atom' },
          payloads: { presence: { proposalId: 'born-presence' } },
        }
      );
      if (!staged.success) throw Error(staged.error);
      await pair.settle();
      const born = bob!
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob!.surface!.session.part(), id)!.startsWith('Born '))!;
      pair.pause();
      expect(
        bob!.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: born, offset: 5, text: ' Peer ' },
        ]).rejected
      ).toBe(false);
      const pendingPeer = native(bob!.surface!.session.currentPackage());
      const rejected = rejectAllReviewGroups(alice!, propertyToken(alice!), propertyContext);
      if (!rejected.success) throw Error(rejected.error);
      pair.reconnect(clients[0] === 41 ? 0 : 1);
      await pair.settle();
      const check = (editors: typeof pair.editors) => {
        for (const editor of editors) {
          const session = editor.surface!.session;
          expect(session.paragraphIds().map((id) => paragraphTextOf(session.part(), id))).toEqual([
            'Start',
            ' Peer ',
          ]);
          const payloads = readPayloadStore(session.currentPackage());
          expect(payloads.status === 'ready' && [...payloads.entries.keys()]).toEqual([]);
        }
      };
      check(pair.editors);
      const decided = pair.editors.map((editor) =>
        native(editor.surface!.session.currentPackage())
      );
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(native(editor.surface!.session.currentPackage())).toBe(pendingPeer);
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      check(pair.editors);
      pair.editors.forEach((editor, index) =>
        expect(native(editor.surface!.session.currentPackage())).toBe(decided[index]!)
      );
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
        check(reopened.editors);
      } finally {
        reopened.destroy();
      }
      pair.ownership.forEach((handle) => handle.assertHealthy());
    } finally {
      pair.destroy();
    }
  });
