import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  withPart,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  imageRemovalFixture,
  imageRemovalContext as context,
  imageRemovalDecodePort,
  imageRemovalMedia as media,
} from '../../test/support/image-removal';
import { expectImageRelationships } from '../../test/support/image-relationships';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { readEditorRevision } from '../projection/revision-cache';
import { planTargetDocument } from '../diff/plan-target';
import { previewDirectPlan } from '../ai/apply-direct';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readWorkingView } from './working-view';
import { recomposeOwnedProposals } from './recompose';
import { readReviewGroups } from './groups';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
const word = (pkg: OoxmlPackage) => {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return native(read.package);
};
for (const shared of [false, true])
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`shared image removal ${shared}/${clients}/${decision} preserves offline peers, native decisions, media, histories, and cold reopening`, async () => {
        const original = await imageRemovalFixture(true, shared);
        const pair = await createTwoEditorFixture({
          document: writeOoxmlPackage(original),
          documentId: context.documentId,
          ownedAttributes: true,
          clientIds: clients,
          editor: {
            modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
            imageDecodePort: imageRemovalDecodePort,
          },
        });
        try {
          const [alice, bob] = pair.editors;
          const before = alice!.surface!.session.currentPackage();
          const paragraphId = alice!.surface!.session.paragraphIds()[0]!;
          const token = (editor: (typeof pair.editors)[number]) =>
            readEditorRevision(editor, context.documentId, context.epoch)!;
          const index = buildBlockIndex(before);
          const cat = [...index.byId.values()].find(
            (block) => block.payload.attributes.reference === 'asset:cat'
          )!;
          const dog = [...index.byId.values()].find(
            (block) => block.payload.attributes.reference === 'asset:dog'
          )!;
          const dogFingerprint = canonicalOoxmlFingerprint(
            findNode(before.parts.get(dog.partName)!, dog.controlId)!
          );
          const initial = await readWorkingView(alice!.surface!.session, token(alice!));
          if (!initial.success) throw Error(initial.error);
          const target = initial.view.projection.text.replace(
            '[![Cat](asset:cat "Photo")](https://keep.test "Keep")',
            ''
          );
          expect(target).not.toBe(initial.view.projection.text);
          const planned = planTargetDocument(
            { pkg: before, revision: token(alice!), projection: initial.view.projection },
            target
          );
          if (!planned.success) throw Error(planned.error);
          const direct = previewDirectPlan(before, planned.plan);
          if (!direct.success) throw Error(direct.error);
          const peerOp = { op: 'insertText', paragraphId, offset: 2, text: ' peer' } as const;
          const withPeer = (pkg: OoxmlPackage) => {
            const edited = applyTreeOp(pkg.parts.get(pkg.mainDocumentPart)!, peerOp);
            if (!edited.ok) throw Error(edited.reason);
            return withPart(pkg, edited.part);
          };
          const rejected = withPeer(before),
            accepted = withPeer(direct.pkg);
          pair.pause();
          expect(bob!.surface!.applyAutomationOps(() => [peerOp]).rejected).toBe(false);
          const staged = await recomposeOwnedProposals(
            alice!,
            { revision: token(alice!), targetMarkdown: target, ownership: [] },
            context,
            { payloads: { presence: { proposalId: 'shared-remove-image' } } }
          );
          if (!staged.success) throw Error(staged.error);
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const pending = pair.editors.map((editor) =>
            native(editor.surface!.session.currentPackage())
          );
          expect(pending[0]).toBe(pending[1]);
          const checkPending = () => {
            for (const editor of pair.editors) {
              const pkg = editor.surface!.session.currentPackage();
              expect(media(pkg)).toEqual(media(before));
              expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
              const groups = readReviewGroups(pkg);
              expect(groups.success && groups.groups.every((group) => !group.readOnly)).toBe(true);
              expect(
                canonicalOoxmlFingerprint(findNode(pkg.parts.get(dog.partName)!, dog.controlId)!)
              ).toBe(dogFingerprint);
            }
          };
          checkPending();
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            const pkg = editor.surface!.session.currentPackage();
            expect(word(pkg)).toBe(word(rejected));
            expect(media(pkg)).toEqual(media(before));
          }
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          pair.editors.forEach((editor, i) =>
            expect(native(editor.surface!.session.currentPackage())).toBe(pending[i]!)
          );
          checkPending();
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(paragraphTextOf(bob!.surface!.session.part(), paragraphId)).not.toContain(' peer');
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          pair.editors.forEach((editor, i) =>
            expect(native(editor.surface!.session.currentPackage())).toBe(pending[i]!)
          );
          const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            alice!,
            token(alice!),
            context
          );
          if (!decided.success) throw Error(decided.error);
          await pair.settle();
          const expected = decision === 'accept' ? accepted : rejected;
          const checkResolved = () => {
            for (const editor of pair.editors) {
              const pkg = editor.surface!.session.currentPackage();
              expect(word(pkg)).toBe(word(expected));
              expect(media(pkg)).toEqual(media(expected));
              const read = readPayloadStore(pkg);
              expect(read.status === 'ready' && [...read.entries.keys()].sort()).toEqual(
                [dog.id, ...(decision === 'reject' ? [cat.id] : [])].sort()
              );
              expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
              expect(
                canonicalOoxmlFingerprint(findNode(pkg.parts.get(dog.partName)!, dog.controlId)!)
              ).toBe(dogFingerprint);
            }
          };
          checkResolved();
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          pair.editors.forEach((editor, i) =>
            expect(native(editor.surface!.session.currentPackage())).toBe(pending[i]!)
          );
          checkPending();
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          checkResolved();
          const geometry = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId, offset: 1 })
          );
          expect(geometry[0]).not.toBeNull();
          expect(geometry[0]).toEqual(geometry[1]);
          const saved = new Uint8Array(await bob!.save());
          for (const room of [false, true]) {
            const reopened = await createTwoEditorFixture({
              document: saved,
              documentId: context.documentId,
              ownedAttributes: true,
              editor: {
                modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
                imageDecodePort: imageRemovalDecodePort,
              },
              ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
            });
            try {
              for (const editor of reopened.editors) {
                const pkg = editor.surface!.session.currentPackage();
                expect(word(pkg)).toBe(word(expected));
                expect(media(pkg)).toEqual(media(expected));
                expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
              }
              reopened.ownership.forEach((handle) => handle.assertHealthy());
            } finally {
              reopened.destroy();
            }
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      });
