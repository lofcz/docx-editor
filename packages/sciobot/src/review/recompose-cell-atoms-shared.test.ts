import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  applyTreeOp,
  TreePackageStore,
  withPart,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import {
  context,
  png,
  decodePort,
  media,
  native,
  bindControls,
  cases,
} from '../../test/support/tracked-cell-atoms';
import { expectImageRelationships } from '../../test/support/image-relationships';
import { buildBlockIndex } from '../blocks/block-index';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from './working-view';
import { readReviewGroups } from './groups';
import { planTargetDocument } from '../diff/plan-target';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { previewDirectPlan } from '../ai/apply-direct';
import { recomposeOwnedProposals } from './recompose';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
// History retains a private caret marker in the live store. Word export removes that marker.
// Compare the complete exported Word tree with the independent native result, and require
// exact live canonical convergence separately. Do not omit any authored Word element.
const wordNative = (pkg: Parameters<typeof native>[0]) => {
  const opened = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!opened.ok) throw Error(opened.reason);
  return native(opened.package);
};
for (const item of cases)
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`shared tracked ${item.name} atoms ${clients}/${decision} preserve peer text, both histories, exact decisions, and cold reopening`, async () => {
        const pair = await createTwoEditorFixture({
          document: docxFixture(item.body),
          documentId: context.documentId,
          clientIds: clients,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE], imageDecodePort: decodePort },
        });
        try {
          const [alice, bob] = pair.editors;
          const token = (editor: typeof alice) =>
            readEditorRevision(editor!, context.documentId, context.epoch)!;
          const before = alice!.surface!.session.currentPackage();
          const peerId = alice!
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(alice!.surface!.session.part(), id) === 'Peer')!;
          const op = { op: 'insertText' as const, paragraphId: peerId, offset: 4, text: ' note' };
          const peerPart = applyTreeOp(before.parts.get(before.mainDocumentPart)!, op);
          if (!peerPart.ok) throw Error(peerPart.reason);
          const rejected = withPart(before, peerPart.part);
          const view = await readWorkingView(alice!.surface!.session, token(alice));
          if (!view.success) throw Error(view.error);
          const plan = planTargetDocument(
            { pkg: before, revision: token(alice), projection: view.view.projection },
            item.target
          );
          if (!plan.success) throw Error(plan.error);
          const resources = await prepareInlineImageResources(
            plan.plan,
            async () => ({ bytes: png, mime: 'image/png' }),
            decodePort
          );
          if (!resources.success) throw Error(resources.error);
          const directStore = new TreePackageStore(
            rejected,
            rejected.parts.get(rejected.mainDocumentPart)!
          );
          const previewed = directStore.transact(
            { kind: 'body' },
            (ctx) => {
              if (
                !ctx.applyPackage((pkg) => {
                  const result = previewDirectPlan(pkg, plan.plan, resources.images);
                  if (!result.success) throw Error(result.error);
                  return result.pkg;
                })
              )
                throw Error('The independent direct preview refused.');
            },
            {
              recordsHistory: false,
              actorId: alice!.surface!.collaborationSession()!.identity.actorId,
            }
          );
          if (!previewed.ok) throw Error(previewed.reason);
          const direct = { pkg: directStore.currentPackage() };
          pair.pause();
          let calls = 0;
          const result = await recomposeOwnedProposals(
            alice!,
            { revision: token(alice), targetMarkdown: item.target, ownership: [] },
            context,
            {
              payloads: { presence: { proposalId: 'shared-cell-presence' } },
              images: {
                decodePort,
                resolve: async (reference) => {
                  expect(reference).toBe('asset:cat');
                  calls++;
                  return { bytes: png, mime: 'image/png' };
                },
              },
            }
          );
          if (!result.success) throw Error(result.error);
          expect(calls).toBe(1);
          expect(
            bob!.surface!.applyAutomationOps(() => [op], { kind: 'body' }, [], false, true)
          ).toMatchObject({ committed: true, rejected: false });
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const pending = alice!.surface!.session.currentPackage(),
            pendingNative = wordNative(pending);
          const accepted = bindControls(direct.pkg, pending, before);
          const check = (expected: string, count: number) => {
            for (const editor of pair.editors) {
              const pkg = editor.surface!.session.currentPackage();
              expect(wordNative(pkg)).toBe(expected);
              expect(native(pkg)).toBe(native(pair.editors[0]!.surface!.session.currentPackage()));
              expect(buildBlockIndex(pkg).byId.size).toBe(count);
              expect(media(pkg).map(([, bytes]) => bytes)).toEqual(count ? [png] : []);
              expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
            }
          };
          check(pendingNative, 4);
          const groups = readReviewGroups(pending);
          if (!groups.success) throw Error(groups.error);
          expect(groups.success && groups.groups.length).toBe(1);
          expect(groups.success && groups.groups[0]!.readOnly).toBe(false);
          expect(groups.success && groups.groups[0]!.properties.length).toBe(4);
          if (decision === 'accept') {
            const pendingCold = await createTwoEditorFixture({
              document: new Uint8Array(await alice!.save()),
              documentId: context.documentId,
              roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
              ownedAttributes: true,
              editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE], imageDecodePort: decodePort },
            });
            try {
              for (const editor of pendingCold.editors) {
                const pkg = editor.surface!.session.currentPackage();
                expect(wordNative(pkg)).toBe(pendingNative);
                expect(buildBlockIndex(pkg).byId.size).toBe(4);
                expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
                const readable = readReviewGroups(pkg);
                expect(readable.success && readable.groups.length).toBe(1);
                expect(readable.success && readable.groups[0]!.readOnly).toBe(false);
              }
              pendingCold.ownership.forEach((handle) => handle.assertHealthy());
            } finally {
              pendingCold.destroy();
            }
          }
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(wordNative(rejected), 0);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(pendingNative, 4);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(paragraphTextOf(editor.surface!.session.part(), peerId)).toBe('Peer');
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(pendingNative, 4);
          const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            bob!,
            token(bob),
            context
          );
          if (!decided.success) throw Error(decided.error);
          await pair.settle();
          const resolved = decision === 'accept' ? wordNative(accepted) : wordNative(rejected),
            count = decision === 'accept' ? 4 : 0;
          check(resolved, count);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(pendingNative, 4);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(resolved, count);
          const cold = await createTwoEditorFixture({
            document: new Uint8Array(await bob!.save()),
            documentId: context.documentId,
            roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
            ownedAttributes: true,
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE], imageDecodePort: decodePort },
          });
          try {
            for (const editor of cold.editors) {
              const pkg = editor.surface!.session.currentPackage();
              expect(wordNative(pkg)).toBe(resolved);
              expect(media(pkg).map(([, bytes]) => bytes)).toEqual(count ? [png] : []);
              expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
            }
            cold.ownership.forEach((h) => h.assertHealthy());
          } finally {
            cold.destroy();
          }
          pair.ownership.forEach((h) => h.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 30000);
