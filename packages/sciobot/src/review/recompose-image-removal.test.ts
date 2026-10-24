import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  imageRemovalFixture,
  imageRemovalContext as context,
  imageRemovalDecodePort,
  imageRemovalMedia as media,
} from '../../test/support/image-removal';
import { expectImageRelationships } from '../../test/support/image-relationships';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { planTargetDocument } from '../diff/plan-target';
import { previewDirectPlan } from '../ai/apply-direct';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { recomposeOwnedProposals } from './recompose';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
const payloads = (pkg: OoxmlPackage) => {
  const read = readPayloadStore(pkg);
  if (read.status !== 'ready') throw Error('Missing image payloads.');
  return [...read.entries].map(([id, entry]) => [id, entry.decoded]);
};
for (const linked of [false, true])
  for (const shared of [false, true])
    for (const decision of ['accept', 'reject'] as const)
      test(`tracked image removal ${linked}/${shared}/${decision} preserves native content, shared media, payloads, history, and reopening`, async () => {
        const fixture = await imageRemovalFixture(linked, shared);
        const editor = createDocxEditor({
          container: document.createElement('div'),
          document: writeOoxmlPackage(fixture),
          imageDecodePort: imageRemovalDecodePort,
          modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
        });
        try {
          const surface = editor.surface!;
          const revision = () => readEditorRevision(editor, context.documentId, context.epoch)!;
          const before = surface.session.currentPackage();
          expect(media(before)).toHaveLength(shared ? 1 : 2);
          const index = buildBlockIndex(before);
          const cat = [...index.byId.values()].find(
            (block) => block.payload.attributes.reference === 'asset:cat'
          )!;
          const dog = [...index.byId.values()].find(
            (block) => block.payload.attributes.reference === 'asset:dog'
          )!;
          const originalDog = canonicalOoxmlFingerprint(
            findNode(before.parts.get(dog.partName)!, dog.controlId)!
          );
          const initial = await readWorkingView(surface.session, revision());
          if (!initial.success) throw Error(initial.error);
          const image = '![Cat](asset:cat "Photo")';
          const target = initial.view.projection.text.replace(
            linked ? `[${image}](https://keep.test "Keep")` : image,
            ''
          );
          expect(target).not.toBe(initial.view.projection.text);
          const planned = planTargetDocument(
            { pkg: before, revision: revision(), projection: initial.view.projection },
            target
          );
          if (!planned.success) throw Error(planned.error);
          expect(planned.plan.operations.map((edit) => edit.kind)).toEqual(
            linked ? ['inline-remove', 'unlink'] : ['inline-remove']
          );
          const direct = previewDirectPlan(before, planned.plan);
          if (!direct.success) throw Error(direct.error);
          const staged = await recomposeOwnedProposals(
            editor,
            { revision: revision(), targetMarkdown: target, ownership: [] },
            context,
            {
              payloads: { presence: { proposalId: 'remove-image' } },
              images: {
                decodePort: imageRemovalDecodePort,
                resolve: async () => {
                  throw Error('Image removal must not fetch resources.');
                },
              },
            }
          );
          expect(staged).toMatchObject({ success: true, committed: true });
          const pending = surface.session.currentPackage();
          expect(media(pending)).toEqual(media(before));
          expect(payloads(pending)).toEqual(payloads(before));
          expectImageRelationships(pending, buildBlockIndex(pending).byId.values());
          const committed = await readCommittedView(surface.session, revision());
          expect(committed.success && committed.view.projection.text).toBe(
            initial.view.projection.text
          );
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(native(before));
          expect(media(surface.session.currentPackage())).toEqual(media(before));
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(native(pending));
          const decide = decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups;
          expect(decide(editor, revision(), context)).toMatchObject({ success: true });
          const expected = decision === 'accept' ? direct.pkg : before;
          const resolved = surface.session.currentPackage();
          expect(native(resolved)).toBe(native(expected));
          expect(media(resolved)).toEqual(media(expected));
          expect(payloads(resolved)).toEqual(payloads(expected));
          expect(buildBlockIndex(resolved).byId.has(cat.id)).toBe(decision === 'reject');
          expect(
            canonicalOoxmlFingerprint(findNode(resolved.parts.get(dog.partName)!, dog.controlId)!)
          ).toBe(originalDog);
          expectImageRelationships(resolved, buildBlockIndex(resolved).byId.values());
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(native(pending));
          expect(media(surface.session.currentPackage())).toEqual(media(before));
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(native(expected));
          const saved = readOoxmlPackage(writeOoxmlPackage(surface.session.currentPackage()));
          if (!saved.ok) throw Error(saved.reason);
          expect(native(saved.package)).toBe(native(expected));
          expect(media(saved.package)).toEqual(media(expected));
          expect(payloads(saved.package)).toEqual(payloads(expected));
          expectImageRelationships(saved.package, buildBlockIndex(saved.package).byId.values());
        } finally {
          editor.destroy();
        }
      });
