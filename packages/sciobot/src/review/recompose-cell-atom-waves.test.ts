import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import {
  context,
  png,
  decodePort,
  media,
  native,
  bindControls,
  cases,
  preparedCellDocument,
} from '../../test/support/tracked-cell-atoms';
import { expectImageRelationships } from '../../test/support/image-relationships';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { planTargetDocument } from '../diff/plan-target';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { previewDirectPlan } from '../ai/apply-direct';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { recomposeOwnedProposals } from './recompose';
import { readReviewGroups } from './groups';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
for (const item of cases)
  for (const decision of ['accept', 'reject'] as const)
    test(`repeated tracked ${item.name} atom waves retain exact native decisions and both creation histories: ${decision}`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: preparedCellDocument(item.body),
        imageDecodePort: decodePort,
        modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      });
      try {
        const surface = editor.surface!,
          token = () => readEditorRevision(editor, context.documentId, context.epoch)!;
        const before = surface.session.currentPackage(),
          original = native(before);
        const initialToken = token();
        const initial = await readWorkingView(surface.session, initialToken);
        if (!initial.success) throw Error(initial.error);
        const first = await recomposeOwnedProposals(
          editor,
          { revision: token(), targetMarkdown: item.target, ownership: [] },
          context,
          {
            payloads: { presence: { proposalId: 'first-cell-wave' } },
            images: { decodePort, resolve: async () => ({ bytes: png, mime: 'image/png' }) },
          }
        );
        if (!first.success) throw Error(first.error);
        const firstPending = native(surface.session.currentPackage());
        const target = item.target
          .replace('$x$', '$y$')
          .replace('Cat', 'Lynx')
          .replace('Ada', 'Grace')
          .replace('Paper', 'Updated');
        const planned = planTargetDocument(
          { pkg: before, revision: initialToken, projection: initial.view.projection },
          target
        );
        if (!planned.success) throw Error(planned.error);
        const resources = await prepareInlineImageResources(
          planned.plan,
          async () => ({ bytes: png, mime: 'image/png' }),
          decodePort
        );
        if (!resources.success) throw Error(resources.error);
        const direct = previewDirectPlan(before, planned.plan, resources.images);
        if (!direct.success) throw Error(direct.error);
        let calls = 0;
        const second = await recomposeOwnedProposals(
          editor,
          {
            revision: token(),
            targetMarkdown: target,
            ownership: first.ownership,
            contentProofs: first.contentProofs,
          },
          context,
          {
            payloads: { presence: { proposalId: 'second-cell-wave' } },
            images: {
              decodePort,
              resolve: async () => {
                calls++;
                return { bytes: png, mime: 'image/png' };
              },
            },
          }
        );
        if (!second.success) throw Error(second.error);
        expect(calls).toBe(1);
        const pending = surface.session.currentPackage(),
          secondPending = native(pending),
          accepted = bindControls(direct.pkg, pending, before);
        expect(buildBlockIndex(pending).byId.size).toBe(4);
        expect(media(pending).map(([, bytes]) => bytes)).toEqual([png]);
        expectImageRelationships(pending, buildBlockIndex(pending).byId.values());
        const groups = readReviewGroups(pending);
        expect(groups.success && groups.groups.length).toBe(1);
        const committed = await readCommittedView(surface.session, token());
        expect(committed.success && committed.view.projection.text).toBe(
          initial.view.projection.text
        );
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(native(surface.session.currentPackage())).toBe(firstPending);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(native(surface.session.currentPackage())).toBe(original);
        expect(media(surface.session.currentPackage())).toEqual([]);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(native(surface.session.currentPackage())).toBe(firstPending);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(native(surface.session.currentPackage())).toBe(secondPending);
        const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
          editor,
          token(),
          context
        );
        if (!decided.success) throw Error(decided.error);
        const resolved = surface.session.currentPackage(),
          resolvedNative = native(resolved);
        expect(resolvedNative).toBe(decision === 'accept' ? native(accepted) : original);
        expect(media(resolved).map(([, bytes]) => bytes)).toEqual(
          decision === 'accept' ? [png] : []
        );
        const store = readPayloadStore(resolved);
        expect(store.status === 'ready' ? store.entries.size : 0).toBe(
          decision === 'accept' ? 4 : 0
        );
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(native(surface.session.currentPackage())).toBe(secondPending);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(native(surface.session.currentPackage())).toBe(resolvedNative);
        const reopened = readOoxmlPackage(writeOoxmlPackage(resolved));
        if (!reopened.ok) throw Error(reopened.reason);
        expect(native(reopened.package)).toBe(resolvedNative);
        expectImageRelationships(reopened.package, buildBlockIndex(reopened.package).byId.values());
      } finally {
        editor.destroy();
      }
    });
