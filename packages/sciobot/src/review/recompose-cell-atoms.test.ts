import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
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
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

import {
  context,
  png,
  decodePort,
  media,
  native,
  allocate,
  bindControls,
  cases,
} from '../../test/support/tracked-cell-atoms';
for (const item of cases)
  for (const decision of ['accept', 'reject'] as const)
    test(`tracked ${item.name} atoms preserve native decisions, resources, history, and reopening: ${decision}`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: docxFixture(item.body),
        imageDecodePort: decodePort,
        modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      });
      try {
        const surface = editor.surface!,
          revision = () => readEditorRevision(editor, context.documentId, context.epoch)!;
        const before = surface.session.currentPackage(),
          original = native(before);
        const view = await readWorkingView(surface.session, revision());
        if (!view.success) throw Error(view.error);
        const planned = planTargetDocument(
          { pkg: before, revision: revision(), projection: view.view.projection },
          item.target,
          { allocateBlockId: allocate() }
        );
        if (!planned.success) throw Error(planned.error);
        const prepared = await prepareInlineImageResources(
          planned.plan,
          async () => ({ bytes: png, mime: 'image/png' }),
          decodePort
        );
        if (!prepared.success) throw Error(prepared.error);
        const direct = previewDirectPlan(before, planned.plan, prepared.images);
        if (!direct.success) throw Error(direct.error);
        const result = await recomposeOwnedProposals(
          editor,
          { revision: revision(), targetMarkdown: item.target, ownership: [] },
          context,
          {
            target: { allocateBlockId: allocate() },
            payloads: { presence: { proposalId: 'cell-presence' } },
            images: { decodePort, resolve: async () => ({ bytes: png, mime: 'image/png' }) },
          }
        );
        if (!result.success) throw Error(result.error);
        const pending = surface.session.currentPackage(),
          pendingNative = native(pending);
        const accepted = bindControls(direct.pkg, pending, before);
        expect(
          [...buildBlockIndex(pending).byId.values()].map((block) => block.payload.type)
        ).toEqual(['inline_equation', 'inline_image', 'mention', 'source']);
        expect(media(pending)).toHaveLength(1);
        expectImageRelationships(pending, buildBlockIndex(pending).byId.values());
        const committed = await readCommittedView(surface.session, revision());
        expect(committed.success && committed.view.projection.text).toBe(view.view.projection.text);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(native(surface.session.currentPackage())).toBe(original);
        expect(media(surface.session.currentPackage())).toEqual(media(before));
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(native(surface.session.currentPackage())).toBe(pendingNative);
        const decide = decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups;
        const decided = decide(editor, revision(), context);
        if (!decided.success) throw Error(decided.error);
        const resolved = surface.session.currentPackage(),
          resolvedNative = native(resolved);
        expect(resolvedNative).toBe(decision === 'accept' ? native(accepted) : original);
        expect(media(resolved).map(([, bytes]) => bytes)).toEqual(
          decision === 'accept' ? [png] : []
        );
        expect(buildBlockIndex(resolved).byId.size).toBe(decision === 'accept' ? 4 : 0);
        const store = readPayloadStore(resolved);
        expect(store.status === 'ready' ? store.entries.size : 0).toBe(
          decision === 'accept' ? 4 : 0
        );
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(native(surface.session.currentPackage())).toBe(pendingNative);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(native(surface.session.currentPackage())).toBe(resolvedNative);
        const opened = readOoxmlPackage(writeOoxmlPackage(resolved));
        if (!opened.ok) throw Error(opened.reason);
        expect(native(opened.package)).toBe(resolvedNative);
        expect(media(opened.package).map(([, bytes]) => bytes)).toEqual(
          media(resolved).map(([, bytes]) => bytes)
        );
      } finally {
        editor.destroy();
      }
    });
