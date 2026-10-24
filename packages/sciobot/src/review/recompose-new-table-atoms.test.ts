import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  readOoxmlPackage,
  writeOoxmlPackage,
  applyTreeOp,
  withPart,
  paragraphTextOf,
} from '@docx-editor.dev/core/store';
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
  atoms,
  preparedCellDocument,
} from '../../test/support/tracked-cell-atoms';
const body =
  '<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:t>Old</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>';
const table = `| A | B |\n| - | - |\n| ${atoms} | Tail |`;
const cases = [
  { name: 'middle table', target: `Before\n\n${table}\n\nOld\n\nAfter`, count: 4 },
  { name: 'terminal table', target: `Before\n\nOld\n\nAfter\n\n${table}`, count: 4 },
  {
    name: 'mixed two-table gap',
    target: `Before\n\nIntro $z$\n\n${table}\n\n| C | D |\n| - | - |\n| E | F |\n\nEnd\n\nOld\n\nAfter`,
    count: 5,
  },
  { name: 'paragraph replacement with table', target: `Before\n\n${table}\n\nAfter`, count: 4 },
];
for (const item of cases)
  for (const decision of ['accept', 'reject'] as const)
    test(`tracked ${item.name} atoms preserve native decisions, resources, history, and reopening: ${decision}`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: preparedCellDocument(body),
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
        let identitySource = pending;
        // The pending source still holds the explicitly removed original paragraph.
        // Exclude only that known deletion from fresh identity pairing, before any decision.
        if (item.name === 'paragraph replacement with table') {
          const removed = surface.session
            .paragraphIds()
            .find(
              (id) => paragraphTextOf(before.parts.get(before.mainDocumentPart)!, id) === 'Old'
            );
          if (!removed) throw Error('Missing original replacement paragraph.');
          const clean = applyTreeOp(pending.parts.get(pending.mainDocumentPart)!, {
            op: 'deleteBlock',
            blockId: removed,
          });
          if (!clean.ok) throw Error(clean.reason);
          identitySource = withPart(pending, clean.part);
        }
        const accepted = bindControls(direct.pkg, identitySource, before);
        expect(
          [...buildBlockIndex(pending).byId.values()].map((block) => block.payload.type)
        ).toEqual(
          item.count === 5
            ? ['inline_equation', 'inline_equation', 'inline_image', 'mention', 'source']
            : ['inline_equation', 'inline_image', 'mention', 'source']
        );
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
        expect(buildBlockIndex(resolved).byId.size).toBe(decision === 'accept' ? item.count : 0);
        const store = readPayloadStore(resolved);
        expect(store.status === 'ready' ? store.entries.size : 0).toBe(
          decision === 'accept' ? item.count : 0
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
    }, 30000);
