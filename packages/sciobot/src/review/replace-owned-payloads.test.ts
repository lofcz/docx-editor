import type { BlockPayload } from '../blocks/payload-schema';
import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  atomicFixture,
  atomicPlan,
  atomicVisual,
  atomicListening,
  proposedVisual,
  proposedListening,
  atomicLabel,
} from '../../test/support/atomic-execution';
import {
  propertyContext,
  propertyEditor,
  propertyToken,
} from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { replaceOwnedProposals } from './replace-owned-proposals';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { readWorkingView } from './working-view';
import { groups } from '../../test/support/review-groups';

for (const action of ['accept', 'reject'] as const)
  test(`owned atomic replacement ${action} retains complete assets, opaque data, labels, history, and Word reopen`, async () => {
    const editor = propertyEditor(atomicFixture());
    const secondTargets: BlockPayload[] = [
      {
        ...proposedVisual,
        attributes: {
          ...proposedVisual.attributes,
          src: 'vfs://second.png',
          fileId: 'second-file',
        },
      },
      { ...proposedListening, attributes: { cache: { asset: 'second-cached-audio' } } },
    ];
    const options = (proposalId: string, prefix: string) => ({
      atomic: {
        proposalId,
        fallbackTargets: ['visual', 'audio'].map((blockId) => ({
          blockId,
          before: `Old ${blockId} fallback`,
          after: `${prefix} ${blockId} fallback`,
        })),
      },
    });
    const payloads = () => {
      const index = buildBlockIndex(editor.surface!.session.currentPackage());
      return ['visual', 'audio'].map((id) => index.byId.get(id)!.payload);
    };
    try {
      const first = replaceOwnedProposals(
        editor,
        [],
        atomicPlan(editor, [proposedVisual, proposedListening]),
        propertyContext,
        options('atomic-wave-one', 'First')
      );
      if (!first.success) throw Error(first.error);
      const firstView = await readWorkingView(editor.surface!.session, propertyToken(editor));
      if (!firstView.success) throw Error(firstView.error);
      const second = replaceOwnedProposals(
        editor,
        first.ownership,
        atomicPlan(editor, secondTargets),
        propertyContext,
        options('atomic-wave-two', 'Second')
      );
      if (!second.success) throw Error(second.error);
      expect(payloads()).toEqual([atomicVisual, atomicListening]);
      const pending = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const secondView = await readWorkingView(editor.surface!.session, propertyToken(editor));
      if (!secondView.success) throw Error(secondView.error);
      expect(secondView.view.projection.text).toContain('src="vfs://second.png"');
      expect(secondView.view.projection.text).toContain('Good evening!');
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      const prior = await readWorkingView(editor.surface!.session, propertyToken(editor));
      if (!prior.success) throw Error(prior.error);
      expect(prior.view.projection.text).toBe(firstView.view.projection.text);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
      const decided = (action === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
        editor,
        propertyToken(editor),
        propertyContext
      );
      if (!decided.success) throw Error(decided.error);
      expect(payloads()).toEqual(
        action === 'accept' ? secondTargets : [atomicVisual, atomicListening]
      );
      for (const id of ['visual', 'audio'])
        expect(paragraphTextOf(editor.surface!.session.part(), atomicLabel(editor, id).id)).toBe(
          `${action === 'accept' ? 'Second' : 'Old'} ${id} fallback`
        );
      expect(groups(editor)).toHaveLength(0);
      const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!saved.ok) throw Error(saved.reason);
      const index = buildBlockIndex(saved.package);
      expect(['visual', 'audio'].map((id) => index.byId.get(id)!.payload)).toEqual(
        action === 'accept' ? secondTargets : [atomicVisual, atomicListening]
      );
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
      expect(payloads()).toEqual([atomicVisual, atomicListening]);
    } finally {
      editor.destroy();
    }
  });
