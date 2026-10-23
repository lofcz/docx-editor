import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import {
  presencePropertyFixture,
  installPresenceOwnership,
} from '../../test/support/presence-ownership';
import {
  propertyContext,
  propertyEditor,
  propertyToken,
  executeProperties,
} from '../../test/support/property-execution';
import { createEditPlan } from '../diff/plan';
import { buildBlockIndex } from '../blocks/block-index';
import { readReviewGroups } from './groups';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import { replaceOwnedProposals } from './replace-owned-proposals';
import { verifyOwnedProposalContentProofs } from './owned-proposal-content';

for (const decision of ['accept', 'reject'] as const)
  test(`an unrelated property ${decision} and its history preserve pending presence ownership and later inverse replacement`, () => {
    const editor = propertyEditor(presencePropertyFixture());
    try {
      const owned = installPresenceOwnership(editor, false);
      // This exercises admission while a different presence proposal remains pending.
      const property = executeProperties(editor);
      if (!property.success) throw Error(property.error);
      const rows = readReviewGroups(editor.surface!.session.currentPackage());
      if (!rows.success) throw Error(rows.error);
      expect(rows.groups).toHaveLength(2);
      const group = rows.groups.find((row) =>
        row.properties.some((member) => member.blockId === 'activity')
      )!;
      const decided = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
        editor,
        group.id,
        propertyToken(editor),
        propertyContext
      );
      if (!decided.success) throw Error(decided.error);
      const current = editor.surface!.session.currentPackage();
      expect(buildBlockIndex(current).byId.get('activity')!.payload.attributes.duration).toBe(
        decision === 'accept' ? 15 : 10
      );
      expect(
        verifyOwnedProposalContentProofs(
          current,
          propertyToken(editor),
          owned.ownership,
          owned.contentProofs
        )
      ).toBeNull();
      const accepted = canonicalOoxmlFingerprint(editor.surface!.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      const undone = readReviewGroups(editor.surface!.session.currentPackage());
      expect(undone.success && undone.groups.length).toBe(2);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(accepted);
      const empty = createEditPlan(propertyToken(editor), []);
      if (!empty.success) throw Error(empty.error);
      const inverse = replaceOwnedProposals(editor, owned.ownership, empty.plan, propertyContext);
      if (!inverse.success) throw Error(inverse.error);
      const index = buildBlockIndex(editor.surface!.session.currentPackage());
      expect(index.byId.has('outer')).toBe(false);
      expect(index.byId.has('inner')).toBe(false);
      expect(index.byId.get('activity')!.payload.attributes.duration).toBe(
        decision === 'accept' ? 15 : 10
      );
    } finally {
      editor.destroy();
    }
  }, 30000);
