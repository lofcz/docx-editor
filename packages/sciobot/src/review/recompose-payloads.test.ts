import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  propertyEditor,
  propertyContext,
  propertyToken,
} from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { recomposeOwnedProposals } from './recompose';
import { groups } from '../../test/support/review-groups';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';

for (const action of ['accept', 'reject'] as const)
  test(`two property and body waves ${action} together and preserve intervening manual text`, async () => {
    const editor = propertyEditor();
    const view = async (committed = false) => {
      const result = await (committed ? readCommittedView : readWorkingView)(
        editor.surface!.session,
        propertyToken(editor)
      );
      if (!result.success) throw Error(result.error);
      return result.view.projection.text;
    };
    try {
      const initial = await view();
      expect(initial).toContain('name="Old activity"');
      const firstTarget = initial
        .replace('name="Old activity"', 'name="First activity"')
        .replace('Body', 'Body first');
      let publications = 0;
      const stop = editor.surface!.session.subscribe(() => publications++);
      const first = await recomposeOwnedProposals(
        editor,
        {
          revision: propertyToken(editor),
          targetMarkdown: firstTarget,
          ownership: [],
          contentProofs: [],
        },
        propertyContext,
        {
          payloads: {
            properties: {
              proposalId: 'wave-one',
              fallbackTargets: [
                {
                  blockId: 'activity',
                  before: 'Old activity — 10 minutes',
                  after: 'First activity — 10 minutes',
                },
              ],
            },
          },
        }
      );
      stop();
      if (!first.success) throw Error(first.error);
      expect(publications).toBe(1);
      expect(await view()).toBe(firstTarget);
      expect(await view(true)).toBe(initial);
      const outsideId = editor
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(editor.surface!.session.part(), id) === 'Outside')!;
      expect(
        editor.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: outsideId, offset: 7, text: ' manual' },
        ]).rejected
      ).toBe(false);
      const firstWithManual = firstTarget.replace('Outside', 'Outside manual');
      const secondTarget = initial
        .replace('name="Old activity"', 'name="Second activity"')
        .replace('Body', 'Body second')
        .replace('Outside', 'Outside manual');
      const second = await recomposeOwnedProposals(
        editor,
        {
          revision: propertyToken(editor),
          targetMarkdown: secondTarget,
          ownership: first.ownership,
          contentProofs: first.contentProofs,
        },
        propertyContext,
        {
          payloads: {
            properties: {
              proposalId: 'wave-two',
              fallbackTargets: [
                {
                  blockId: 'activity',
                  before: 'Old activity — 10 minutes',
                  after: 'Second activity — 10 minutes',
                },
              ],
            },
          },
        }
      );
      if (!second.success) throw Error(second.error);
      expect(await view()).toBe(secondTarget);
      const original = initial.replace('Outside', 'Outside manual');
      expect(await view(true)).toBe(original);
      expect(
        buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('activity')!.payload
          .attributes.name
      ).toBe('Old activity');
      const pending = canonicalOoxmlFingerprint(editor.surface!.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(await view()).toBe(firstWithManual);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
      const result = (action === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
        editor,
        propertyToken(editor),
        propertyContext
      );
      if (!result.success) throw Error(result.error);
      expect(await view()).toBe(action === 'accept' ? secondTarget : original);
      expect(await view(true)).toBe(action === 'accept' ? secondTarget : original);
      expect(groups(editor)).toHaveLength(0);
      expect(
        buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('activity')!.payload
          .attributes.name
      ).toBe(action === 'accept' ? 'Second activity' : 'Old activity');
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
      expect(await view()).toBe(secondTarget);
      const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!saved.ok) throw Error(saved.reason);
      const reopened = propertyEditor(saved.package);
      try {
        const read = await readWorkingView(reopened.surface!.session, propertyToken(reopened));
        if (!read.success) throw Error(read.error);
        expect(read.view.projection.text).toBe(secondTarget);
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  });

for (const mismatch of ['missing-fallback', 'stale-label'] as const)
  test(`later property ${mismatch} rolls back native edits, selected inverses, metadata, and history`, async () => {
    const editor = propertyEditor();
    try {
      const initial = await readWorkingView(editor.surface!.session, propertyToken(editor));
      if (!initial.success) throw Error(initial.error);
      const markdown = initial.view.projection.text;
      const first = await recomposeOwnedProposals(
        editor,
        {
          revision: propertyToken(editor),
          targetMarkdown: markdown
            .replace('name="Old activity"', 'name="First activity"')
            .replace('Body', 'Body first'),
          ownership: [],
          contentProofs: [],
        },
        propertyContext,
        {
          payloads: {
            properties: {
              proposalId: 'rollback-one',
              fallbackTargets: [
                {
                  blockId: 'activity',
                  before: 'Old activity — 10 minutes',
                  after: 'First activity — 10 minutes',
                },
              ],
            },
          },
        }
      );
      if (!first.success) throw Error(first.error);
      const session = editor.surface!.session;
      const before = session.currentPackage();
      const token = propertyToken(editor);
      const groupIds = groups(editor).map((group) => group.id);
      const undo = session.canUndo(),
        redo = session.canRedo();
      const failed = await recomposeOwnedProposals(
        editor,
        {
          revision: token,
          targetMarkdown: markdown
            .replace('name="Old activity"', 'name="Second activity"')
            .replace('Body', 'Body second'),
          ownership: first.ownership,
          contentProofs: first.contentProofs,
        },
        propertyContext,
        mismatch === 'missing-fallback'
          ? {}
          : {
              payloads: {
                properties: {
                  proposalId: 'rollback-two',
                  fallbackTargets: [
                    {
                      blockId: 'activity',
                      before: 'First activity — 10 minutes',
                      after: 'Second activity — 10 minutes',
                    },
                  ],
                },
              },
            }
      );
      expect(failed.success).toBe(false);
      if (failed.success) throw Error('Unexpected replacement');
      expect(failed.code).toBe(
        mismatch === 'missing-fallback' ? 'unsupported-operation' : 'stale-revision'
      );
      expect(session.currentPackage()).toBe(before);
      expect(propertyToken(editor)).toEqual(token);
      expect(groups(editor).map((group) => group.id)).toEqual(groupIds);
      expect(session.canUndo()).toBe(undo);
      expect(session.canRedo()).toBe(redo);
    } finally {
      editor.destroy();
    }
  });
