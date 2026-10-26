import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { planDomainProperties } from '../diff/domain-properties';
import {
  propertyEditor,
  propertyContext,
  propertyToken,
} from '../../test/support/property-execution';
import { recomposeOwnedProperties } from './recompose-properties';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { readReviewGroups } from './groups';
import type { OwnedProposalReplacementResult } from './replace-owned-proposals';

for (const action of ['accept', 'reject'] as const)
  test(`explicit resource field waves retain native body, decisions, reopen, and history: ${action}`, async () => {
    const editor = propertyEditor();
    try {
      const session = editor.surface!.session;
      const original = buildBlockIndex(session.currentPackage()).byId.get('activity')!;
      const bodyId = session
        .paragraphIds()
        .find((id) => paragraphTextOf(session.part(), id) === 'Body')!;
      const body = canonicalOoxmlFingerprint(findNode(session.part(), bodyId)!);
      const source = () => ({
        pkg: session.currentPackage(),
        revision: propertyToken(editor),
        blockId: 'activity',
      });
      // Frozen AI pairing behavior remains strict, even after the field path is available.
      expect(planDomainProperties(source(), { set: { resource: 'Workbook' } }).success).toBe(false);
      let ledger: Extract<OwnedProposalReplacementResult, { success: true }> | undefined;
      const change = (name: string, resource: string, id: string) => {
        const result = recomposeOwnedProperties(
          editor,
          {
            revision: propertyToken(editor),
            blockId: 'activity',
            target: { set: { name, duration: 10, resource } },
            ownership: ledger?.ownership ?? [],
            contentProofs: ledger?.contentProofs ?? [],
            proposalId: id,
            fallback: {
              blockId: 'activity',
              before: 'Old activity — 10 minutes',
              after: `${name} — 10 minutes`,
            },
          },
          propertyContext
        );
        if (!result.success) throw Error(result.error);
        ledger = result;
      };
      change('Old activity', 'Workbook', 'field-first');
      change('New activity', 'Workbook', 'field-second');
      change('New activity', 'Second workbook', 'field-third');
      const read = async (committed = false) => {
        const result = await (committed ? readCommittedView : readWorkingView)(
          session,
          propertyToken(editor)
        );
        if (!result.success) throw Error(result.error);
        return result.view.projection.text;
      };
      const committed = await read(true),
        proposed = await read();
      expect(committed).toContain('resource="keep"');
      expect(proposed).toContain('resource="Second workbook"');
      expect(proposed).toContain('name="New activity"');
      expect(buildBlockIndex(session.currentPackage()).byId.get('activity')!.payload).toEqual(
        original.payload
      );
      const groups = readReviewGroups(session.currentPackage());
      if (!groups.success) throw Error(groups.error);
      expect(groups.groups).toHaveLength(1);
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw Error(reopened.reason);
      const reopenedEditor = propertyEditor(reopened.package);
      try {
        const pending = await readWorkingView(
          reopenedEditor.surface!.session,
          propertyToken(reopenedEditor)
        );
        if (!pending.success) throw Error(pending.error);
        expect(pending.view.projection.text).toBe(proposed);
      } finally {
        reopenedEditor.destroy();
      }
      const decision = (action === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
        editor,
        propertyToken(editor),
        propertyContext
      );
      if (!decision.success) throw Error(decision.error);
      expect(await read()).toBe(action === 'accept' ? proposed : committed);
      expect(canonicalOoxmlFingerprint(findNode(session.part(), bodyId)!)).toBe(body);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(await read()).toBe(proposed);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(await read()).toBe(action === 'accept' ? proposed : committed);
    } finally {
      editor.destroy();
    }
  });

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const decision of ['accept', 'reject'] as const)
    test(`resource field replacement preserves offline peer content: ${clients}/${decision}`, async () => {
      const { createTwoEditorFixture } = await import('../../test/support/collaboration');
      const { propertyRecordFixture } = await import('../../test/support/property-record');
      const { writeOoxmlPackage } = await import('@docx-editor.dev/core/store');
      const { SCIOBOT_NATIVE_REVIEW_MODULE } = await import('./native-module');
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(propertyRecordFixture()),
        documentId: propertyContext.documentId,
        ownedAttributes: true,
        clientIds: clients,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors;
        let ledger: Extract<OwnedProposalReplacementResult, { success: true }> | undefined;
        const change = (resource: string, proposalId: string) => {
          const result = recomposeOwnedProperties(
            alice!,
            {
              revision: propertyToken(alice!),
              blockId: 'activity',
              target: { set: { resource } },
              ownership: ledger?.ownership ?? [],
              contentProofs: ledger?.contentProofs ?? [],
              proposalId,
              fallback: {
                blockId: 'activity',
                before: 'Old activity — 10 minutes',
                after: 'Old activity — 10 minutes',
              },
            },
            propertyContext
          );
          if (!result.success) throw Error(result.error);
          ledger = result;
        };
        change('Workbook A', 'field-a');
        await pair.settle();
        pair.pause();
        const bodyId = bob!
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'Body')!;
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: bodyId, offset: 4, text: ' peer' },
          ]).rejected
        ).toBe(false);
        change('Workbook B', 'field-b');
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const check = async (resource: string, groups: number) => {
          for (const editor of pair.editors) {
            const read = await readWorkingView(editor.surface!.session, propertyToken(editor));
            if (!read.success) throw Error(read.error);
            expect(read.view.projection.text).toContain(`resource="${resource}"`);
            expect(paragraphTextOf(editor.surface!.session.part(), bodyId)).toBe('Body peer');
            const pending = readReviewGroups(editor.surface!.session.currentPackage());
            if (!pending.success) throw Error(pending.error);
            expect(pending.groups).toHaveLength(groups);
          }
        };
        await check('Workbook B', 1);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check('Workbook A', 1);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check('Workbook B', 1);
        const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
          bob!,
          propertyToken(bob!),
          propertyContext
        );
        if (!decided.success) throw Error(decided.error);
        await pair.settle();
        await check(decision === 'accept' ? 'Workbook B' : 'keep', 0);
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check('Workbook B', 1);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(decision === 'accept' ? 'Workbook B' : 'keep', 0);
        pair.ownership.forEach((scope) => scope.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 30000);
