// The test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { propertyRecordFixture } from '../../test/support/property-record';
import {
  propertyContext,
  propertyEditor,
  propertyToken,
} from '../../test/support/property-execution';
import { readPayloadGroupDecisions } from './payload-group-decisions';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { recomposeOwnedProposals } from './recompose';
import { readReviewGroups } from './groups';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

async function views(editor: DocxEditorInstance) {
  const before = await readCommittedView(editor.surface!.session, propertyToken(editor));
  const after = await readWorkingView(editor.surface!.session, propertyToken(editor));
  if (!before.success || !after.success) throw Error('Missing native views');
  return [before.view.projection.text, after.view.projection.text] as const;
}
const groupsOf = (editor: DocxEditorInstance) => {
  const rows = readReviewGroups(editor.surface!.session.currentPackage());
  if (!rows.success) throw Error(rows.error);
  return rows.groups;
};
/** A changed `resource` replaces the activity block: one removal and one adjacent insertion. */
async function replaceActivity(editor: DocxEditorInstance) {
  const [, initial] = await views(editor);
  const target = initial.replace('resource="keep"', 'resource="Lab"');
  expect(target).not.toBe(initial);
  const result = await recomposeOwnedProposals(
    editor,
    { revision: propertyToken(editor), targetMarkdown: target, ownership: [], contentProofs: [] },
    propertyContext,
    { createPresenceProposalId: () => 'replacement-wave' }
  );
  if (!result.success) throw Error(`${result.code}: ${result.error}`);
  return { initial, target };
}

for (const action of ['accept', 'reject'] as const)
  test(`an activity replacement in one presence record is one review group (${action})`, async () => {
    const editor = propertyEditor();
    const { initial, target } = await replaceActivity(editor);
    const [group, ...rest] = groupsOf(editor);
    expect(rest).toEqual([]);
    expect(group!.readOnly).toBe(false);
    expect(group!.properties.map((member) => member.proposalId)).toEqual([
      'replacement-wave',
      'replacement-wave',
    ]);
    expect(group!.properties.map((member) => member.blockId)).toContain('activity');
    expect(await views(editor)).toEqual([initial, target]);
    const decided = (action === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
      editor,
      group!.id,
      propertyToken(editor),
      propertyContext
    );
    if (!decided.success) throw Error(decided.error);
    const resolved = action === 'accept' ? target : initial;
    expect(await views(editor)).toEqual([resolved, resolved]);
    expect(groupsOf(editor)).toEqual([]);
    // The decision covers each block key, so a reader without pairing sees both decided.
    const decisions = readPayloadGroupDecisions(editor.surface!.session.currentPackage());
    expect(decisions.status === 'ready' && decisions.decisions.size).toBe(2);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groupsOf(editor).map((row) => row.id)).toEqual([group!.id]);
    expect(await views(editor)).toEqual([initial, target]);
  });

test('a removal and an insertion that are not adjacent stay separate groups', async () => {
  const editor = propertyEditor();
  const [, initial] = await views(editor);
  // The new activity follows Outside, so the removed activity is not its sibling.
  const target =
    initial.replace(/<activity[^>]*>\s*\*\*Body\*\*\s*<\/activity>\s*/, '') +
    '\n<activity name="Other" duration="5">\n\nNew body\n\n</activity>\n';
  const result = await recomposeOwnedProposals(
    editor,
    { revision: propertyToken(editor), targetMarkdown: target, ownership: [], contentProofs: [] },
    propertyContext,
    { createPresenceProposalId: () => 'separate-wave' }
  );
  if (!result.success) throw Error(`${result.code}: ${result.error}`);
  expect(groupsOf(editor).length).toBeGreaterThan(1);
});

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const action of ['accept', 'reject'] as const)
    test(`a replacement group converges with peer typing, history, reconnect, and reopening (${clients}, ${action})`, async () => {
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(propertyRecordFixture()),
        documentId: propertyContext.documentId,
        ownedAttributes: true,
        clientIds: clients,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const { initial, target } = await replaceActivity(alice);
        await pair.settle();
        const ids = pair.editors.map((editor) => groupsOf(editor).map((row) => row.id));
        expect(ids[0]).toHaveLength(1);
        expect(ids[1]).toEqual(ids[0]!);
        const outside = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Outside')!;
        pair.pause();
        expect(
          bob.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: outside, offset: 7, text: ' Peer' },
          ]).rejected
        ).toBe(false);
        const decided = (action === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
          alice,
          ids[0]![0]!,
          propertyToken(alice),
          propertyContext
        );
        if (!decided.success) throw Error(decided.error);
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const resolved = (action === 'accept' ? target : initial).replace(
          '**Outside**',
          '**Outside Peer**'
        );
        for (const editor of pair.editors) {
          expect(groupsOf(editor)).toEqual([]);
          expect(await views(editor)).toEqual([resolved, resolved]);
        }
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(groupsOf(editor).map((row) => row.id)).toEqual(ids[0]!);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) expect(groupsOf(editor)).toEqual([]);
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        const reopened = await createTwoEditorFixture({
          document: new Uint8Array(await bob.save()),
          documentId: propertyContext.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
        });
        try {
          await reopened.settle();
          for (const editor of reopened.editors)
            expect(await views(editor)).toEqual([resolved, resolved]);
          reopened.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          reopened.destroy();
        }
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 30000);
