import { expect, test } from 'bun:test';
import { readOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { groupToken } from '../../test/support/review-groups';
import { readReviewGroups } from './groups';
import { recomposeOwnedProposals } from './recompose';
import { rejectAllReviewGroups } from './reject-all';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';

const context = { ...propertyContext, date: '2026-10-21T12:00:00.000Z' };
const body =
  '<w:p><w:r><w:t>Alpha one</w:t></w:r></w:p><w:p><w:r><w:t>Keep</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Beta two</w:t></w:r></w:p>';
const target = 'Gamma one\n\nKeep\n\nDelta two\n\nNew tail\n';

async function propose(groupNativeChanges?: boolean) {
  const read = readOoxmlPackage(docxFixture(body));
  if (!read.ok) throw new Error(read.reason);
  const editor = propertyEditor(read.package);
  const result = await recomposeOwnedProposals(
    editor,
    { revision: groupToken(editor), targetMarkdown: target, ownership: [] },
    context,
    groupNativeChanges === undefined ? {} : { groupNativeChanges }
  );
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  const groups = readReviewGroups(editor.surface!.session.currentPackage());
  if (!groups.success) throw new Error(groups.error);
  return { editor, groups: groups.groups };
}
const views = async (editor: Awaited<ReturnType<typeof propose>>['editor']) => {
  const token = groupToken(editor),
    session = editor.surface!.session;
  const committed = await readCommittedView(session, token),
    working = await readWorkingView(session, token);
  if (!committed.success || !working.success) throw new Error('Missing review views.');
  return [committed.view.projection.text, working.view.projection.text];
};

test('a recomposition keeps independent native changes as separate review groups by default', async () => {
  const { groups } = await propose();
  expect(groups.length).toBeGreaterThan(1);
});

test('a grouped recomposition is one review decision with the same views', async () => {
  const separate = await propose(false);
  const grouped = await propose(true);
  expect(grouped.groups).toHaveLength(1);
  expect(await views(grouped.editor)).toEqual(await views(separate.editor));
  const token = groupToken(grouped.editor);
  const rejected = rejectAllReviewGroups(grouped.editor, token, context);
  expect(rejected.success).toBe(true);
  const [committed, working] = await views(grouped.editor);
  expect(working).toBe(committed);
  expect(committed).toContain('Alpha one');
});

test('a grouped recomposition refuses a non-boolean grouping', async () => {
  const read = readOoxmlPackage(docxFixture(body));
  if (!read.ok) throw new Error(read.reason);
  const editor = propertyEditor(read.package);
  const result = await recomposeOwnedProposals(
    editor,
    { revision: groupToken(editor), targetMarkdown: target, ownership: [] },
    context,
    { groupNativeChanges: 'yes' as unknown as boolean }
  );
  expect(result.success).toBe(false);
});
