import { expect, test } from 'bun:test';
import { applyTreeOp, withPart, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { crossStoryGroupFixture } from '../../test/support/group-decisions';
import { readReviewGroups } from './groups';
import { groupRemovalBudget } from './group-removal-budget';

test('the canonical removal budget counts the two unwrapped Word revision wrappers, preserving their text and other groups', () => {
  const before = crossStoryGroupFixture(),
    bytes = writeOoxmlPackage(before),
    read = readReviewGroups(before);
  if (!read.success) throw new Error(read.error);
  const group = read.groups.find((group) => group.id === 'g1:stored:rg_cross_story')!;
  let after = before;
  for (const member of group.native) {
    const result = applyTreeOp(after.parts.get(member.partName)!, {
      op: 'acceptRevision',
      revision: member.address,
      siteNodeIds: member.siteNodeIds,
    });
    if (!result.ok) throw new Error(result.reason);
    after = withPart(after, result.part);
  }
  expect(groupRemovalBudget(before, before)).toBe(0);
  expect(groupRemovalBudget(before, after)).toBe(2);
  expect(groupRemovalBudget(after, before)).toBe(0);
  expect(writeOoxmlPackage(before)).toEqual(bytes);
  expect(readReviewGroups(after)).toMatchObject({
    success: true,
    groups: [
      expect.objectContaining({
        native: [expect.objectContaining({ address: { id: '8', author: 'Peer' } })],
      }),
    ],
  });
});

test('removal budgets include every descendant of a removed subtree or part, without mutating the original package', () => {
  const before = crossStoryGroupFixture(),
    header = before.parts.get('/word/header1.xml')!,
    bytes = writeOoxmlPackage(before),
    emptied = withPart(before, { ...header, root: { ...header.root, children: [] } });
  expect(groupRemovalBudget(before, emptied)).toBe(5);
  // Every stage counts: a node created then removed has no final-package difference.
  expect(groupRemovalBudget(emptied, before) + groupRemovalBudget(before, emptied)).toBe(5);
  expect(groupRemovalBudget(emptied, emptied)).toBe(0);
  const parts = new Map(before.parts);
  parts.delete(header.name);
  expect(groupRemovalBudget(before, { ...before, parts })).toBe(6);
  expect(writeOoxmlPackage(before)).toEqual(bytes);
});
