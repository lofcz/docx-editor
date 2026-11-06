import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  withBinaryPart,
  withPart,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { galleryPng } from '../../test/support/native-gallery';
import { crossStoryGroupFixture } from '../../test/support/group-decisions';
import { readReviewGroups } from './groups';
import { groupRemovalBudget } from './group-removal-budget';

test('new image descriptors reserve their one retained override flag without a general allowance', () => {
  const before = crossStoryGroupFixture();
  const once = withBinaryPart(before, '/word/media/new.png', galleryPng, 'image/png');
  expect(groupRemovalBudget(before, once)).toBe(1);
  expect(groupRemovalBudget(once, once)).toBe(0);
  expect(
    groupRemovalBudget(once, withBinaryPart(once, '/word/media/new.png', galleryPng, 'image/png'))
  ).toBe(0);
  const twice = withBinaryPart(once, '/word/media/second.png', galleryPng, 'image/png');
  expect(groupRemovalBudget(once, twice)).toBe(1);
  expect(groupRemovalBudget(before, twice)).toBe(2);
  expect(
    groupRemovalBudget(
      before,
      withBinaryPart(before, '/opaque.bin', galleryPng, 'application/octet-stream')
    )
  ).toBe(0);
});

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
