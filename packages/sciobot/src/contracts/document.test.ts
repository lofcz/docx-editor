import { expect, test } from 'bun:test';
import { sameRevision, type NativeEditFailure, type RevisionToken } from './document';

const revision: RevisionToken = { documentId: 'lesson-1', epoch: 'mount-1', revision: 3 };

test('revision equality includes document identity and load epoch', () => {
  expect(sameRevision(revision, { ...revision })).toBe(true);
  expect(sameRevision(revision, { ...revision, revision: 4 })).toBe(false);
  expect(sameRevision(revision, { ...revision, documentId: 'lesson-2' })).toBe(false);
  expect(sameRevision(revision, { ...revision, epoch: 'mount-2' })).toBe(false);
});

test('stale reads and unsupported edits retain distinct serializable failures', () => {
  const stale: NativeEditFailure = {
    success: false,
    code: 'stale-revision',
    error: 'Read the current document before editing.',
    expectedRevision: revision,
    actualRevision: { ...revision, revision: 4 },
  };
  const unsupported: NativeEditFailure = {
    success: false,
    code: 'unsupported-operation',
    error: 'The selected boundary does not support this operation.',
  };
  expect(JSON.parse(JSON.stringify(stale))).toEqual(stale);
  expect(JSON.parse(JSON.stringify(unsupported))).toEqual(unsupported);
  expect(stale.code).not.toBe(unsupported.code);
});
