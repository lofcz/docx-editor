import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import {
  acceptedMixedPackage,
  bulkSource,
  payloadSnapshot,
  proposalSnapshot,
  proposeMixedBulk,
  typeBulkSentence,
} from '../../test/support/bulk-decisions';
import { propertyContext, propertyEditor } from '../../test/support/property-execution';
import { groupFixture, groupToken, groups } from '../../test/support/review-groups';
import {
  packageWithPropertyRecord,
  propertyRecord,
  propertyRecordFixture,
} from '../../test/support/property-record';
import { crossStoryGroupFixture } from '../../test/support/group-decisions';
import { docxFixture } from '../../test/support/document';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { restoreReviewState } from './restore';

const word = (editor: DocxEditorInstance) =>
  canonicalOoxmlFingerprint(editor.surface!.session.part());
function restored(editor: DocxEditorInstance, epoch = 'reload-2') {
  const result = restoreReviewState(bulkSource(editor), { ...groupToken(editor), epoch });
  if (!result.success) throw new Error(result.error);
  return result.state;
}
async function reopen(editor: DocxEditorInstance) {
  const read = readOoxmlPackage(new Uint8Array(await editor.save()));
  if (!read.ok) throw new Error(read.reason);
  return propertyEditor(read.package);
}

test('reload restores nested text, property and atomic groups with fresh native owners and one decision undo', async () => {
  const original = propertyEditor(groupFixture());
  try {
    proposeMixedBulk(original, () => typeBulkSentence(original));
    const ids = groups(original).map((group) => group.id),
      raw = proposalSnapshot(bulkSource(original));
    const editor = await reopen(original);
    try {
      const state = restored(editor),
        pendingWord = word(editor),
        pendingPayloads = payloadSnapshot(bulkSource(editor)),
        expected = acceptedMixedPackage(bulkSource(editor)),
        index = buildBlockIndex(bulkSource(editor));
      expect(state.hasPendingReview).toBe(true);
      expect(state.groups.map((group) => group.id)).toEqual(ids);
      expect(state.reviewableGroupIds).toEqual(ids);
      expect(state.groups).toHaveLength(6);
      expect(state.pendingProperties).toHaveLength(2);
      expect(editor.surface!.session.canUndo()).toBe(false);
      for (const proposal of state.pendingProperties) {
        expect(proposal.plan.revision).toEqual(state.revision);
        for (const operation of proposal.plan.operations) {
          if (operation.kind !== 'payload') throw new Error('Unexpected recovered operation');
          expect(operation.controlId).toBe(index.byId.get(operation.blockId)!.controlId);
        }
      }
      expect(proposalSnapshot(bulkSource(editor))).toEqual(raw);
      expect(
        acceptAllReviewGroups(editor, state.revision, {
          ...propertyContext,
          epoch: state.revision.epoch,
        }).success
      ).toBe(true);
      expect(restored(editor).hasPendingReview).toBe(false);
      expect(word(editor)).toBe(
        canonicalOoxmlFingerprint(expected.parts.get(expected.mainDocumentPart)!)
      );
      expect(payloadSnapshot(bulkSource(editor))).toEqual(payloadSnapshot(expected));
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(restored(editor).groups.map((group) => group.id)).toEqual(ids);
      expect(word(editor)).toBe(pendingWord);
      expect(payloadSnapshot(bulkSource(editor))).toEqual(pendingPayloads);
      expect(editor.surface!.session.bodyText()).toContain('Manual sentence. Keep tail');
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      const saved = await reopen(editor);
      try {
        expect(restored(saved)).toMatchObject({
          groups: [],
          pendingProperties: [],
          reviewableGroupIds: [],
          hasPendingReview: false,
        });
      } finally {
        saved.destroy();
      }
    } finally {
      editor.destroy();
    }
  } finally {
    original.destroy();
  }
});

test('reload retains prior mixed decisions and recovers only the remaining proposal blocks', async () => {
  const original = propertyEditor(groupFixture());
  try {
    proposeMixedBulk(original);
    const first = groups(original)[0]!;
    expect(
      acceptReviewGroup(original, first.id, groupToken(original), propertyContext).success
    ).toBe(true);
    const visual = groups(original).find((group) =>
      group.properties.some((member) => member.blockId === 'visual')
    )!;
    expect(
      rejectReviewGroup(original, visual.id, groupToken(original), propertyContext).success
    ).toBe(true);
    const before = restored(original),
      payloads = payloadSnapshot(bulkSource(original)),
      editor = await reopen(original);
    try {
      const state = restored(editor);
      expect(state.groups.map((group) => group.id)).toEqual(before.groups.map((group) => group.id));
      expect(state.groups).toHaveLength(4);
      expect(state.pendingProperties).toHaveLength(2);
      expect(
        state.pendingProperties
          .flatMap((proposal) =>
            proposal.plan.operations.map((operation) =>
              operation.kind === 'payload' ? operation.blockId : 'unexpected'
            )
          )
          .sort()
      ).toEqual(['activity', 'activity', 'audio']);
      expect(payloadSnapshot(bulkSource(editor))).toEqual(payloads);
      expect(editor.surface!.session.bodyText()).toContain('The car sat.');
    } finally {
      editor.destroy();
    }
  } finally {
    original.destroy();
  }
});

for (const action of ['accept', 'reject'] as const)
  test(
    'reload of a fully ' + action + 'ed review exposes no stale groups or proposal controls',
    async () => {
      const original = propertyEditor(groupFixture());
      try {
        proposeMixedBulk(original);
        const decide = action === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups;
        expect(decide(original, groupToken(original), propertyContext).success).toBe(true);
        const editor = await reopen(original);
        try {
          expect(restored(editor)).toMatchObject({
            hasPendingReview: false,
            groups: [],
            pendingProperties: [],
            reviewableGroupIds: [],
          });
          expect(proposalSnapshot(bulkSource(editor))).toHaveLength(2);
        } finally {
          editor.destroy();
        }
      } finally {
        original.destroy();
      }
    }
  );

test('repeated recovery is immutable and creates no revision, history, metadata or display-mode change', () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeMixedBulk(editor);
    typeBulkSentence(editor);
    const pkg = bulkSource(editor),
      token = { ...groupToken(editor) },
      bytes = writeOoxmlPackage(pkg),
      before = editor.surface!.session.packageRevision(),
      ids = groups(editor).map((group) => group.id),
      mode = editor.snapshot().reviewDisplayMode;
    for (let index = 0; index < 3; index++) {
      const result = restoreReviewState(pkg, token);
      if (!result.success) throw new Error(result.error);
      expect(Object.isFrozen(result.state)).toBe(true);
      expect(Object.isFrozen(result.state.revision)).toBe(true);
      expect(Object.isFrozen(result.state.groups)).toBe(true);
      expect(Object.isFrozen(result.state.pendingProperties)).toBe(true);
      expect(Object.isFrozen(result.state.reviewableGroupIds)).toBe(true);
    }
    const result = restoreReviewState(pkg, token);
    if (!result.success) throw new Error(result.error);
    token.epoch = 'caller-mutated';
    expect(result.state.revision.epoch).toBe(propertyContext.epoch);
    expect(writeOoxmlPackage(pkg)).toEqual(bytes);
    expect(editor.surface!.session.packageRevision()).toBe(before);
    expect(editor.snapshot().reviewDisplayMode).toBe(mode);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.surface!.session.bodyText()).not.toContain('Manual sentence.');
    expect(restored(editor).groups.map((group) => group.id)).toEqual(ids);
  } finally {
    editor.destroy();
  }
});

test('imported body/header groups recover without application chat or property records', async () => {
  const original = propertyEditor(crossStoryGroupFixture());
  try {
    const editor = await reopen(original);
    try {
      const state = restored(editor);
      expect(state.hasPendingReview).toBe(true);
      expect(state.groups.map((group) => group.id)).toEqual(
        groups(original).map((group) => group.id)
      );
      expect(state.groups).toHaveLength(2);
      expect(state.pendingProperties).toEqual([]);
      expect(state.reviewableGroupIds).toHaveLength(2);
    } finally {
      editor.destroy();
    }
  } finally {
    original.destroy();
  }
});

test('ambiguous imported groups remain visible and read-only after recovery', () => {
  const read = readOoxmlPackage(
    docxFixture(
      '<w:p><w:ins w:id="4" w:author="Teacher"><w:r><w:t>First</w:t></w:r></w:ins><w:r><w:t> gap </w:t></w:r><w:ins w:id="4" w:author="Teacher"><w:r><w:t>Second</w:t></w:r></w:ins></w:p>'
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const result = restoreReviewState(read.package, { ...propertyContext, revision: 0 });
  if (!result.success) throw new Error(result.error);
  expect(result.state.groups).toHaveLength(2);
  expect(result.state.hasPendingReview).toBe(true);
  expect(result.state.reviewableGroupIds).toEqual([]);
  expect(result.state.groups.every((group) => group.readOnly)).toBe(true);
});

test('invalid identities and invalid persisted proposal chains fail without modifying the package', () => {
  let pkg = packageWithPropertyRecord(propertyRecordFixture(), {
    ...propertyRecord('one'),
    predecessors: ['two'],
  });
  pkg = packageWithPropertyRecord(pkg, { ...propertyRecord('two'), predecessors: ['one'] });
  const bytes = writeOoxmlPackage(pkg);
  expect(restoreReviewState(pkg, { ...propertyContext, revision: 0 })).toMatchObject({
    success: false,
    code: 'invalid-document',
    error: expect.stringContaining('cycle'),
  });
  for (const revision of [
    { documentId: '', epoch: 'load', revision: 0 },
    { documentId: 'document', epoch: '', revision: 0 },
    { documentId: 'document', epoch: 'load', revision: -1 },
    { documentId: 'document', epoch: 'load', revision: 0.5 },
  ])
    expect(restoreReviewState(pkg, revision)).toMatchObject({
      success: false,
      code: 'invalid-input',
    });
  expect(writeOoxmlPackage(pkg)).toEqual(bytes);
});
