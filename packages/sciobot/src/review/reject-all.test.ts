import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  collectRevisionSites,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { rejectAllReviewGroups } from './reject-all';
import { rejectReviewGroup } from './reject-group';
import { propertyEditor, propertyContext } from '../../test/support/property-execution';
import {
  groupFixture,
  groups,
  groupToken,
  proposeGroupText,
} from '../../test/support/review-groups';
import {
  crossStoryGroupFixture,
  groupDecisionFixture,
  proposeThreeProperties,
} from '../../test/support/group-decisions';
import { docxFixture } from '../../test/support/document';
import {
  rejectedMixedPackage,
  bulkSource,
  payloadSnapshot,
  proposalSnapshot,
  proposeMixedBulk,
  typeBulkSentence,
} from '../../test/support/bulk-decisions';

const word = (editor: ReturnType<typeof propertyEditor>) =>
  canonicalOoxmlFingerprint(editor.surface!.session.part());
test('bulk rejection resolves mixed native/property/atomic groups in one history frame with exact Word and payload values', async () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeMixedBulk(editor, () => typeBulkSentence(editor));
    const before = groups(editor),
      ids = before.map((group) => group.id),
      pendingWord = word(editor),
      pendingPayloads = payloadSnapshot(bulkSource(editor)),
      proposals = proposalSnapshot(bulkSource(editor)),
      expected = rejectedMixedPackage(bulkSource(editor));
    // The dog edit has two disjoint intervals; the cat edit has one.
    expect(before).toHaveLength(6);
    const token = groupToken(editor);
    expect(rejectAllReviewGroups(editor, token, propertyContext).success).toBe(true);
    expect(editor.surface!.session.bodyText()).toContain('Manual sentence. Keep tail');
    expect(groups(editor)).toEqual([]);
    expect(word(editor)).toBe(
      canonicalOoxmlFingerprint(expected.parts.get(expected.mainDocumentPart)!)
    );
    expect(payloadSnapshot(bulkSource(editor))).toEqual(payloadSnapshot(expected));
    expect(proposalSnapshot(bulkSource(editor))).toEqual(proposals);
    expect(collectRevisionSites(editor.surface!.session.part())).toEqual([]);
    const resolved = word(editor),
      revision = editor.surface!.session.packageRevision();
    expect(rejectAllReviewGroups(editor, token, propertyContext)).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    expect(rejectAllReviewGroups(editor, groupToken(editor), propertyContext).success).toBe(true);
    expect(editor.surface!.session.packageRevision()).toBe(revision);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(payloadSnapshot(reopened.package)).toEqual(payloadSnapshot(expected));
    expect(
      canonicalOoxmlFingerprint(reopened.package.parts.get(reopened.package.mainDocumentPart)!)
    ).toBe(canonicalOoxmlFingerprint(expected.parts.get(expected.mainDocumentPart)!));
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor).map((group) => group.id)).toEqual(ids);
    expect(word(editor)).toBe(pendingWord);
    expect(editor.surface!.session.bodyText()).toContain('Manual sentence. Keep tail');
    expect(payloadSnapshot(bulkSource(editor))).toEqual(pendingPayloads);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(groups(editor)).toEqual([]);
    expect(word(editor)).toBe(resolved);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor)).toHaveLength(4);
    expect(
      editor
        .surface!.session.paragraphIds()
        .some(
          (id) =>
            paragraphTextOf(editor.surface!.session.part(), id) === 'Manual sentence. Keep tail'
        )
    ).toBe(true);
  } finally {
    editor.destroy();
  }
});

for (const metadata of [false, true])
  test(
    'bulk rejection keeps complete multi-block proposal data and metadata-only fallback; metadata=' +
      metadata,
    () => {
      const editor = propertyEditor(groupDecisionFixture());
      try {
        proposeThreeProperties(editor, metadata);
        const ids = groups(editor).map((group) => group.id),
          raw = proposalSnapshot(bulkSource(editor)),
          before = word(editor);
        expect(rejectAllReviewGroups(editor, groupToken(editor), propertyContext).success).toBe(
          true
        );
        expect(groups(editor)).toEqual([]);
        for (const id of ['first', 'middle', 'last'])
          expect(buildBlockIndex(bulkSource(editor)).byId.get(id)!.payload.attributes.name).toBe(
            'Old ' + id
          );
        expect(proposalSnapshot(bulkSource(editor))).toEqual(raw);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(groups(editor).map((group) => group.id)).toEqual(ids);
        expect(word(editor)).toBe(before);
      } finally {
        editor.destroy();
      }
    }
  );

test('bulk rejection includes body/header groups and independent imported groups, preserving all identities on one undo', async () => {
  const editor = propertyEditor(crossStoryGroupFixture());
  try {
    const ids = groups(editor).map((group) => group.id),
      parts = [...bulkSource(editor).parts.values()].filter(
        (part) => collectRevisionSites(part).length
      ),
      expected = parts.map((part) => {
        const rejected = applyTreeOp(part, { op: 'rejectAllRevisions' });
        if (!rejected.ok) throw new Error(rejected.reason);
        return [part.name, canonicalOoxmlFingerprint(rejected.part)];
      }),
      before = parts.map((part) => [part.name, canonicalOoxmlFingerprint(part)]);
    expect(ids).toHaveLength(2);
    expect(rejectAllReviewGroups(editor, groupToken(editor), propertyContext).success).toBe(true);
    expect(groups(editor)).toEqual([]);
    expect(
      parts.map((part) => [
        part.name,
        canonicalOoxmlFingerprint(bulkSource(editor).parts.get(part.name)!),
      ])
    ).toEqual(expected);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor).map((group) => group.id)).toEqual(ids);
    expect(
      parts.map((part) => [
        part.name,
        canonicalOoxmlFingerprint(bulkSource(editor).parts.get(part.name)!),
      ])
    ).toEqual(before);
    const read = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!read.ok) throw new Error(read.reason);
    const reopened = propertyEditor(read.package);
    try {
      expect(groups(reopened).map((group) => group.id)).toEqual(ids);
    } finally {
      reopened.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('empty and repeated bulk rejection add no history entry', () => {
  const editor = propertyEditor(groupFixture());
  try {
    const revision = editor.surface!.session.packageRevision();
    expect(editor.surface!.session.canUndo()).toBe(false);
    expect(rejectAllReviewGroups(editor, groupToken(editor), propertyContext).success).toBe(true);
    expect(editor.surface!.session.packageRevision()).toBe(revision);
    expect(editor.surface!.session.canUndo()).toBe(false);
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    const ids = groups(editor).map((group) => group.id);
    expect(rejectAllReviewGroups(editor, groupToken(editor), propertyContext).success).toBe(true);
    expect(rejectAllReviewGroups(editor, groupToken(editor), propertyContext).success).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor).map((group) => group.id)).toEqual(ids);
  } finally {
    editor.destroy();
  }
});

test('one read-only imported group prevents rejection of every otherwise valid group', () => {
  const loaded = readOoxmlPackage(
    docxFixture(
      '<w:p><w:ins w:id="1" w:author="Good"><w:r><w:t>Good</w:t></w:r></w:ins></w:p><w:p><w:ins w:id="4" w:author="Teacher"><w:r><w:t>First</w:t></w:r></w:ins><w:r><w:t> gap </w:t></w:r><w:ins w:id="4" w:author="Teacher"><w:r><w:t>Second</w:t></w:r></w:ins></w:p>'
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const editor = propertyEditor(loaded.package);
  try {
    const before = writeOoxmlPackage(bulkSource(editor)),
      ids = groups(editor).map((group) => group.id),
      revision = editor.surface!.session.packageRevision();
    expect(groups(editor).some((group) => !group.readOnly)).toBe(true);
    expect(groups(editor).some((group) => group.readOnly)).toBe(true);
    expect(rejectAllReviewGroups(editor, groupToken(editor), propertyContext).success).toBe(false);
    expect(groups(editor).map((group) => group.id)).toEqual(ids);
    expect(writeOoxmlPackage(bulkSource(editor))).toEqual(before);
    expect(editor.surface!.session.packageRevision()).toBe(revision);
    expect(editor.surface!.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

test('a late locked payload refuses the whole batch and releases the guard for the next valid action', () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeMixedBulk(editor);
    const ids = groups(editor).map((group) => group.id),
      control = buildBlockIndex(bulkSource(editor)).byId.get('activity')!.controlId;
    editor.surface!.applyAutomationOps(() => [
      { op: 'setContentControlProperties', controlId: control, lock: 'contentLocked' },
    ]);
    const before = writeOoxmlPackage(bulkSource(editor)),
      revision = editor.surface!.session.packageRevision();
    expect(rejectAllReviewGroups(editor, groupToken(editor), propertyContext).success).toBe(false);
    expect(writeOoxmlPackage(bulkSource(editor))).toEqual(before);
    expect(editor.surface!.session.packageRevision()).toBe(revision);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor).map((group) => group.id)).toEqual(ids);
    expect(rejectAllReviewGroups(editor, groupToken(editor), propertyContext).success).toBe(true);
  } finally {
    editor.destroy();
  }
});

test('a second decision from a synchronous change listener refuses while bulk rejection runs', () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeMixedBulk(editor);
    const ids = groups(editor).map((group) => group.id),
      nested: unknown[] = [];
    const unsubscribe = editor.on('change', () => {
      if (nested.length) return;
      nested.push(rejectAllReviewGroups(editor, groupToken(editor), propertyContext));
      nested.push(rejectReviewGroup(editor, ids[0]!, groupToken(editor), propertyContext));
    });
    try {
      expect(rejectAllReviewGroups(editor, groupToken(editor), propertyContext).success).toBe(true);
    } finally {
      unsubscribe();
    }
    expect(nested).toHaveLength(2);
    for (const result of nested)
      expect(result).toMatchObject({
        success: false,
        error: expect.stringContaining('already running'),
      });
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor).map((group) => group.id)).toEqual(ids);
  } finally {
    editor.destroy();
  }
});

test('document switches with an equal numeric revision still refuse the previous document and epoch', async () => {
  const editor = propertyEditor(crossStoryGroupFixture());
  try {
    const old = groupToken(editor);
    await editor.load(writeOoxmlPackage(crossStoryGroupFixture()));
    const identity = { documentId: 'another-document', epoch: 'another-mount' },
      current = editor.surface!.session.packageRevision(),
      before = writeOoxmlPackage(bulkSource(editor));
    expect(current).toBe(old.revision);
    expect(rejectAllReviewGroups(editor, old, identity)).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    expect(
      rejectAllReviewGroups(editor, { ...old, documentId: identity.documentId }, identity)
    ).toMatchObject({ success: false, code: 'stale-revision' });
    expect(writeOoxmlPackage(bulkSource(editor))).toEqual(before);
    expect(editor.surface!.session.canUndo()).toBe(false);
    expect(
      rejectAllReviewGroups(editor, { ...identity, revision: current }, identity).success
    ).toBe(true);
  } finally {
    editor.destroy();
  }
});

test('buffered typing, viewing, and a missing review module refuse before bulk mutation', () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeMixedBulk(editor);
    const ids = groups(editor).map((group) => group.id),
      token = groupToken(editor),
      surface = editor.surface!,
      id = surface.session.paragraphIds().at(-1)!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    surface.insertPlainText('Manual: ');
    expect(rejectAllReviewGroups(editor, token, propertyContext)).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    expect(groups(editor).map((group) => group.id)).toEqual(ids);
    const before = writeOoxmlPackage(bulkSource(editor));
    surface.setEditingMode('view');
    expect(rejectAllReviewGroups(editor, groupToken(editor), propertyContext).success).toBe(false);
    expect(writeOoxmlPackage(bulkSource(editor))).toEqual(before);
    surface.setEditingMode('edit');
    const bare = propertyEditor(bulkSource(editor), false);
    try {
      expect(rejectAllReviewGroups(bare, groupToken(bare), propertyContext).success).toBe(false);
    } finally {
      bare.destroy();
    }
  } finally {
    editor.destroy();
  }
});
