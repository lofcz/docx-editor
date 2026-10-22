import { crossStoryGroupFixture } from '../../test/support/group-decisions';
import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { acceptReviewGroup } from './accept-group';
import { readStoredPropertyProposals } from './stored-property-proposals';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { readPropertyReviewStore } from './property-record-store';
import { readReviewGroups } from './groups';
import { decidePropertyProposal } from './decide-properties';
import { decideAtomicProposal } from './track-atomic';
import { reviewGroupRevisionReservations } from './group-allocation';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { docxFixture } from '../../test/support/document';
import { propertyEditor, propertyContext } from '../../test/support/property-execution';
import {
  atomicFixture,
  executeAtomic,
  proposedListening,
  proposedVisual,
} from '../../test/support/atomic-execution';
import {
  groupFixture,
  groups,
  groupToken,
  proposeGroupText,
} from '../../test/support/review-groups';
import {
  acceptBlockGroup,
  blockLabel,
  decisionBlocks,
  groupDecisionFixture,
  proposeThreeProperties,
} from '../../test/support/group-decisions';

for (const metadataMiddle of [false, true])
  test(`accepting property group two retains groups one and three, immutable proposal JSON, projections, and one history unit; metadata=${metadataMiddle}`, async () => {
    const editor = propertyEditor(groupDecisionFixture());
    try {
      const proposal = proposeThreeProperties(editor, metadataMiddle).proposal!,
        before = groups(editor).map((group) => group.id),
        labelsBefore = new Map(decisionBlocks.map((id) => [id, blockLabel(editor, id)])),
        part = canonicalOoxmlFingerprint(editor.surface!.session.part()),
        store = readPropertyReviewStore(editor.surface!.session.currentPackage());
      if (store.status !== 'ready') throw new Error('Missing record');
      const raw = store.records.get(proposal.id)!.raw;
      expect(before).toHaveLength(3);
      expect(groups(editor).every((group) => !group.readOnly)).toBe(true);
      expect(acceptBlockGroup(editor).success).toBe(true);
      expect(groups(editor).map((group) => group.id)).toEqual([before[0]!, before[2]!]);
      const index = buildBlockIndex(editor.surface!.session.currentPackage());
      for (const id of decisionBlocks) {
        expect(index.byId.get(id)!.payload.attributes.name).toBe(
          id === 'middle' ? 'New middle' : `Old ${id}`
        );
        expect(blockLabel(editor, id)).toBe(
          id === 'middle' ? (metadataMiddle ? 'Old middle' : 'New middle') : labelsBefore.get(id)!
        );
      }
      const record = readPropertyReviewStore(editor.surface!.session.currentPackage());
      expect(record.status === 'ready' && record.records.get(proposal.id)!.raw).toBe(raw);
      const pending = readStoredPropertyProposals(
        editor.surface!.session.currentPackage(),
        groupToken(editor)
      );
      if (!pending.success) throw new Error(pending.error);
      expect(pending.pending[0]!.plan.operations).toHaveLength(4);
      const [committed, working] = await Promise.all([
        readCommittedView(editor.surface!.session, groupToken(editor)),
        readWorkingView(editor.surface!.session, groupToken(editor)),
      ]);
      if (!committed.success || !working.success) throw new Error('Missing views');
      expect(committed.view.projection.text).toContain('name="New middle"');
      expect(committed.view.projection.text).toContain('name="Old first"');
      expect(working.view.projection.text).toContain('name="New first"');
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      expect(readReviewGroups(reopened.package)).toMatchObject({
        success: true,
        groups: [{ id: before[0] }, { id: before[2] }],
      });
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(groups(editor).map((group) => group.id)).toEqual(before);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(part);
      expect(
        buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('middle')!.payload
          .attributes.name
      ).toBe('Old middle');
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(groups(editor).map((group) => group.id)).toEqual([before[0]!, before[2]!]);
    } finally {
      editor.destroy();
    }
  });

for (const remainingDecision of ['accept', 'reject'] as const)
  test(`legacy whole-property ${remainingDecision} after one accepted group touches only the remaining blocks`, () => {
    const editor = propertyEditor(groupDecisionFixture());
    try {
      proposeThreeProperties(editor);
      expect(acceptBlockGroup(editor).success).toBe(true);
      const pendingIds = groups(editor).map((group) => group.id);
      expect(
        decidePropertyProposal(
          editor,
          'three-blocks',
          remainingDecision,
          groupToken(editor),
          propertyContext
        ).success
      ).toBe(true);
      expect(groups(editor)).toEqual([]);
      const index = buildBlockIndex(editor.surface!.session.currentPackage());
      expect(index.byId.get('middle')!.payload.attributes.name).toBe('New middle');
      expect(index.byId.get('first')!.payload.attributes.name).toBe(
        remainingDecision === 'accept' ? 'New first' : 'Old first'
      );
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(groups(editor).map((group) => group.id)).toEqual(pendingIds);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(groups(editor)).toHaveLength(3);
    } finally {
      editor.destroy();
    }
  });

test('complete atomic groups retain original assets, accept one payload only, and allow the legacy decision on the other', () => {
  const editor = propertyEditor(atomicFixture());
  try {
    expect(executeAtomic(editor, [proposedVisual, proposedListening]).success).toBe(true);
    const first = groups(editor),
      group = first.find((group) => group.properties[0]?.blockId === 'visual')!;
    expect(acceptReviewGroup(editor, group.id, groupToken(editor), propertyContext).success).toBe(
      true
    );
    const index = buildBlockIndex(editor.surface!.session.currentPackage());
    expect(index.byId.get('visual')!.payload).toEqual(proposedVisual);
    expect(index.byId.get('audio')!.payload).not.toEqual(proposedListening);
    expect(groups(editor)).toHaveLength(1);
    expect(
      decideAtomicProposal(editor, 'atomic-proposal', 'reject', groupToken(editor), propertyContext)
        .success
    ).toBe(true);
    expect(groups(editor)).toEqual([]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor)).toHaveLength(1);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor).map((group) => group.id)).toEqual(first.map((group) => group.id));
  } finally {
    editor.destroy();
  }
});

test('native group two resolves without other changes, and undo restores its stable identity', () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    proposeGroupText(editor, 'A dog ran.', 'A dog slept.');
    proposeGroupText(editor, 'Keep tail', 'Keep tail!');
    const before = groups(editor),
      fingerprint = canonicalOoxmlFingerprint(editor.surface!.session.part());
    expect(before).toHaveLength(3);
    expect(
      acceptReviewGroup(editor, before[1]!.id, groupToken(editor), propertyContext).success
    ).toBe(true);
    expect(groups(editor).map((group) => group.id)).toEqual([before[0]!.id, before[2]!.id]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(fingerprint);
    expect(groups(editor).map((group) => group.id)).toEqual(before.map((group) => group.id));
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  } finally {
    editor.destroy();
  }
});

test('all accepted blocks leave no pending projection and reserve fallback IDs after legacy removal', () => {
  const editor = propertyEditor(groupDecisionFixture());
  try {
    const proposal = proposeThreeProperties(editor).proposal!;
    for (const id of decisionBlocks) expect(acceptBlockGroup(editor, id).success).toBe(true);
    expect(groups(editor)).toEqual([]);
    expect(
      readStoredPropertyProposals(editor.surface!.session.currentPackage(), groupToken(editor))
    ).toEqual({ success: true, pending: [] });
    const reserved = reviewGroupRevisionReservations(
      editor.surface!.session.currentPackage()
    ).flatMap((row) => row.ids);
    expect(proposal.nativeRevisions.every((row) => reserved.includes(row.id))).toBe(true);
    for (let i = 0; i < 3; i++) expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor)).toHaveLength(3);
    expect(
      decidePropertyProposal(editor, proposal.id, 'accept', groupToken(editor), propertyContext)
        .success
    ).toBe(true);
    expect(
      reviewGroupRevisionReservations(editor.surface!.session.currentPackage()).flatMap(
        (row) => row.ids
      )
    ).toEqual(reserved);
  } finally {
    editor.destroy();
  }
});

test('stale, viewing, missing-module, wrong-owner and locked group decisions refuse atomically', () => {
  const editor = propertyEditor(groupDecisionFixture());
  try {
    proposeThreeProperties(editor);
    const group = groups(editor)[1]!,
      token = groupToken(editor),
      before = writeOoxmlPackage(editor.surface!.session.currentPackage());
    for (const [id, expected, identity] of [
      [group.id, { ...token, revision: token.revision - 1 }, propertyContext],
      ['missing', token, propertyContext],
      [group.id, token, { ...propertyContext, epoch: 'other' }],
    ] as const)
      expect(acceptReviewGroup(editor, id, expected, identity).success).toBe(false);
    editor.surface!.setEditingMode('view');
    expect(acceptReviewGroup(editor, group.id, token, propertyContext).success).toBe(false);
    expect(writeOoxmlPackage(editor.surface!.session.currentPackage())).toEqual(before);
    editor.surface!.setEditingMode('edit');
    const owner = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('middle')!;
    editor.surface!.applyAutomationOps(
      () => [
        { op: 'setContentControlProperties', controlId: owner.controlId, lock: 'contentLocked' },
      ],
      { kind: 'body' }
    );
    const locked = writeOoxmlPackage(editor.surface!.session.currentPackage());
    expect(acceptReviewGroup(editor, group.id, groupToken(editor), propertyContext).success).toBe(
      false
    );
    expect(writeOoxmlPackage(editor.surface!.session.currentPackage())).toEqual(locked);
    const noReview = propertyEditor(editor.surface!.session.currentPackage(), false);
    try {
      expect(
        acceptReviewGroup(noReview, group.id, groupToken(noReview), propertyContext).success
      ).toBe(false);
    } finally {
      noReview.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('one stored native group resolves body and header members atomically with one undo and exact save/reopen', async () => {
  const editor = propertyEditor(crossStoryGroupFixture());
  try {
    const before = new Map(
        [...editor.surface!.session.currentPackage().parts].map(([name, part]) => [
          name,
          canonicalOoxmlFingerprint(part),
        ])
      ),
      id = 'g1:stored:rg_cross_story';
    expect(groups(editor).find((group) => group.id === id)!.native).toHaveLength(2);
    expect(acceptReviewGroup(editor, id, groupToken(editor), propertyContext).success).toBe(true);
    expect(
      collectRevisionSites(editor.surface!.session.currentPackage().parts.get('/word/header1.xml')!)
    ).toEqual([]);
    expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(1);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    for (const name of ['/word/document.xml', '/word/header1.xml'])
      expect(
        canonicalOoxmlFingerprint(editor.surface!.session.currentPackage().parts.get(name)!)
      ).toBe(before.get(name)!);
    expect(groups(editor).find((group) => group.id === id)!.native).toHaveLength(2);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(collectRevisionSites(reopened.package.parts.get('/word/header1.xml')!)).toEqual([]);
    expect(collectRevisionSites(reopened.package.parts.get('/word/document.xml')!)).toHaveLength(1);
  } finally {
    editor.destroy();
  }
});

test('imported native acceptance retains Word addresses for later decision undo without adding application attributes', () => {
  const read = readOoxmlPackage(
    docxFixture(
      '<w:p><w:ins w:id="5" w:author="Teacher"><w:r><w:t>Pending</w:t></w:r></w:ins></w:p>'
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const editor = propertyEditor(read.package);
  try {
    const before = groups(editor)[0]!.id;
    expect(acceptReviewGroup(editor, before, groupToken(editor), propertyContext).success).toBe(
      true
    );
    expect(groups(editor)).toEqual([]);
    expect(reviewGroupRevisionReservations(editor.surface!.session.currentPackage())).toEqual([
      { partName: '/word/document.xml', ids: ['5'] },
    ]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor)[0]!.id).toBe(before);
  } finally {
    editor.destroy();
  }
});

test('buffered manual input makes an earlier decision token stale without removing that input', () => {
  const editor = propertyEditor(groupDecisionFixture());
  try {
    proposeThreeProperties(editor);
    const before = groups(editor).map((group) => group.id),
      token = groupToken(editor),
      surface = editor.surface!,
      id = surface.session.paragraphIds().at(-1)!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    surface.insertPlainText('Manual: ');
    expect(acceptReviewGroup(editor, before[1]!, token, propertyContext)).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    expect(groups(editor).map((group) => group.id)).toEqual(before);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor).map((group) => group.id)).toEqual(before);
    const store = readPayloadStore(surface.session.currentPackage());
    expect(store.status === 'ready' && store.entries.get('middle')!.label).toBe('Stored middle');
  } finally {
    editor.destroy();
  }
});
