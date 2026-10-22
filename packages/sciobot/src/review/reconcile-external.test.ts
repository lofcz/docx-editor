import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import {
  atomicListening,
  atomicVisual,
  proposedListening,
  proposedVisual,
} from '../../test/support/atomic-execution';
import {
  bulkSource,
  payloadSnapshot,
  proposalSnapshot,
  proposeMixedBulk,
} from '../../test/support/bulk-decisions';
import { externalReviewDocument } from '../../test/support/external-review';
import { groupFixture, groupToken, groups } from '../../test/support/review-groups';
import {
  executeProperties,
  propertyContext,
  propertyEditor,
} from '../../test/support/property-execution';
import { acceptAllReviewGroups } from './accept-all';
import { inspectExternalReview, reconcileExternalReview } from './reconcile-external';
import { restoreReviewState } from './restore';
import oracle from '../__tests__/fixtures/domain-properties.json';
import type { JsonObject } from '../data/json';
import {
  propertyRecordFixture,
  propertyRecord,
  packageWithPropertyRecord,
} from '../../test/support/property-record';
import { propertyLabel } from '../../test/support/property-execution';
import { decidePropertyProposal } from './decide-properties';
import { decideAtomicProposal } from './track-atomic';

async function external(choices: Parameters<typeof externalReviewDocument>[1]) {
  const original = propertyEditor(groupFixture());
  try {
    proposeMixedBulk(original);
    const bytes = externalReviewDocument(new Uint8Array(await original.save()), choices),
      read = readOoxmlPackage(bytes);
    if (!read.ok) throw new Error(read.reason);
    return propertyEditor(read.package);
  } finally {
    original.destroy();
  }
}
const word = (editor: ReturnType<typeof propertyEditor>) =>
  canonicalOoxmlFingerprint(editor.surface!.session.part());
function report(editor: ReturnType<typeof propertyEditor>) {
  const result = inspectExternalReview(bulkSource(editor));
  if (!result.success) throw new Error(result.error);
  return result.report;
}

for (const decision of ['accept', 'reject'] as const)
  test(
    'external WordprocessingML ' +
      decision +
      ' selects exact property/atomic values without rewriting Word; undo and reopen',
    async () => {
      const editor = await external(
        Object.fromEntries(['activity', 'visual', 'audio'].map((id) => [id, { decision }]))
      );
      try {
        const beforeWord = word(editor),
          beforePayloads = payloadSnapshot(bulkSource(editor)),
          raw = proposalSnapshot(bulkSource(editor)),
          ids = groups(editor).map((group) => group.id),
          before = groupToken(editor);
        expect(report(editor).entries.map((entry) => entry.status)).toEqual(
          Array(3).fill(decision === 'accept' ? 'accepted' : 'rejected')
        );
        expect(editor.surface!.session.canUndo()).toBe(false);
        const result = reconcileExternalReview(editor, before, propertyContext);
        if (!result.success) throw new Error(result.error);
        expect(result.reconciled).toHaveLength(3);
        expect(result.report.hasConflicts).toBe(false);
        expect(word(editor)).toBe(beforeWord);
        expect(proposalSnapshot(bulkSource(editor))).toEqual(raw);
        const index = buildBlockIndex(bulkSource(editor));
        expect(index.byId.get('activity')!.payload.attributes).toEqual({
          name: decision === 'accept' ? 'New activity' : 'Old activity',
          duration: decision === 'accept' ? 15 : 10,
          resource: 'keep',
        });
        expect(index.byId.get('visual')!.payload).toEqual(
          decision === 'accept' ? proposedVisual : atomicVisual
        );
        expect(index.byId.get('audio')!.payload).toEqual(
          decision === 'accept' ? proposedListening : atomicListening
        );
        expect(groups(editor)).toHaveLength(3);
        expect(report(editor).entries.every((entry) => entry.status === 'resolved')).toBe(true);
        const revision = editor.surface!.session.packageRevision();
        expect(reconcileExternalReview(editor, groupToken(editor), propertyContext)).toMatchObject({
          success: true,
          reconciled: [],
        });
        expect(editor.surface!.session.packageRevision()).toBe(revision);
        expect(reconcileExternalReview(editor, before, propertyContext)).toMatchObject({
          success: false,
          code: 'stale-revision',
        });
        const read = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!read.ok) throw new Error(read.reason);
        const reopened = propertyEditor(read.package);
        try {
          expect(report(reopened).entries.every((entry) => entry.status === 'resolved')).toBe(true);
          expect(payloadSnapshot(bulkSource(reopened))).toEqual(
            payloadSnapshot(bulkSource(editor))
          );
          expect(word(reopened)).toBe(beforeWord);
        } finally {
          reopened.destroy();
        }
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(word(editor)).toBe(beforeWord);
        expect(payloadSnapshot(bulkSource(editor))).toEqual(beforePayloads);
        expect(groups(editor).map((group) => group.id)).toEqual(ids);
        expect(
          report(editor).entries.every(
            (entry) => entry.status === (decision === 'accept' ? 'accepted' : 'rejected')
          )
        ).toBe(true);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(word(editor)).toBe(beforeWord);
        expect(report(editor).entries.every((entry) => entry.status === 'resolved')).toBe(true);
      } finally {
        editor.destroy();
      }
    }
  );

test('mixed external decisions reconcile independent groups and report an edited fallback without rolling it back', async () => {
  const editor = await external({
    activity: { decision: 'accept' },
    visual: { decision: 'reject' },
    audio: { decision: 'accept', append: ' User correction.' },
  });
  try {
    const before = word(editor),
      raw = proposalSnapshot(bulkSource(editor));
    expect(report(editor)).toMatchObject({
      hasConflicts: true,
      entries: expect.arrayContaining([
        {
          proposalId: 'atomic-proposal',
          blockId: 'audio',
          status: 'conflict',
          conflict: 'changed-fallback',
        },
      ]),
    });
    const result = reconcileExternalReview(editor, groupToken(editor), propertyContext);
    if (!result.success) throw new Error(result.error);
    expect(result.reconciled).toHaveLength(2);
    expect(result.report.hasConflicts).toBe(true);
    expect(word(editor)).toBe(before);
    expect(editor.surface!.session.bodyText()).toContain('New audio fallback User correction.');
    expect(buildBlockIndex(bulkSource(editor)).byId.get('audio')!.payload).toEqual(atomicListening);
    expect(buildBlockIndex(bulkSource(editor)).byId.get('activity')!.payload.attributes.name).toBe(
      'New activity'
    );
    expect(proposalSnapshot(bulkSource(editor))).toEqual(raw);
    const restored = restoreReviewState(bulkSource(editor), groupToken(editor));
    if (!restored.success) throw new Error(restored.error);
    expect(
      restored.state.groups.some(
        (group) => group.properties.some((member) => member.blockId === 'audio') && group.readOnly
      )
    ).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(word(editor)).toBe(before);
    expect(buildBlockIndex(bulkSource(editor)).byId.get('activity')!.payload.attributes.name).toBe(
      'Old activity'
    );
  } finally {
    editor.destroy();
  }
});

test('partial external revision decisions remain conflicts and keep exact native content and metadata', async () => {
  const editor = await external({ activity: { decision: 'accept', only: 'ins' } });
  try {
    const bytes = writeOoxmlPackage(bulkSource(editor)),
      before = word(editor),
      revision = editor.surface!.session.packageRevision();
    expect(report(editor).entries.find((entry) => entry.blockId === 'activity')).toMatchObject({
      status: 'conflict',
      conflict: 'changed-fallback',
    });
    expect(reconcileExternalReview(editor, groupToken(editor), propertyContext)).toMatchObject({
      success: true,
      reconciled: [],
      report: { hasConflicts: true },
    });
    expect(word(editor)).toBe(before);
    expect(writeOoxmlPackage(bulkSource(editor))).toEqual(bytes);
    expect(editor.surface!.session.packageRevision()).toBe(revision);
    expect(editor.surface!.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

test('unchanged metadata-only fallback cannot prove a decision in Word and creates no history', () => {
  const editor = propertyEditor();
  try {
    expect(
      executeProperties(editor, { set: { duration: 15 } }, 'Old activity — 10 minutes').success
    ).toBe(true);
    const before = writeOoxmlPackage(bulkSource(editor)),
      revision = editor.surface!.session.packageRevision();
    expect(report(editor).entries).toMatchObject([{ status: 'pending' }]);
    expect(reconcileExternalReview(editor, groupToken(editor), propertyContext)).toMatchObject({
      success: true,
      reconciled: [],
      report: { hasConflicts: false },
    });
    expect(editor.surface!.session.packageRevision()).toBe(revision);
    expect(writeOoxmlPackage(bulkSource(editor))).toEqual(before);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(report(editor).entries).toEqual([]);
  } finally {
    editor.destroy();
  }
});

test('changed committed attributes conflict while unrelated current payload data survives clear reconciliation', async () => {
  const editor = await external({ activity: { decision: 'accept' } });
  try {
    const owner = buildBlockIndex(bulkSource(editor)).byId.get('activity')!;
    const changed = withPayloadRecord(
      bulkSource(editor),
      {
        ...owner.payload,
        attributes: { ...owner.payload.attributes, resource: 'Manual resource' },
        extension: { manual: true },
      },
      'Stored activity label'
    );
    if (!changed.ok) throw new Error(changed.reason);
    const fresh = propertyEditor(changed.pkg);
    try {
      expect(reconcileExternalReview(fresh, groupToken(fresh), propertyContext)).toMatchObject({
        success: true,
        reconciled: [{ blockId: 'activity' }],
      });
      expect(buildBlockIndex(bulkSource(fresh)).byId.get('activity')!.payload).toMatchObject({
        attributes: { name: 'New activity', resource: 'Manual resource' },
        extension: { manual: true },
      });
    } finally {
      fresh.destroy();
    }
    const conflict = withPayloadRecord(
      bulkSource(editor),
      { ...owner.payload, attributes: { ...owner.payload.attributes, name: 'Manual name' } },
      'Stored activity label'
    );
    if (!conflict.ok) throw new Error(conflict.reason);
    const stale = propertyEditor(conflict.pkg);
    try {
      const before = writeOoxmlPackage(bulkSource(stale));
      expect(report(stale).entries.find((entry) => entry.blockId === 'activity')).toMatchObject({
        status: 'conflict',
        conflict: 'changed-committed-value',
      });
      expect(reconcileExternalReview(stale, groupToken(stale), propertyContext)).toMatchObject({
        success: true,
        reconciled: [],
      });
      expect(writeOoxmlPackage(bulkSource(stale))).toEqual(before);
    } finally {
      stale.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('external reconciliation shares the complete group-decision reentry guard', async () => {
  const editor = await external({ activity: { decision: 'accept' } });
  try {
    const nested: unknown[] = [];
    const stop = editor.on('change', () => {
      if (nested.length) return;
      nested.push(reconcileExternalReview(editor, groupToken(editor), propertyContext));
      nested.push(acceptAllReviewGroups(editor, groupToken(editor), propertyContext));
      nested.push(
        decidePropertyProposal(
          editor,
          'property-proposal',
          'accept',
          groupToken(editor),
          propertyContext
        )
      );
      nested.push(
        decideAtomicProposal(
          editor,
          'atomic-proposal',
          'accept',
          groupToken(editor),
          propertyContext
        )
      );
    });
    try {
      expect(reconcileExternalReview(editor, groupToken(editor), propertyContext).success).toBe(
        true
      );
    } finally {
      stop();
    }
    expect(nested).toHaveLength(4);
    for (const result of nested)
      expect(result).toMatchObject({
        success: false,
        error: expect.stringContaining('already running'),
      });
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(reconcileExternalReview(editor, groupToken(editor), propertyContext).success).toBe(true);
  } finally {
    editor.destroy();
  }
});

for (const item of oracle.cases)
  for (const decision of ['accept', 'reject'] as const)
    test(
      'external review retains captured property presence ' + item.name + '/' + decision,
      async () => {
        const pkg = propertyRecordFixture(),
          owner = buildBlockIndex(pkg).byId.get('activity')!,
          changed = withPayloadRecord(
            pkg,
            { ...owner.payload, attributes: item.before as JsonObject },
            'Original stored label'
          );
        if (!changed.ok) throw new Error(changed.reason);
        const original = propertyEditor(changed.pkg);
        let editor: ReturnType<typeof propertyEditor> | undefined;
        try {
          const proposed = executeProperties(original, {
            set: Object.fromEntries(
              item.keys
                .filter((key) => Object.hasOwn(item.after, key))
                .map((key) => [key, (item.after as JsonObject)[key]!])
            ),
            remove: item.keys.filter((key) => !Object.hasOwn(item.after, key)),
          });
          if (!proposed.success) throw new Error(proposed.error);
          const read = readOoxmlPackage(
            externalReviewDocument(new Uint8Array(await original.save()), {
              activity: { decision },
            })
          );
          if (!read.ok) throw new Error(read.reason);
          editor = propertyEditor(read.package);
          const before = word(editor);
          expect(
            reconcileExternalReview(editor, groupToken(editor), propertyContext)
          ).toMatchObject({ success: true, reconciled: [{ blockId: 'activity', decision }] });
          expect(
            buildBlockIndex(bulkSource(editor)).byId.get('activity')!.payload.attributes
          ).toEqual(
            decision === 'accept' ? (item.working as JsonObject) : (item.committed as JsonObject)
          );
          expect(word(editor)).toBe(before);
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(
            buildBlockIndex(bulkSource(editor)).byId.get('activity')!.payload.attributes
          ).toEqual(item.before as JsonObject);
          expect(word(editor)).toBe(before);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        } finally {
          original.destroy();
          editor?.destroy();
        }
      }
    );

test('stale epochs, numeric revisions, viewing and missing review support refuse before reconciliation', async () => {
  const editor = await external({ activity: { decision: 'accept' } });
  try {
    const token = groupToken(editor),
      bytes = writeOoxmlPackage(bulkSource(editor));
    for (const [expected, identity] of [
      [{ ...token, revision: token.revision + 1 }, propertyContext],
      [token, { ...propertyContext, epoch: 'another-mount' }],
      [token, { ...propertyContext, documentId: 'another-document' }],
    ] as const)
      expect(reconcileExternalReview(editor, expected, identity)).toMatchObject({
        success: false,
        code: 'stale-revision',
      });
    editor.surface!.setEditingMode('view');
    expect(reconcileExternalReview(editor, groupToken(editor), propertyContext).success).toBe(
      false
    );
    expect(writeOoxmlPackage(bulkSource(editor))).toEqual(bytes);
    editor.surface!.setEditingMode('edit');
    const bare = propertyEditor(bulkSource(editor), false);
    try {
      expect(reconcileExternalReview(bare, groupToken(bare), propertyContext).success).toBe(false);
    } finally {
      bare.destroy();
    }
    expect(editor.surface!.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

test('unseen native input invalidates the token and a later fresh read preserves that fallback as a conflict', async () => {
  const editor = await external({ activity: { decision: 'accept' } });
  try {
    const token = groupToken(editor),
      paragraphId = propertyLabel(editor).id;
    const offset = editor.surface!.session.paragraphIds().includes(paragraphId)
      ? 'New activity — 15 minutes'.length
      : -1;
    expect(offset).toBeGreaterThan(0);
    editor.surface!.setSelection({
      anchor: { paragraphId, offset },
      head: { paragraphId, offset },
    });
    editor.surface!.insertPlainText(' User edit');
    expect(reconcileExternalReview(editor, token, propertyContext)).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    const before = word(editor);
    expect(reconcileExternalReview(editor, groupToken(editor), propertyContext)).toMatchObject({
      success: true,
      reconciled: [],
      report: { hasConflicts: true },
    });
    expect(word(editor)).toBe(before);
    expect(editor.surface!.session.bodyText()).toContain('New activity — 15 minutes User edit');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(report(editor).entries.find((entry) => entry.blockId === 'activity')).toMatchObject({
      status: 'accepted',
    });
  } finally {
    editor.destroy();
  }
});

test('missing or locked owners remain conflicts while native Word content stays intact', async () => {
  for (const remove of [false, true]) {
    const editor = await external({ activity: { decision: 'accept' } });
    try {
      const controlId = buildBlockIndex(bulkSource(editor)).byId.get('activity')!.controlId;
      expect(
        editor.surface!.applyAutomationOps(() =>
          remove
            ? [{ op: 'removeContentControl', controlId, keepContent: false }]
            : [{ op: 'setContentControlProperties', controlId, lock: 'contentLocked' }]
        ).committed
      ).toBe(true);
      const before = writeOoxmlPackage(bulkSource(editor)),
        beforeWord = word(editor);
      expect(report(editor).entries.find((entry) => entry.blockId === 'activity')).toMatchObject({
        status: 'conflict',
        conflict: remove ? 'missing-owner' : 'restricted-owner',
      });
      expect(reconcileExternalReview(editor, groupToken(editor), propertyContext)).toMatchObject({
        success: true,
        reconciled: [],
        report: { hasConflicts: true },
      });
      expect(writeOoxmlPackage(bulkSource(editor))).toEqual(before);
      expect(word(editor)).toBe(beforeWord);
    } finally {
      editor.destroy();
    }
  }
});

test('cyclic persisted proposal causality refuses without writes', () => {
  let pkg = propertyRecordFixture();
  for (const [id, predecessor] of [
    ['cycle-one', 'cycle-two'],
    ['cycle-two', 'cycle-one'],
  ])
    pkg = packageWithPropertyRecord(pkg, { ...propertyRecord(id), predecessors: [predecessor!] });
  const editor = propertyEditor(pkg);
  try {
    const before = writeOoxmlPackage(bulkSource(editor));
    expect(inspectExternalReview(bulkSource(editor))).toMatchObject({
      success: false,
      code: 'invalid-document',
      error: expect.stringContaining('cycle'),
    });
    expect(reconcileExternalReview(editor, groupToken(editor), propertyContext)).toMatchObject({
      success: false,
      code: 'invalid-document',
    });
    expect(writeOoxmlPackage(bulkSource(editor))).toEqual(before);
    expect(editor.surface!.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});
