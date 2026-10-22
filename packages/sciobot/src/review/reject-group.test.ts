import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import { docxFixture } from '../../test/support/document';
import { propertyEditor, propertyContext } from '../../test/support/property-execution';
import {
  atomicFixture,
  atomicVisual,
  atomicListening,
  executeAtomic,
  proposedVisual,
  proposedListening,
} from '../../test/support/atomic-execution';
import {
  blockLabel,
  crossStoryGroupFixture,
  decisionBlocks,
  groupDecisionFixture,
  proposeThreeProperties,
} from '../../test/support/group-decisions';
import {
  groupFixture,
  groups,
  groupToken,
  proposeGroupText,
} from '../../test/support/review-groups';
import { rejectReviewGroup } from './reject-group';
import { readReviewGroups } from './groups';
import { readPropertyReviewStore } from './property-record-store';
import { readStoredPropertyProposals } from './stored-property-proposals';
import { reviewGroupRevisionReservations } from './group-allocation';
import { decidePropertyProposal } from './decide-properties';
import { decideAtomicProposal } from './track-atomic';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { planTableFormat } from '../diff/table-format';
import { trackTableFormatPlan } from './track-table-format';
import { cellReviewDocument, cellScopeFor } from '../../test/support/table-review';
import { parseTargetMarkdown } from '../projection/parse-target';

function plainText(node: unknown): string {
  if (!node || typeof node !== 'object') return '';
  const value = node as { value?: unknown; children?: unknown };
  if (typeof value.value === 'string') return value.value;
  return Array.isArray(value.children) ? value.children.map(plainText).join('') : '';
}

const source = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
function payloads(editor: DocxEditorInstance) {
  const read = readPayloadStore(source(editor));
  if (read.status !== 'ready') throw new Error('Missing payload store');
  return [...read.entries].map(([id, entry]) => [id, entry.raw, entry.label]);
}
function rejectBlock(editor: DocxEditorInstance, blockId = 'middle') {
  const group = groups(editor).find((group) =>
    group.properties.some((member) => member.blockId === blockId)
  );
  if (!group) throw new Error('Missing group');
  return rejectReviewGroup(editor, group.id, groupToken(editor), propertyContext);
}

for (const metadata of [false, true])
  test(`reject group two preserves neighbors, immutable payloads, native labels, one undo unit and reopening; metadata=${metadata}`, async () => {
    const editor = propertyEditor(groupDecisionFixture());
    try {
      const proposal = proposeThreeProperties(editor, metadata).proposal!;
      const pending = groups(editor),
        ids = pending.map((group) => group.id);
      const initial = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const stored = payloads(editor),
        labels = decisionBlocks.map((id) => blockLabel(editor, id));
      const record = readPropertyReviewStore(source(editor));
      if (record.status !== 'ready') throw new Error('Missing proposal');
      const raw = record.records.get(proposal.id)!.raw;
      expect(ids).toHaveLength(3);
      expect(rejectBlock(editor).success).toBe(true);
      expect(groups(editor).map((group) => group.id)).toEqual([ids[0]!, ids[2]!]);
      expect(groups(editor).map((group) => group.native)).toEqual([
        pending[0]!.native,
        pending[2]!.native,
      ]);
      expect(payloads(editor)).toEqual(stored);
      for (const id of decisionBlocks) {
        expect(buildBlockIndex(source(editor)).byId.get(id)!.payload.attributes).toEqual({
          name: `Old ${id}`,
          duration: 10,
          resource: 'keep',
        });
        expect(blockLabel(editor, id)).toBe(
          id === 'middle' ? 'Old middle' : labels[decisionBlocks.indexOf(id)]!
        );
      }
      const after = readPropertyReviewStore(source(editor));
      expect(after.status === 'ready' && after.records.get(proposal.id)!.raw).toBe(raw);
      const remaining = readStoredPropertyProposals(source(editor), groupToken(editor));
      if (!remaining.success) throw new Error(remaining.error);
      expect(remaining.pending[0]!.plan.operations).toHaveLength(4);
      const [committed, working] = await Promise.all([
        readCommittedView(editor.surface!.session, groupToken(editor)),
        readWorkingView(editor.surface!.session, groupToken(editor)),
      ]);
      if (!committed.success || !working.success) throw new Error('Missing views');
      expect(committed.view.projection.text).toContain('name="Old middle"');
      expect(working.view.projection.text).toContain('name="Old middle"');
      expect(working.view.projection.text).toContain('name="New first"');
      expect(working.view.projection.text).toContain('name="New last"');
      expect(working.view.projection.text).not.toContain('name="New middle"');
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      expect(readReviewGroups(reopened.package)).toMatchObject({
        success: true,
        groups: [{ id: ids[0] }, { id: ids[2] }],
      });
      const saved = readPayloadStore(reopened.package);
      expect(
        saved.status === 'ready' &&
          [...saved.entries].map(([id, entry]) => [id, entry.raw, entry.label])
      ).toEqual(stored);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(groups(editor).map((group) => group.id)).toEqual(ids);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(initial);
      expect(decisionBlocks.map((id) => blockLabel(editor, id))).toEqual(labels);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(groups(editor).map((group) => group.id)).toEqual([ids[0]!, ids[2]!]);
    } finally {
      editor.destroy();
    }
  });

for (const duration of [undefined, 0])
  test(`rejecting an optional property retains its exact absent or zero baseline; duration=${duration}`, () => {
    let pkg = groupDecisionFixture();
    const owner = buildBlockIndex(pkg).byId.get('middle')!;
    const attributes = { ...owner.payload.attributes };
    delete attributes.duration;
    if (duration !== undefined) attributes.duration = duration;
    const updated = withPayloadRecord(pkg, { ...owner.payload, attributes }, 'Stored middle');
    if (!updated.ok) throw new Error(updated.reason);
    pkg = updated.pkg;
    const editor = propertyEditor(pkg);
    try {
      proposeThreeProperties(editor);
      const ids = groups(editor).map((group) => group.id);
      expect(rejectBlock(editor).success).toBe(true);
      const current = buildBlockIndex(source(editor)).byId.get('middle')!.payload.attributes;
      expect(Object.hasOwn(current, 'duration')).toBe(duration !== undefined);
      if (duration === undefined) expect(current.duration).toBeUndefined();
      else expect(current.duration).toBe(duration);
      expect(current.name).toBe('Old middle');
      expect(groups(editor).map((group) => group.id)).toEqual([ids[0]!, ids[2]!]);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(groups(editor).map((group) => group.id)).toEqual(ids);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      const again = buildBlockIndex(source(editor)).byId.get('middle')!.payload.attributes;
      expect(Object.hasOwn(again, 'duration')).toBe(duration !== undefined);
      if (duration === undefined) expect(again.duration).toBeUndefined();
      else expect(again.duration).toBe(duration);
    } finally {
      editor.destroy();
    }
  });

for (const action of ['accept', 'reject'] as const)
  test(`a legacy property ${action} after one rejected group touches only pending blocks`, () => {
    const editor = propertyEditor(groupDecisionFixture());
    try {
      const proposal = proposeThreeProperties(editor).proposal!;
      expect(rejectBlock(editor).success).toBe(true);
      const pending = groups(editor).map((group) => group.id);
      expect(
        decidePropertyProposal(editor, proposal.id, action, groupToken(editor), propertyContext)
          .success
      ).toBe(true);
      expect(groups(editor)).toEqual([]);
      for (const id of decisionBlocks)
        expect(buildBlockIndex(source(editor)).byId.get(id)!.payload.attributes.name).toBe(
          id === 'middle' || action === 'reject' ? `Old ${id}` : `New ${id}`
        );
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(groups(editor).map((group) => group.id)).toEqual(pending);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(groups(editor)).toHaveLength(3);
    } finally {
      editor.destroy();
    }
  });

for (const action of ['accept', 'reject'] as const)
  test(`reject one atomic group preserves original assets and permits legacy ${action} on its neighbor`, async () => {
    const editor = propertyEditor(atomicFixture());
    try {
      expect(executeAtomic(editor, [proposedVisual, proposedListening]).success).toBe(true);
      const pending = groups(editor).map((group) => group.id),
        stored = payloads(editor);
      expect(rejectBlock(editor, 'visual').success).toBe(true);
      expect(buildBlockIndex(source(editor)).byId.get('visual')!.payload).toEqual(atomicVisual);
      expect(buildBlockIndex(source(editor)).byId.get('audio')!.payload).toEqual(atomicListening);
      expect(payloads(editor)).toEqual(stored);
      expect(groups(editor)).toHaveLength(1);
      expect(
        decideAtomicProposal(editor, 'atomic-proposal', action, groupToken(editor), propertyContext)
          .success
      ).toBe(true);
      expect(groups(editor)).toEqual([]);
      expect(buildBlockIndex(source(editor)).byId.get('visual')!.payload).toEqual(atomicVisual);
      expect(buildBlockIndex(source(editor)).byId.get('audio')!.payload).toEqual(
        action === 'accept' ? proposedListening : atomicListening
      );
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(groups(editor)).toHaveLength(1);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(groups(editor).map((group) => group.id)).toEqual(pending);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(groups(editor)).toHaveLength(1);
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      expect(buildBlockIndex(reopened.package).byId.get('visual')!.payload).toEqual(atomicVisual);
      expect(readReviewGroups(reopened.package)).toMatchObject({
        success: true,
        groups: [{ id: pending[1] }],
      });
    } finally {
      editor.destroy();
    }
  });

test('rejecting native group two restores only its old text and the same pending identity on undo', async () => {
  const editor = propertyEditor(groupFixture());
  try {
    proposeGroupText(editor, 'The cat sat.', 'The car sat.');
    proposeGroupText(editor, 'A dog ran.', 'A dog slept.');
    proposeGroupText(editor, 'Keep tail', 'Keep tail!');
    const pending = groups(editor),
      ids = pending.map((group) => group.id),
      before = canonicalOoxmlFingerprint(editor.surface!.session.part());
    expect(ids).toHaveLength(3);
    expect(rejectReviewGroup(editor, ids[1]!, groupToken(editor), propertyContext).success).toBe(
      true
    );
    expect(groups(editor).map((group) => group.id)).toEqual([ids[0]!, ids[2]!]);
    expect(groups(editor).map((group) => group.native)).toEqual([
      pending[0]!.native,
      pending[2]!.native,
    ]);
    expect(paragraphTextOf(editor.surface!.session.part(), pending[1]!.anchor!.paragraphId!)).toBe(
      'A dog ran.'
    );
    const working = await readWorkingView(editor.surface!.session, groupToken(editor));
    if (!working.success) throw new Error(working.error);
    const parsed = parseTargetMarkdown(working.view.projection.text);
    if (!parsed.success) throw new Error(parsed.error);
    const text = plainText(parsed.target.normalizedAst);
    expect(text).toContain('The car sat.');
    expect(text).toContain('A dog ran.');
    expect(text).not.toContain('A dog slept.');
    expect(text).toContain('Keep tail!');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
    expect(groups(editor).map((group) => group.id)).toEqual(ids);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    expect(readReviewGroups(saved.package)).toMatchObject({
      success: true,
      groups: [{ id: ids[0] }, { id: ids[2] }],
    });
  } finally {
    editor.destroy();
  }
});

test('one rejected body/header group restores old native content together and retains its unrelated group', async () => {
  const editor = propertyEditor(crossStoryGroupFixture());
  try {
    const id = 'g1:stored:rg_cross_story',
      ids = groups(editor).map((group) => group.id),
      before = new Map(
        [...source(editor).parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)])
      );
    expect(rejectReviewGroup(editor, id, groupToken(editor), propertyContext).success).toBe(true);
    expect(groups(editor).map((group) => group.id)).toEqual(ids.filter((value) => value !== id));
    expect(
      paragraphTextOf(editor.surface!.session.part(), editor.surface!.session.paragraphIds()[0]!)
    ).toBe('Body ');
    expect(collectRevisionSites(source(editor).parts.get('/word/header1.xml')!)).toEqual([]);
    expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(1);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor).map((group) => group.id)).toEqual(ids);
    for (const name of ['/word/document.xml', '/word/header1.xml'])
      expect(canonicalOoxmlFingerprint(source(editor).parts.get(name)!)).toBe(before.get(name)!);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    expect(collectRevisionSites(saved.package.parts.get('/word/header1.xml')!)).toEqual([]);
    expect(collectRevisionSites(saved.package.parts.get('/word/document.xml')!)).toHaveLength(1);
  } finally {
    editor.destroy();
  }
});

test('imported rejection reserves removed Word IDs and restores the original group identity on undo', () => {
  const read = readOoxmlPackage(
    docxFixture(
      '<w:p><w:ins w:id="5" w:author="Teacher"><w:r><w:t>Pending</w:t></w:r></w:ins></w:p>'
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const editor = propertyEditor(read.package);
  try {
    const id = groups(editor)[0]!.id;
    expect(rejectReviewGroup(editor, id, groupToken(editor), propertyContext).success).toBe(true);
    expect(groups(editor)).toEqual([]);
    expect(editor.surface!.session.bodyText()).toBe('');
    expect(reviewGroupRevisionReservations(source(editor))).toEqual([
      { partName: '/word/document.xml', ids: ['5'] },
    ]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor)[0]!.id).toBe(id);
    expect(editor.surface!.session.bodyText()).toBe('Pending');
  } finally {
    editor.destroy();
  }
});

test('stale, viewing, unknown, wrong-epoch, missing-module and locked rejection refuse before mutation', () => {
  const editor = propertyEditor(groupDecisionFixture());
  try {
    proposeThreeProperties(editor);
    const id = groups(editor)[1]!.id,
      token = groupToken(editor),
      before = writeOoxmlPackage(source(editor));
    for (const [groupId, expected, identity] of [
      [id, { ...token, revision: token.revision - 1 }, propertyContext],
      ['missing', token, propertyContext],
      [id, token, { ...propertyContext, epoch: 'other' }],
    ] as const)
      expect(rejectReviewGroup(editor, groupId, expected, identity).success).toBe(false);
    editor.surface!.setEditingMode('view');
    expect(rejectReviewGroup(editor, id, token, propertyContext).success).toBe(false);
    expect(writeOoxmlPackage(source(editor))).toEqual(before);
    editor.surface!.setEditingMode('edit');
    const control = buildBlockIndex(source(editor)).byId.get('middle')!.controlId;
    editor.surface!.applyAutomationOps(() => [
      { op: 'setContentControlProperties', controlId: control, lock: 'contentLocked' },
    ]);
    const locked = writeOoxmlPackage(source(editor));
    expect(rejectReviewGroup(editor, id, groupToken(editor), propertyContext).success).toBe(false);
    expect(writeOoxmlPackage(source(editor))).toEqual(locked);
    const bare = propertyEditor(source(editor), false);
    try {
      expect(rejectReviewGroup(bare, id, groupToken(bare), propertyContext).success).toBe(false);
    } finally {
      bare.destroy();
    }
  } finally {
    editor.destroy();
  }
});

test('native property rejection restores absent cell alignment with exact Word state and one history frame', async () => {
  const loaded = readOoxmlPackage(cellReviewDocument());
  if (!loaded.ok) throw new Error(loaded.reason);
  const editor = propertyEditor(loaded.package);
  try {
    const original = canonicalOoxmlFingerprint(editor.surface!.session.part()),
      scope = cellScopeFor(editor);
    const plan = planTableFormat(
      {
        revision: groupToken(editor),
        part: editor.surface!.session.part(),
        tableId: scope.tableId,
      },
      { op: 'setTableCellVerticalAlignment', cellIds: [scope.cellId], alignment: 'center' }
    );
    if (!plan.success) throw new Error(plan.error);
    expect(trackTableFormatPlan(editor, plan.plan, propertyContext).success).toBe(true);
    const id = groups(editor)[0]!.id,
      pending = canonicalOoxmlFingerprint(editor.surface!.session.part());
    expect(rejectReviewGroup(editor, id, groupToken(editor), propertyContext).success).toBe(true);
    expect(groups(editor)).toEqual([]);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(original);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(groups(editor)[0]!.id).toBe(id);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(original);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    expect(canonicalOoxmlFingerprint(saved.package.parts.get('/word/document.xml')!)).toBe(
      original
    );
  } finally {
    editor.destroy();
  }
});
