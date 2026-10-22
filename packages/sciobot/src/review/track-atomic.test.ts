import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import oracle from '../__tests__/fixtures/atomic-assets.json';
import {
  atomicFixture,
  atomicLabel,
  atomicListening,
  atomicPlan,
  atomicVisual,
  executeAtomic,
  proposedListening,
  proposedVisual,
} from '../../test/support/atomic-execution';
import {
  propertyContext,
  propertyEditor,
  propertyToken,
} from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { insertBlock } from '../blocks/insert';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import type { BlockPayload, JsonObject } from '../blocks/payload-schema';
import { createEditPlan } from '../diff/plan';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { readPropertyReviewStore, withPropertyReviewRecord } from './property-record-store';
import { trackPropertyPlan } from './track-properties';
import { decidePropertyProposal } from './decide-properties';
import { decideAtomicProposal, trackAtomicPlan } from './track-atomic';

function records(editor: ReturnType<typeof propertyEditor>) {
  const read = readPropertyReviewStore(editor.surface!.session.currentPackage());
  if (read.status !== 'ready') throw new Error('Missing review store');
  return read.records;
}
function payload(editor: ReturnType<typeof propertyEditor>, id = 'visual') {
  return buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(id)!.payload;
}
for (const decision of ['accept', 'reject'] as const)
  test(`atomic ${decision} selects the complete payload, assets, and fallback in one history unit`, async () => {
    const editor = propertyEditor(atomicFixture());
    try {
      const session = editor.surface!.session,
        before = writeOoxmlPackage(session.currentPackage()),
        index = buildBlockIndex(session.currentPackage()),
        phase = index.byId.get('phase')!.controlId,
        control = index.byId.get('visual')!.controlId,
        untouched = session
          .paragraphIds()
          .find((id) => paragraphTextOf(session.part(), id) === 'Keep tail')!,
        fingerprint = canonicalOoxmlFingerprint(findNode(session.part(), untouched)!),
        revision = session.packageRevision();
      let publications = 0;
      const stop = session.subscribe(() => publications++),
        result = executeAtomic(editor, [proposedVisual, proposedListening]);
      stop();
      if (!result.success || !result.proposal)
        throw new Error(result.success ? 'Missing receipt' : result.error);
      expect(publications).toBe(1);
      expect(session.packageRevision()).toBe(revision + 1);
      expect(result.proposal.blocks.map((block) => block.changes[0])).toEqual([
        {
          name: '$payload',
          before: { present: true, value: atomicVisual },
          after: { present: true, value: proposedVisual },
        },
        {
          name: '$payload',
          before: { present: true, value: atomicListening },
          after: { present: true, value: proposedListening },
        },
      ]);
      expect(payload(editor)).toEqual(atomicVisual);
      expect(payload(editor, 'audio')).toEqual(atomicListening);
      const [committed, working] = await Promise.all([
        readCommittedView(session, propertyToken(editor)),
        readWorkingView(session, propertyToken(editor)),
      ]);
      if (!committed.success || !working.success) throw new Error('Atomic views failed');
      expect(committed.view.projection.text).toContain('src="vfs://pristine.png"');
      expect(committed.view.projection.text).toContain('audio="original-audio"');
      expect(working.view.projection.text).toContain('src="vfs://new.png"');
      expect(working.view.projection.text).toContain('audio="new-audio"');
      expect(committed.view.projection.text).toContain('Good morning!');
      expect(working.view.projection.text).toContain('Good evening!');
      const reopened = propertyEditor(session.currentPackage());
      try {
        const pending = records(reopened).get('atomic-proposal')!;
        expect(pending.decoded).toEqual(records(editor).get('atomic-proposal')!.decoded);
        const decided = decideAtomicProposal(
          reopened,
          'atomic-proposal',
          decision,
          propertyToken(reopened),
          propertyContext
        );
        if (!decided.success) throw new Error(decided.error);
        expect(payload(reopened)).toEqual(decision === 'accept' ? proposedVisual : atomicVisual);
        expect(payload(reopened, 'audio')).toEqual(
          decision === 'accept' ? proposedListening : atomicListening
        );
      } finally {
        reopened.destroy();
      }
      const decided = decideAtomicProposal(
        editor,
        'atomic-proposal',
        decision,
        propertyToken(editor),
        propertyContext
      );
      if (!decided.success) throw new Error(decided.error);
      expect(payload(editor)).toEqual(decision === 'accept' ? proposedVisual : atomicVisual);
      expect(payload(editor, 'audio')).toEqual(
        decision === 'accept' ? proposedListening : atomicListening
      );
      for (const id of ['visual', 'audio'])
        expect(paragraphTextOf(session.part(), atomicLabel(editor, id).id)).toBe(
          `${decision === 'accept' ? 'New' : 'Old'} ${id} fallback`
        );
      expect(records(editor).size).toBe(0);
      expect(collectRevisionSites(session.part())).toHaveLength(0);
      expect(buildBlockIndex(session.currentPackage()).byId.get('phase')!.controlId).toBe(phase);
      expect(buildBlockIndex(session.currentPackage()).byId.get('visual')!.controlId).toBe(control);
      expect(canonicalOoxmlFingerprint(findNode(session.part(), untouched)!)).toBe(fingerprint);
      const saved = readOoxmlPackage(writeOoxmlPackage(session.currentPackage()));
      if (!saved.ok) throw new Error(saved.reason);
      expect(buildBlockIndex(saved.package).byId.get('visual')!.payload).toEqual(
        decision === 'accept' ? proposedVisual : atomicVisual
      );
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(records(editor).size).toBe(1);
      expect(payload(editor)).toEqual(atomicVisual);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(records(editor).size).toBe(0);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
      expect(session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  });

function listeningPayload(node: JsonObject): BlockPayload {
  return {
    id: 'audio',
    type: 'listening',
    version: 1,
    attributes: {},
    ...Object.fromEntries(
      Object.entries(node).filter(([key]) => !['id', 'type', 'children'].includes(key))
    ),
  };
}
for (const item of oracle.cases.filter((item) => item.rootChanges.length))
  for (const decision of ['accept', 'reject'] as const)
    test(`atomic ${decision} matches captured listening assets: ${item.name}`, () => {
      const before = listeningPayload(item.before as unknown as JsonObject),
        after = listeningPayload(item.after as unknown as JsonObject),
        editor = propertyEditor(atomicFixture([before]));
      try {
        const result = executeAtomic(editor, [after]);
        if (!result.success) throw new Error(result.error);
        expect(payload(editor, 'audio')).toEqual(before);
        const decided = decideAtomicProposal(
          editor,
          'atomic-proposal',
          decision,
          propertyToken(editor),
          propertyContext
        );
        if (!decided.success) throw new Error(decided.error);
        expect(payload(editor, 'audio')).toEqual(decision === 'accept' ? after : before);
        expect(payload(editor, 'audio').blockKey).toBe('budget-keep');
      } finally {
        editor.destroy();
      }
    });

test('atomic no-op preserves revision and does not add history', () => {
  const editor = propertyEditor(atomicFixture());
  try {
    const revision = propertyToken(editor);
    const plan = createEditPlan(revision, []);
    if (!plan.success) throw new Error(plan.error);
    expect(trackAtomicPlan(editor, plan.plan, [], propertyContext)).toEqual({
      success: true,
      revision,
      proposal: null,
    });
    expect(editor.surface!.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

test('unchanged native fallback still retains a complete metadata-only atomic proposal', () => {
  const editor = propertyEditor(atomicFixture());
  try {
    const result = executeAtomic(
      editor,
      [proposedVisual],
      'atomic-proposal',
      'Old visual fallback'
    );
    if (!result.success || !result.proposal) throw new Error('Missing atomic proposal');
    expect(result.proposal.nativeRevisions).toEqual([]);
    expect(
      decideAtomicProposal(
        editor,
        'atomic-proposal',
        'accept',
        propertyToken(editor),
        propertyContext
      ).success
    ).toBe(true);
    expect(payload(editor)).toEqual(proposedVisual);
  } finally {
    editor.destroy();
  }
});

test('an inserted atomic widget retains its generated label, native caret paragraph, and control identity', () => {
  const editor = propertyEditor(atomicFixture([]));
  try {
    const session = editor.surface!.session,
      phase = buildBlockIndex(session.currentPackage()).byId.get('phase')!;
    const inserted = insertBlock(editor, {
      parentControlId: phase.controlId,
      beforeId: null,
      payload: atomicVisual,
      label: 'Stored visual label',
      generatedLabel: 'Old visual fallback',
    });
    if (!inserted.success) throw new Error(inserted.error);
    const caret = findNode(session.part(), inserted.caret.paragraphId)!;
    expect(executeAtomic(editor).success).toBe(true);
    const decided = decideAtomicProposal(
      editor,
      'atomic-proposal',
      'accept',
      propertyToken(editor),
      propertyContext
    );
    if (!decided.success) throw new Error(decided.error);
    expect(payload(editor)).toEqual(proposedVisual);
    expect(buildBlockIndex(session.currentPackage()).byId.get('visual')!.controlId).toBe(
      inserted.controlId
    );
    expect(canonicalOoxmlFingerprint(findNode(session.part(), caret.id)!)).toBe(
      canonicalOoxmlFingerprint(caret)
    );
  } finally {
    editor.destroy();
  }
});

for (const mode of ['view', 'locked', 'no-review'] as const)
  test(`atomic ${mode} refuses without a payload or history write`, () => {
    const editor = propertyEditor(atomicFixture(), mode !== 'no-review');
    try {
      const session = editor.surface!.session;
      if (mode === 'view') editor.surface!.setEditingMode('view');
      if (mode === 'locked') {
        const owner = buildBlockIndex(session.currentPackage()).byId.get('visual')!;
        expect(
          editor.surface!.applyAutomationOps(
            () => [
              {
                op: 'setContentControlProperties',
                controlId: owner.controlId,
                lock: 'contentLocked',
              },
            ],
            { kind: 'body' }
          ).rejected
        ).toBe(false);
      }
      const bytes = writeOoxmlPackage(session.currentPackage()),
        revision = session.packageRevision(),
        canUndo = session.canUndo();
      expect(executeAtomic(editor).success).toBe(false);
      expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
      expect(session.packageRevision()).toBe(revision);
      expect(session.canUndo()).toBe(canUndo);
    } finally {
      editor.destroy();
    }
  });

test('a root replacement cannot change a container, stable identity, or payload type', () => {
  for (const invalid of ['container', 'identity', 'type'] as const) {
    const editor = propertyEditor(atomicFixture());
    try {
      const session = editor.surface!.session,
        owner = buildBlockIndex(session.currentPackage()).byId.get(
          invalid === 'container' ? 'phase' : 'visual'
        )!,
        after = {
          ...owner.payload,
          ...(invalid === 'identity'
            ? { id: 'other' }
            : invalid === 'type'
              ? { type: 'source' }
              : { attributes: { phaseType: 'reflection' } }),
        },
        plan = createEditPlan(propertyToken(editor), [
          {
            kind: 'payload',
            blockId: owner.id,
            partName: owner.partName,
            controlId: owner.controlId,
            path: [],
            before: { present: true, value: owner.payload },
            after: { present: true, value: after },
          },
        ]);
      if (!plan.success) throw new Error(plan.error);
      const bytes = writeOoxmlPackage(session.currentPackage());
      expect(
        trackAtomicPlan(
          editor,
          plan.plan,
          [{ blockId: owner.id, before: 'Old visual fallback', after: 'New fallback' }],
          propertyContext
        ).success
      ).toBe(false);
      expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
      expect(session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  }
});

for (const scenario of [
  'stale-plan',
  'stale-fallback',
  'missing-fallback',
  'property-executor',
  'wrong-identity',
  'overlap',
  'changed-payload',
  'missing-receipt',
  'partial-receipt',
  'wrong-decision-executor',
] as const)
  test(`atomic refusal is write-free: ${scenario}`, () => {
    let editor = propertyEditor(atomicFixture());
    try {
      const plan = atomicPlan(editor, [proposedVisual]),
        target = { blockId: 'visual', before: 'Old visual fallback', after: 'New visual fallback' };
      let decide = false;
      if (
        [
          'overlap',
          'changed-payload',
          'missing-receipt',
          'partial-receipt',
          'wrong-decision-executor',
        ].includes(scenario)
      ) {
        const result = executeAtomic(editor);
        if (!result.success) throw new Error(result.error);
        decide = scenario !== 'overlap';
      }
      if (scenario === 'stale-plan')
        editor.surface!.applyAutomationOps(
          () => [
            { op: 'insertText', paragraphId: atomicLabel(editor).id, offset: 0, text: 'Peer ' },
          ],
          { kind: 'body' }
        );
      if (
        scenario === 'changed-payload' ||
        scenario === 'missing-receipt' ||
        scenario === 'partial-receipt'
      ) {
        const pkg = editor.surface!.session.currentPackage();
        let changed;
        if (scenario === 'changed-payload')
          changed = withPayloadRecord(
            pkg,
            { ...atomicVisual, attributes: { ...atomicVisual.attributes, fileId: 'peer-file' } },
            'Stored visual label'
          );
        else {
          const record = records(editor).get('atomic-proposal')!.decoded;
          if (record.status !== 'supported') throw new Error('Invalid record');
          expect(record.value.nativeRevisions.length).toBeGreaterThan(1);
          changed = withPropertyReviewRecord(pkg, {
            ...record.value,
            nativeRevisions:
              scenario === 'partial-receipt' ? record.value.nativeRevisions.slice(0, -1) : [],
          });
        }
        if (!changed.ok) throw new Error(changed.reason);
        editor.destroy();
        editor = propertyEditor(changed.pkg);
      }
      const session = editor.surface!.session,
        before = writeOoxmlPackage(session.currentPackage()),
        revision = session.packageRevision();
      const result = decide
        ? scenario === 'wrong-decision-executor'
          ? decidePropertyProposal(
              editor,
              'atomic-proposal',
              'accept',
              propertyToken(editor),
              propertyContext
            )
          : decideAtomicProposal(
              editor,
              'atomic-proposal',
              'accept',
              propertyToken(editor),
              propertyContext
            )
        : scenario === 'property-executor'
          ? trackPropertyPlan(editor, plan, [target], propertyContext)
          : trackAtomicPlan(
              editor,
              plan,
              scenario === 'missing-fallback'
                ? []
                : [{ ...target, ...(scenario === 'stale-fallback' ? { before: 'Wrong' } : {}) }],
              {
                ...propertyContext,
                proposalId: 'other-proposal',
                ...(scenario === 'wrong-identity' ? { epoch: 'other-epoch' } : {}),
              }
            );
      expect(result.success).toBe(false);
      expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
      expect(session.packageRevision()).toBe(revision);
    } finally {
      editor.destroy();
    }
  });
