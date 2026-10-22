import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import oracle from '../__tests__/fixtures/domain-properties.json';
import {
  executeProperties,
  nativePropertyPlan,
  propertyContext,
  propertyEditor,
  propertyLabel,
  propertyToken,
} from '../../test/support/property-execution';
import { propertyActivity, propertyRecordFixture } from '../../test/support/property-record';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import type { JsonObject } from '../data/json';
import { createEditPlan } from '../diff/plan';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { readPropertyReviewStore, withPropertyReviewRecord } from './property-record-store';
import { nativePropertyAddress, propertyAddressKey } from './property-fallback';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackPropertyPlan } from './track-properties';
import { decidePropertyProposal } from './decide-properties';

function attrs(editor: ReturnType<typeof propertyEditor>) {
  return propertyActivity(editor.surface!.session.currentPackage()).payload.attributes;
}
function records(editor: ReturnType<typeof propertyEditor>) {
  const read = readPropertyReviewStore(editor.surface!.session.currentPackage());
  if (read.status !== 'ready') throw new Error('Missing native property records');
  return read.records;
}
for (const decision of ['accept', 'reject'] as const)
  test(`property ${decision} resolves metadata and tracked Word fallback in one history unit`, async () => {
    const editor = propertyEditor();
    try {
      const session = editor.surface!.session,
        pkg = session.currentPackage(),
        phase = buildBlockIndex(pkg).byId.get('phase')!,
        activity = propertyActivity(pkg),
        label = propertyLabel(editor),
        body = storyParagraphs(
          contentControlContentNodeOf(findNode(session.part(), activity.controlId)!)!
        ).find((paragraph) => paragraphTextOf(session.part(), paragraph.id) === 'Body')!,
        kept = canonicalOoxmlFingerprint(body),
        before = canonicalOoxmlFingerprint(session.part()),
        beforeAttributes = attrs(editor),
        revision = session.packageRevision();
      let publications = 0;
      const stop = session.subscribe(() => publications++),
        result = executeProperties(editor);
      stop();
      if (!result.success || !result.proposal)
        throw new Error(result.success ? 'Missing proposal' : result.error);
      expect(publications).toBe(1);
      expect(session.packageRevision()).toBe(revision + 1);
      expect(attrs(editor)).toEqual(beforeAttributes);
      expect(
        result.proposal.blocks[0]!.changes.map((change) => [
          change.name,
          change.before,
          change.after,
        ])
      ).toEqual([
        [
          'name',
          { present: true, value: 'Old activity' },
          { present: true, value: 'New activity' },
        ],
        ['duration', { present: true, value: 10 }, { present: true, value: 15 }],
      ]);
      expect(result.proposal.nativeRevisions.length).toBeGreaterThan(0);
      const actual = collectRevisionSites(session.part()).map((site) =>
        propertyAddressKey(session.part().name, nativePropertyAddress(site.node)!)
      );
      expect(
        result.proposal.nativeRevisions
          .map((address) => propertyAddressKey(address.partName, address))
          .sort()
      ).toEqual([...new Set(actual)].sort());
      expect(
        result.proposal.nativeRevisions.every(
          (address) => address.author === 'ScioBot' && address.date === '2026-10-19T09:00:00.000Z'
        )
      ).toBe(true);
      const token = propertyToken(editor),
        [committed, working] = await Promise.all([
          readCommittedView(session, token),
          readWorkingView(session, token),
        ]);
      if (!committed.success || !working.success) throw new Error('Native property views failed');
      expect(committed.view.projection.text).toContain('name="Old activity" duration="10"');
      expect(working.view.projection.text).toContain('name="New activity" duration="15"');
      const pending = canonicalOoxmlFingerprint(session.part());
      expect(
        decidePropertyProposal(
          editor,
          result.proposal.id,
          decision,
          propertyToken(editor),
          propertyContext
        ).success
      ).toBe(true);
      expect(paragraphTextOf(session.part(), label.id)).toBe(
        decision === 'accept' ? 'New activity — 15 minutes' : 'Old activity — 10 minutes'
      );
      expect(attrs(editor)).toEqual(
        decision === 'accept'
          ? { ...beforeAttributes, name: 'New activity', duration: 15 }
          : beforeAttributes
      );
      expect(collectRevisionSites(session.part())).toHaveLength(0);
      expect(records(editor).size).toBe(0);
      expect(buildBlockIndex(session.currentPackage()).byId.get('phase')!.controlId).toBe(
        phase.controlId
      );
      expect(propertyActivity(session.currentPackage()).controlId).toBe(activity.controlId);
      expect(canonicalOoxmlFingerprint(findNode(session.part(), body.id)!)).toBe(kept);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
      expect(attrs(editor)).toEqual(beforeAttributes);
      expect(records(editor).size).toBe(1);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(records(editor).size).toBe(0);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
      expect(records(editor).size).toBe(0);
      expect(session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  });

for (const item of oracle.cases)
  for (const decision of ['accept', 'reject'] as const)
    test(`live property presence matches captured ${item.name}; ${decision}`, () => {
      const original = propertyRecordFixture(),
        owner = propertyActivity(original),
        changed = withPayloadRecord(
          original,
          { ...owner.payload, attributes: item.before as JsonObject },
          'Stored original label'
        );
      if (!changed.ok) throw new Error(changed.reason);
      const editor = propertyEditor(changed.pkg);
      try {
        const result = executeProperties(editor, {
          set: Object.fromEntries(
            item.keys
              .filter((key) => Object.hasOwn(item.after, key))
              .map((key) => [key, (item.after as JsonObject)[key]!])
          ),
          remove: item.keys.filter((key) => !Object.hasOwn(item.after, key)),
        });
        if (!result.success || !result.proposal)
          throw new Error(result.success ? 'Missing proposal' : result.error);
        expect(
          decidePropertyProposal(
            editor,
            result.proposal.id,
            decision,
            propertyToken(editor),
            propertyContext
          ).success
        ).toBe(true);
        expect(attrs(editor)).toEqual(
          decision === 'accept' ? (item.working as JsonObject) : (item.committed as JsonObject)
        );
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(attrs(editor)).toEqual(item.before as JsonObject);
        expect(records(editor).size).toBe(1);
      } finally {
        editor.destroy();
      }
    });

for (const decision of ['accept', 'reject'] as const)
  test(`pending property records and exact native fallback addresses survive save/reopen before ${decision}`, async () => {
    const editor = propertyEditor();
    let reopened: ReturnType<typeof propertyEditor> | undefined;
    try {
      const result = executeProperties(editor);
      if (!result.success || !result.proposal) throw new Error('Missing proposal');
      const bytes = new Uint8Array(await editor.save()),
        loaded = readOoxmlPackage(bytes);
      if (!loaded.ok) throw new Error(loaded.reason);
      reopened = createDocxEditor({
        container: document.createElement('div'),
        document: bytes,
        modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      });
      expect(records(reopened).get(result.proposal.id)!.decoded).toEqual({
        status: 'supported',
        value: result.proposal,
      });
      expect(
        decidePropertyProposal(
          reopened,
          result.proposal.id,
          decision,
          propertyToken(reopened, 'reload'),
          { ...propertyContext, epoch: 'reload' }
        ).success
      ).toBe(true);
      expect(attrs(reopened).duration).toBe(decision === 'accept' ? 15 : 10);
      expect(paragraphTextOf(reopened.surface!.session.part(), propertyLabel(reopened).id)).toBe(
        decision === 'accept' ? 'New activity — 15 minutes' : 'Old activity — 10 minutes'
      );
      expect(editor.surface!.session.canUndo()).toBe(true);
    } finally {
      reopened?.destroy();
      editor.destroy();
    }
  });

test('viewing, missing native review support, stale plans, mismatched fallback text, and mixed edits refuse atomically', () => {
  for (const review of [false, true]) {
    const editor = propertyEditor(undefined, review);
    try {
      const session = editor.surface!.session,
        bytes = writeOoxmlPackage(session.currentPackage()),
        plan = nativePropertyPlan(editor, { set: { duration: 15 } }),
        target = [{ blockId: 'activity', before: 'Old activity — 10 minutes', after: 'New label' }];
      if (!review)
        expect(trackPropertyPlan(editor, plan, target, propertyContext).success).toBe(false);
      expect(
        trackPropertyPlan(
          editor,
          { ...plan, revision: { ...plan.revision, revision: 100 } },
          target,
          propertyContext
        ).success
      ).toBe(false);
      expect(
        trackPropertyPlan(editor, plan, [{ ...target[0]!, before: 'Wrong' }], propertyContext)
          .success
      ).toBe(false);
      editor.surface!.setEditingMode('view');
      expect(trackPropertyPlan(editor, plan, target, propertyContext).success).toBe(false);
      editor.surface!.setEditingMode('edit');
      const mixed = createEditPlan(plan.revision, [
        ...plan.operations,
        {
          kind: 'text',
          partName: session.part().name,
          paragraphId: propertyLabel(editor).id,
          start: 0,
          end: 0,
          before: '',
          after: 'Unowned',
        },
      ]);
      if (!mixed.success) throw new Error(mixed.error);
      expect(trackPropertyPlan(editor, mixed.plan, target, propertyContext).success).toBe(false);
      expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
      expect(session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  }
});

test('metadata-only requests preserve generated text while their decisions remain separate native history units', () => {
  const editor = propertyEditor();
  try {
    const label = propertyLabel(editor),
      before = canonicalOoxmlFingerprint(label),
      result = executeProperties(editor, { set: { duration: 15 } }, 'Old activity — 10 minutes');
    if (!result.success || !result.proposal) throw new Error('Missing proposal');
    expect(result.proposal.nativeRevisions).toEqual([]);
    expect(canonicalOoxmlFingerprint(findNode(editor.surface!.session.part(), label.id)!)).toBe(
      before
    );
    expect(
      decidePropertyProposal(
        editor,
        result.proposal.id,
        'accept',
        propertyToken(editor),
        propertyContext
      ).success
    ).toBe(true);
    expect(attrs(editor).duration).toBe(15);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(attrs(editor).duration).toBe(10);
    expect(records(editor).size).toBe(1);
  } finally {
    editor.destroy();
  }
});

for (const after of ['Žák 🧑🏽‍🚀\t0\nminutes', ''])
  for (const decision of ['accept', 'reject'] as const)
    test(`generated Unicode, tab, hard-break, and empty targets retain exact native outcomes: ${JSON.stringify(after)}; ${decision}`, () => {
      const editor = propertyEditor();
      try {
        const result = executeProperties(editor, { set: { name: 'Žák', duration: 0 } }, after);
        if (!result.success || !result.proposal)
          throw new Error(result.success ? 'Missing proposal' : result.error);
        expect(
          decidePropertyProposal(
            editor,
            result.proposal.id,
            decision,
            propertyToken(editor),
            propertyContext
          ).success
        ).toBe(true);
        expect(paragraphTextOf(editor.surface!.session.part(), propertyLabel(editor).id)).toBe(
          decision === 'accept' ? after : 'Old activity — 10 minutes'
        );
        expect(attrs(editor).duration).toBe(decision === 'accept' ? 0 : 10);
      } finally {
        editor.destroy();
      }
    });

test('metadata-only writes respect control locks and never add a proposal or history unit', () => {
  const editor = propertyEditor();
  try {
    const surface = editor.surface!,
      session = surface.session,
      owner = propertyActivity(session.currentPackage());
    expect(
      surface.applyAutomationOps(() => [
        { op: 'setContentControlProperties', controlId: owner.controlId, lock: 'contentLocked' },
      ]).committed
    ).toBe(true);
    const bytes = writeOoxmlPackage(session.currentPackage()),
      revision = session.packageRevision();
    expect(
      executeProperties(editor, { set: { duration: 15 } }, 'Old activity — 10 minutes')
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(session.packageRevision()).toBe(revision);
    expect(records(editor).size).toBe(0);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

test('external native review and incomplete saved receipts require reconciliation instead of silently deciding metadata', () => {
  for (const mode of ['external-review', 'incomplete-receipt'] as const) {
    const editor = propertyEditor();
    try {
      const result = executeProperties(editor);
      if (!result.success || !result.proposal) throw new Error('Missing property proposal');
      const surface = editor.surface!,
        session = surface.session;
      if (mode === 'external-review') {
        expect(
          surface.applyAutomationOps(
            () => [{ op: 'acceptAllRevisions' }],
            { kind: 'body' },
            undefined,
            true,
            true
          ).committed
        ).toBe(true);
      } else {
        expect(
          surface.applyAutomationOps(
            () => [],
            { kind: 'body' },
            [
              (pkg) => {
                const changed = withPropertyReviewRecord(pkg, {
                  ...result.proposal!,
                  nativeRevisions: [],
                });
                if (!changed.ok) throw new Error(changed.reason);
                return changed.pkg;
              },
            ],
            false,
            true
          ).committed
        ).toBe(true);
      }
      const bytes = writeOoxmlPackage(session.currentPackage()),
        revision = session.packageRevision();
      expect(
        decidePropertyProposal(
          editor,
          result.proposal.id,
          'accept',
          propertyToken(editor),
          propertyContext
        ).success
      ).toBe(false);
      expect(attrs(editor).duration).toBe(10);
      expect(records(editor).size).toBe(1);
      expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
      expect(session.packageRevision()).toBe(revision);
    } finally {
      editor.destroy();
    }
  }
});

test('stale mount identities and changed selected properties refuse decisions while unrelated metadata survives acceptance', () => {
  for (const selected of [false, true]) {
    const editor = propertyEditor();
    try {
      const result = executeProperties(editor);
      if (!result.success || !result.proposal) throw new Error('Missing property proposal');
      const surface = editor.surface!,
        session = surface.session,
        stale = propertyToken(editor);
      expect(
        decidePropertyProposal(
          editor,
          result.proposal.id,
          'accept',
          { ...stale, epoch: 'foreign-mount' },
          propertyContext
        )
      ).toMatchObject({ success: false, code: 'stale-revision' });
      expect(
        surface.applyAutomationOps(
          () => [],
          { kind: 'body' },
          [
            (pkg) => {
              const owner = propertyActivity(pkg),
                attributes = {
                  ...owner.payload.attributes,
                  ...(selected ? { duration: 12 } : { resource: 'Peer resource' }),
                },
                changed = withPayloadRecord(
                  pkg,
                  { ...owner.payload, attributes },
                  'activity label'
                );
              if (!changed.ok) throw new Error(changed.reason);
              return changed.pkg;
            },
          ],
          false,
          true
        ).committed
      ).toBe(true);
      const bytes = writeOoxmlPackage(session.currentPackage()),
        decision = decidePropertyProposal(
          editor,
          result.proposal.id,
          'accept',
          propertyToken(editor),
          propertyContext
        );
      if (selected) {
        expect(decision).toMatchObject({ success: false, code: 'stale-revision' });
        expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
      } else {
        expect(decision.success).toBe(true);
        expect(attrs(editor).resource).toBe('Peer resource');
        expect(attrs(editor).duration).toBe(15);
      }
    } finally {
      editor.destroy();
    }
  }
});
