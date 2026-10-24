import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  withPart,
  type OoxmlPackage,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { planTargetDocument } from '../diff/plan-target';
import { previewDirectPlan } from '../ai/apply-direct';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readWorkingView } from '../review/working-view';
import { readCommittedView } from '../review/committed-view';
import { readReviewGroups } from '../review/groups';
import { recomposeOwnedProposals } from '../review/recompose';
import { acceptAllReviewGroups } from '../review/accept-all';
import { rejectAllReviewGroups } from '../review/reject-all';

const context = {
  documentId: 'born-inline-paragraph',
  epoch: 'load',
  author: 'AI',
  date: '2026-10-24T12:00:00.000Z',
};
const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
const atoms = {
  mention: '<scio-mention key="user" value="Ada" />',
  source: '<source title="Paper" url="https://source.test" />',
  math: '$x+y$',
};
for (const kind of ['mention', 'source', 'math'] as const)
  for (const mark of ['plain', 'code', 'link'] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`tracked new paragraph atom ${kind}/${mark}/${decision} owns its paragraph and payload`, async () => {
        const editor = createDocxEditor({
          container: document.createElement('div'),
          document: docxFixture('<w:p><w:r><w:t>Start</w:t></w:r></w:p>'),
          modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
        });
        const token = () => readEditorRevision(editor, context.documentId, context.epoch)!;
        try {
          const before = editor.surface!.session.currentPackage();
          const initial = await readWorkingView(editor.surface!.session, token());
          if (!initial.success) throw Error(initial.error);
          const authored =
            mark === 'code'
              ? '<code>' + atoms[kind] + '</code>'
              : mark === 'link'
                ? '[' + atoms[kind] + '](https://example.test)'
                : atoms[kind];
          const target = 'Start\n\nLead. ' + authored + ' Tail.';
          const options = { allocateBlockId: () => 'born-atom' };
          const plan = planTargetDocument(
            { pkg: before, revision: token(), projection: initial.view.projection },
            target,
            options
          );
          if (!plan.success) throw Error(plan.error);
          expect(
            plan.plan.operations.some(
              (edit) => edit.kind === 'siblings' && edit.atoms?.length === 1
            )
          ).toBe(true);
          const direct = previewDirectPlan(before, plan.plan);
          if (!direct.success) throw Error(direct.error);
          const result = await recomposeOwnedProposals(
            editor,
            { revision: token(), targetMarkdown: target, ownership: [] },
            context,
            { target: options, payloads: { presence: { proposalId: 'born-presence' } } }
          );
          if (!result.success) throw Error(result.error);
          const pending = native(editor.surface!.session.currentPackage());
          const accepted = direct.pkg;
          const groups = readReviewGroups(editor.surface!.session.currentPackage());
          if (!groups.success) throw Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          expect(groups.groups[0]!.readOnly).toBe(false);
          expect(groups.groups[0]!.properties).toEqual([
            { proposalId: 'born-presence', blockId: 'born-atom' },
          ]);
          expect(groups.groups[0]!.native).toHaveLength(1);
          expect(groups.groups[0]!.native[0]!.siteNodeIds.length).toBeGreaterThanOrEqual(2);
          const committed = await readCommittedView(editor.surface!.session, token());
          expect(committed.success && committed.view.projection.text).toBe(
            initial.view.projection.text
          );
          const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            editor,
            token(),
            context
          );
          if (!decided.success) throw Error(decided.error);
          const actual = editor.surface!.session.currentPackage();
          expect(native(actual)).toBe(native(decision === 'accept' ? accepted : before));
          const reopened = readOoxmlPackage(writeOoxmlPackage(actual));
          if (!reopened.ok) throw Error(reopened.reason);
          expect(native(reopened.package)).toBe(native(decision === 'accept' ? accepted : before));
          const payloads = readPayloadStore(actual);
          expect(payloads.status === 'ready' && [...payloads.entries.keys()]).toEqual(
            decision === 'accept' ? ['born-atom'] : []
          );
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(native(editor.surface!.session.currentPackage())).toBe(pending);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(native(editor.surface!.session.currentPackage())).toBe(
            native(decision === 'accept' ? accepted : before)
          );
        } finally {
          editor.destroy();
        }
      });

test('multiple born paragraph atoms share one complete review group and later AI waves replace it exactly', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture('<w:p><w:r><w:t>Start</w:t></w:r></w:p>'),
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
  const token = () => readEditorRevision(editor, context.documentId, context.epoch)!;
  try {
    const before = native(editor.surface!.session.currentPackage());
    let count = 0;
    const first = await recomposeOwnedProposals(
      editor,
      {
        revision: token(),
        targetMarkdown:
          'Start\n\nLead ' +
          atoms.mention +
          ' and ' +
          atoms.math +
          '.\n\nNext ' +
          atoms.source +
          '.',
        ownership: [],
      },
      context,
      {
        target: { allocateBlockId: () => 'first-' + ++count },
        payloads: { presence: { proposalId: 'first-born' } },
      }
    );
    if (!first.success) throw Error(first.error);
    const firstWord = native(editor.surface!.session.currentPackage());
    const groups = readReviewGroups(editor.surface!.session.currentPackage());
    if (!groups.success) throw Error(groups.error);
    expect(groups.groups).toHaveLength(1);
    expect(groups.groups[0]!.properties).toHaveLength(3);
    expect(groups.groups[0]!.native).toHaveLength(1);
    const second = await recomposeOwnedProposals(
      editor,
      {
        revision: token(),
        targetMarkdown: 'Start\n\nReplaced ' + atoms.math + '.',
        ownership: first.ownership,
        contentProofs: first.contentProofs,
      },
      context,
      {
        target: { allocateBlockId: () => 'second-atom' },
        payloads: { presence: { proposalId: 'second-born' } },
      }
    );
    if (!second.success) throw Error(second.error);
    expect([...buildBlockIndex(editor.surface!.session.currentPackage()).byId.keys()]).toEqual([
      'second-atom',
    ]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(native(editor.surface!.session.currentPackage())).toBe(firstWord);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const rejected = rejectAllReviewGroups(editor, token(), context);
    if (!rejected.success) throw Error(rejected.error);
    expect(native(editor.surface!.session.currentPackage())).toBe(before);
    expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(0);
  } finally {
    editor.destroy();
  }
});

test('paragraph birth ownership refuses forged membership and reserves its native IDs after resolution', async () => {
  const { readPropertyReviewStore } = await import('./property-record-store');
  const { decodeBlockPresenceRecord, encodeBlockPresenceRecord } =
    await import('./block-presence-schema');
  const { decodePropertyReviewRecord } = await import('./property-record-schema');
  const { bindBlockPresenceRecord } = await import('./block-presence-plan');
  const { readInlinePresenceExtension } = await import('./inline-presence-native');
  const { reviewGroupRevisionReservations } = await import('./group-allocation');
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture('<w:p><w:r><w:t>Start</w:t></w:r></w:p>'),
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
  const token = () => readEditorRevision(editor, context.documentId, context.epoch)!;
  try {
    const staged = await recomposeOwnedProposals(
      editor,
      { revision: token(), targetMarkdown: 'Start\n\nBorn ' + atoms.mention + '.', ownership: [] },
      context,
      {
        target: { allocateBlockId: () => 'born-atom' },
        payloads: { presence: { proposalId: 'born-presence' } },
      }
    );
    if (!staged.success) throw Error(staged.error);
    const pkg = editor.surface!.session.currentPackage();
    const store = readPropertyReviewStore(pkg);
    if (store.status !== 'ready') throw Error('Missing presence store.');
    const decoded = decodeBlockPresenceRecord(store.records.get('born-presence')!.raw);
    if (decoded.status !== 'supported') throw Error('Missing supported presence record.');
    const record = decoded.value;
    expect(record.schemaVersion).toBe(4);
    expect(decodePropertyReviewRecord(encodeBlockPresenceRecord(record))).toMatchObject({
      status: 'unsupported-version',
      version: 4,
    });
    const owners = readInlinePresenceExtension(record)!;
    const owner = owners.get('born-atom')!;
    expect(owner.gap).toHaveLength(1);
    const forged = (gap: unknown) =>
      ({
        ...record,
        extension: { nativeInline: { schemaVersion: 2, owners: [{ ...owner, gap }] } },
      }) as unknown as typeof record;
    for (const invalid of [
      [],
      [owner.revision, owner.revision],
      [{ ...owner.revision, author: 'Foreign' }],
      [{ ...owner.revision, id: 'missing' }],
    ]) {
      expect(bindBlockPresenceRecord(pkg, forged(invalid)).success).toBe(false);
      expect(editor.surface!.session.currentPackage()).toBe(pkg);
    }
    expect(() => readInlinePresenceExtension({ ...record, schemaVersion: 3 })).toThrow();
    const control = buildBlockIndex(pkg).byId.get('born-atom')!;
    const changed = applyTreeOp(pkg.parts.get(control.partName)!, {
      op: 'setContentControlProperties',
      controlId: control.controlId,
      lock: 'unlocked',
    });
    if (!changed.ok) throw Error(changed.reason);
    expect(readReviewGroups(withPart(pkg, changed.part)).success).toBe(false);
    const reservations = reviewGroupRevisionReservations(pkg);
    const rejected = rejectAllReviewGroups(editor, token(), context);
    if (!rejected.success) throw Error(rejected.error);
    expect(reviewGroupRevisionReservations(editor.surface!.session.currentPackage())).toEqual(
      reservations
    );
    const reopened = readOoxmlPackage(writeOoxmlPackage(editor.surface!.session.currentPackage()));
    if (!reopened.ok) throw Error(reopened.reason);
    expect(reviewGroupRevisionReservations(reopened.package)).toEqual(reservations);
  } finally {
    editor.destroy();
  }
});

test('a born atom paragraph beside a new application block keeps independent presence groups', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture('<w:p><w:r><w:t>Start</w:t></w:r></w:p>'),
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
  const token = () => readEditorRevision(editor, context.documentId, context.epoch)!;
  try {
    const before = native(editor.surface!.session.currentPackage());
    let count = 0;
    const staged = await recomposeOwnedProposals(
      editor,
      {
        revision: token(),
        targetMarkdown:
          'Start\n\nBorn ' +
          atoms.mention +
          '.\n\n<activity name="Task" duration="5">\n\nBody\n\n</activity>',
        ownership: [],
      },
      context,
      {
        target: { allocateBlockId: () => 'mixed-' + ++count },
        payloads: { presence: { proposalId: 'mixed-presence' } },
      }
    );
    if (!staged.success) throw Error(staged.error);
    const groups = readReviewGroups(editor.surface!.session.currentPackage());
    if (!groups.success) throw Error(groups.error);
    expect(groups.groups).toHaveLength(2);
    expect(groups.groups.every((group) => !group.readOnly && group.properties.length === 1)).toBe(
      true
    );
    const rejected = rejectAllReviewGroups(editor, token(), context);
    if (!rejected.success) throw Error(rejected.error);
    expect(native(editor.surface!.session.currentPackage())).toBe(before);
    expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(0);
  } finally {
    editor.destroy();
  }
});
