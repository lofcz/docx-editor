import { expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import type { CanonicalPrimitiveJournal } from '@docx-editor.dev/core/collaboration/replication';
import { canonicalOoxmlFingerprint, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { propertyEditor } from '../../test/support/property-execution';
import { listDocument, listParagraph, checklistParagraph } from '../../test/support/list-document';
import { readEditorRevision } from '../projection/revision-cache';
import { planListMembership } from '../diff/lists';
import type { ListTarget } from '../diff/list-types';
import { applyDirectPlan } from '../ai/apply-direct';
import { trackChecklistPlan } from './track-checklists';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
const context = {
  documentId: 'checklist-review',
  epoch: 'mount',
  author: 'Alice',
  date: '2026-10-22T09:00:00Z',
};
const cases: [string, string, ListTarget][] = [
  ['create', listParagraph('Task'), { kind: 'checklist', level: 0, checked: false }],
  ['check', checklistParagraph(false), { kind: 'checklist', level: 0, checked: true }],
  ['nest', checklistParagraph(false), { kind: 'checklist', level: 2, checked: false }],
  ['remove', checklistParagraph(false), { kind: 'none' }],
];
for (const [name, body, target] of cases)
  for (const action of ['accept', 'reject'] as const)
    test(`tracked checklist ${name}: exact ${action}, history, saved decision`, async () => {
      const read = readOoxmlPackage(listDocument(body));
      if (!read.ok) throw Error(read.reason);
      const editor = propertyEditor(read.package),
        direct = propertyEditor(read.package);
      const plan = (value: typeof editor) => {
        const session = value.surface!.session;
        const result = planListMembership(
          {
            pkg: session.currentPackage(),
            partName: session.part().name,
            revision: readEditorRevision(value, context.documentId, context.epoch)!,
          },
          [{ paragraphId: session.paragraphIds()[0]!, target }]
        );
        if (!result.success) throw Error(result.error);
        return result.plan;
      };
      const fingerprint = (value: typeof editor) =>
        canonicalOoxmlFingerprint(value.surface!.session.part());
      const projection = async (value: typeof editor, original: boolean) => {
        const result = await (original ? readCommittedView : readWorkingView)(
          value.surface!.session,
          readEditorRevision(value, context.documentId, context.epoch)!
        );
        if (!result.success) throw Error(result.error);
        return result.view.projection.text;
      };
      try {
        const before = fingerprint(editor),
          original = await projection(editor, false);
        const originalControlIds = contentControlsIn(editor.surface!.session.part().root).map(
          (entry) => entry.node.id
        );
        const originalApplicationIds = contentControlsIn(editor.surface!.session.part().root).map(
          (entry) => contentControlPropertiesOf(entry.node).id
        );
        const originalControls = contentControlsIn(editor.surface!.session.part().root).map(
          (entry) => canonicalOoxmlFingerprint(entry.node)
        );
        const applied = applyDirectPlan(direct, plan(direct), context);
        if (!applied.success) throw Error(applied.error);
        if (name === 'nest')
          expect(
            contentControlsIn(direct.surface!.session.part().root).map((entry) =>
              canonicalOoxmlFingerprint(entry.node)
            )
          ).toEqual(originalControls);
        const expected = fingerprint(direct),
          proposed = await projection(direct, false);
        const journals: CanonicalPrimitiveJournal[] = [];
        const port = editor.surface!.session.collaborationPort(context.documentId);
        const stop = port.observePrimitiveJournal((journal) => journals.push(journal));
        const result = trackChecklistPlan(editor, plan(editor), context);
        port.flushPendingJournals();
        stop();
        expect(journals).toHaveLength(1);
        expect(
          journals[0]!.effects.filter(
            (effect) => effect.kind === 'spliceText' && effect.deleteCount > 0
          )
        ).toEqual([]);
        if (!result.success) throw Error(result.error);
        expect(await projection(editor, true)).toBe(original);
        expect(await projection(editor, false)).toBe(proposed);
        const pending = fingerprint(editor);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(pending);
        const bytes = new Uint8Array(await editor.save());
        const xml = strFromU8(unzipSync(bytes)['word/document.xml']!);
        if (name === 'check' || name === 'remove') {
          const deletion = xml.match(/<w:del\b[^>]*>([\s\S]*?)<\/w:del>/)![1]!;
          expect(deletion).toContain('<w:delText');
          expect(deletion).not.toMatch(/<w:t(?:[ >])/);
        }
        const reopened = readOoxmlPackage(bytes);
        if (!reopened.ok) throw Error(reopened.reason);
        const saved = propertyEditor(reopened.package);
        try {
          expect(await projection(saved, true)).toBe(original);
          expect(await projection(saved, false)).toBe(proposed);
          expect(saved.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
          expect(fingerprint(saved)).toBe(action === 'accept' ? expected : before);
        } finally {
          saved.destroy();
        }
        expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(action === 'accept' ? expected : before);
        if (action === 'reject')
          expect(
            contentControlsIn(editor.surface!.session.part().root).map((entry) => entry.node.id)
          ).toEqual(originalControlIds);
        if (name === 'check' || name === 'nest' || action === 'reject')
          expect(
            contentControlsIn(editor.surface!.session.part().root).map(
              (entry) => contentControlPropertiesOf(entry.node).id
            )
          ).toEqual(originalApplicationIds);
      } finally {
        editor.destroy();
        direct.destroy();
      }
    });

import {
  contentControlsIn,
  contentControlPropertiesOf,
  findNode,
} from '@docx-editor.dev/core/store';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';

test('later checklist proposals preserve older decisions, native control IDs, and unrelated forms after saving', async () => {
  const body =
    checklistParagraph(false) +
    checklistParagraph(false).replace('w:val="8"', 'w:val="9"') +
    checklistParagraph(true, 0, 'ordinary-form').replace('w:val="8"', 'w:val="10"');
  const read = readOoxmlPackage(listDocument(body));
  if (!read.ok) throw Error(read.reason);
  const editor = propertyEditor(read.package),
    expected = propertyEditor(read.package);
  const token = (value: typeof editor) =>
    readEditorRevision(value, context.documentId, context.epoch)!;
  const planned = (value: typeof editor, index: number) => {
    const session = value.surface!.session;
    const result = planListMembership(
      { pkg: session.currentPackage(), partName: session.part().name, revision: token(value) },
      [
        {
          paragraphId: session.paragraphIds()[index]!,
          target: { kind: 'checklist', level: index, checked: true },
        },
      ]
    );
    if (!result.success) throw Error(result.error);
    return result.plan;
  };
  try {
    const part = editor.surface!.session.part();
    const controls = contentControlsIn(part.root);
    const originalIds = controls.map((control) => control.node.id);
    const foreign = canonicalOoxmlFingerprint(controls[2]!.node);
    const first = trackChecklistPlan(editor, planned(editor, 0), context);
    if (!first.success) throw Error(first.error);
    const second = trackChecklistPlan(editor, planned(editor, 1), context);
    if (!second.success) throw Error(second.error);
    expect(first.ownership).toHaveLength(1);
    expect(second.ownership).toHaveLength(1);
    expect(first.ownership[0]!.groupId).not.toBe(second.ownership[0]!.groupId);
    const current = editor.surface!.session.part();
    for (const id of originalIds) expect(findNode(current, id)).not.toBeNull();
    expect(canonicalOoxmlFingerprint(findNode(current, originalIds[2]!)!)).toBe(foreign);
    const applied = applyDirectPlan(expected, planned(expected, 1), context);
    if (!applied.success) throw Error(applied.error);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw Error(saved.reason);
    const reopened = propertyEditor(saved.package);
    try {
      const result = rejectReviewGroup(
        reopened,
        first.ownership[0]!.groupId,
        token(reopened),
        context
      );
      if (!result.success) throw Error(result.error);
      expect(reopened.getReviewItems().length).toBeGreaterThan(0);
      const accept = acceptReviewGroup(
        reopened,
        second.ownership[0]!.groupId,
        token(reopened),
        context
      );
      if (!accept.success) throw Error(accept.error);
      expect(canonicalOoxmlFingerprint(reopened.surface!.session.part())).toBe(
        canonicalOoxmlFingerprint(expected.surface!.session.part())
      );
      expect(
        contentControlsIn(reopened.surface!.session.part().root).map(
          (control) => contentControlPropertiesOf(control.node).tag
        )
      ).toEqual(['scio:list-checkbox', 'scio:list-checkbox', 'ordinary-form']);
    } finally {
      reopened.destroy();
    }
  } finally {
    editor.destroy();
    expected.destroy();
  }
});

for (const failure of ['stale', 'locked'] as const)
  test(`tracked checklist ${failure} refusal preserves package and history`, () => {
    const body = checklistParagraph(false).replace(
      '<w:sdtPr>',
      failure === 'locked' ? '<w:sdtPr><w:lock w:val="contentLocked"/>' : '<w:sdtPr>'
    );
    const read = readOoxmlPackage(listDocument(body));
    if (!read.ok) throw Error(read.reason);
    const editor = propertyEditor(read.package);
    try {
      const session = editor.surface!.session;
      const paragraphId = session.paragraphIds()[0]!;
      const planned = planListMembership(
        {
          pkg: session.currentPackage(),
          partName: session.part().name,
          revision: readEditorRevision(editor, context.documentId, context.epoch)!,
        },
        [{ paragraphId, target: { kind: 'checklist', level: 0, checked: true } }]
      );
      if (!planned.success) throw Error(planned.error);
      if (failure === 'stale')
        expect(
          editor.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId, offset: 5, text: '!' },
          ]).rejected
        ).toBe(false);
      const before = canonicalOoxmlFingerprint(session.part());
      const snapshot = editor.snapshot();
      const result = trackChecklistPlan(editor, planned.plan, context);
      expect(result.success).toBe(false);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
      expect(editor.snapshot().canUndo).toBe(snapshot.canUndo);
      expect(editor.snapshot().canRedo).toBe(snapshot.canRedo);
    } finally {
      editor.destroy();
    }
  });

for (const firstKind of ['nest', 'check'] as const)
  for (const firstAction of ['accept', 'reject'] as const)
    for (const secondAction of ['accept', 'reject'] as const)
      test(`later checklist changes preserve earlier ${firstKind} decisions: ${firstAction}/${secondAction}`, async () => {
        const read = readOoxmlPackage(listDocument(checklistParagraph(false)));
        if (!read.ok) throw Error(read.reason);
        const editor = propertyEditor(read.package),
          direct = propertyEditor(read.package);
        const token = (value: typeof editor) =>
          readEditorRevision(value, context.documentId, context.epoch)!;
        const plan = (value: typeof editor, level: number, checked: boolean) => {
          const session = value.surface!.session;
          const planned = planListMembership(
            {
              pkg: session.currentPackage(),
              partName: session.part().name,
              revision: token(value),
            },
            [
              {
                paragraphId: session.paragraphIds()[0]!,
                target: { kind: 'checklist', level, checked },
              },
            ]
          );
          if (!planned.success) throw Error(planned.error);
          return planned.plan;
        };
        try {
          const first = trackChecklistPlan(
            editor,
            plan(editor, firstKind === 'nest' ? 2 : 0, firstKind === 'check'),
            context
          );
          if (!first.success) throw Error(first.error);
          const second = trackChecklistPlan(editor, plan(editor, 2, true), context);
          if (!second.success) throw Error(second.error);
          if (firstAction === 'accept' || secondAction === 'accept') {
            const level = (firstKind === 'nest' ? firstAction : secondAction) === 'accept' ? 2 : 0;
            const checked = (firstKind === 'check' ? firstAction : secondAction) === 'accept';
            const applied = applyDirectPlan(direct, plan(direct, level, checked), context);
            if (!applied.success) throw Error(applied.error);
          }
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw Error(saved.reason);
          const reopened = propertyEditor(saved.package);
          try {
            const rejected = (firstAction === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
              reopened,
              first.ownership[0]!.groupId,
              token(reopened),
              context
            );
            if (!rejected.success) throw Error(rejected.error);
            const accepted = (secondAction === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
              reopened,
              second.ownership[0]!.groupId,
              token(reopened),
              context
            );
            if (!accepted.success) throw Error(accepted.error);
            expect(canonicalOoxmlFingerprint(reopened.surface!.session.part())).toBe(
              canonicalOoxmlFingerprint(direct.surface!.session.part())
            );
          } finally {
            reopened.destroy();
          }
        } finally {
          editor.destroy();
          direct.destroy();
        }
      });
