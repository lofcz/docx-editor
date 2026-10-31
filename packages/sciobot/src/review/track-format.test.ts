import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { propertyEditor } from '../../test/support/property-execution';
import { listDocument, listParagraph, checklistParagraph } from '../../test/support/list-document';
import { readEditorRevision } from '../projection/revision-cache';
import { planTargetDocument } from '../diff/plan-target';
import { readWorkingView } from './working-view';
import { planRunFormat } from '../diff/run-format';
import { planParagraphFormat } from '../diff/paragraph-format';
import { SCIOBOT_CODE_STYLE_ID } from '../native/inline-code';
import type { OoxmlProperty } from '@docx-editor.dev/core/store';
import { planListMembership } from '../diff/lists';
import type { ListTarget } from '../diff/list-types';
import { applyDirectPlan } from '../ai/apply-direct';
import { trackFormatPlan } from './track-format';

const context = {
  documentId: 'format-review',
  epoch: 'mount',
  author: 'Alice',
  date: '2026-10-23T00:00:00Z',
};
const cases: [string, string, ListTarget][] = [
  [
    'existing',
    listParagraph('Item'),
    { kind: 'ordered', level: 1, instance: { kind: 'existing', numId: '1' } },
  ],
  [
    'new',
    listParagraph('Item'),
    { kind: 'ordered', level: 2, instance: { kind: 'new', key: 'review', start: 5 } },
  ],
  [
    'bullet',
    listParagraph('Item', '1'),
    { kind: 'bullet', level: 0, instance: { kind: 'new', key: 'review' } },
  ],
  ['remove', listParagraph('Item', '1', 2), { kind: 'none' }],
  [
    'checkbox-to-ordered',
    checklistParagraph(true, 1),
    { kind: 'ordered', level: 1, instance: { kind: 'new', key: 'review', start: 3 } },
  ],
  [
    'checkbox-to-bullet',
    checklistParagraph(false),
    { kind: 'bullet', level: 0, instance: { kind: 'existing', numId: '2' } },
  ],
];
for (const [name, body, target] of cases)
  for (const action of ['accept', 'reject'] as const)
    test(`ordinary list ${name} supports ${action}, history, and saved reopen`, async () => {
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
          [
            {
              paragraphId: session.paragraphIds()[0]!,
              target,
            },
          ]
        );
        if (!result.success) throw Error(result.error);
        return result.plan;
      };
      const fingerprint = (value: typeof editor) =>
        canonicalOoxmlFingerprint(value.surface!.session.part());
      try {
        const before = fingerprint(editor);
        const applied = applyDirectPlan(direct, plan(direct), context);
        if (!applied.success) throw Error(applied.error);
        const expected = fingerprint(direct);
        const proposed = trackFormatPlan(editor, plan(editor), context);
        if (!proposed.success) throw Error(proposed.error);
        const pending = fingerprint(editor);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(pending);
        expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(action === 'accept' ? expected : before);
        const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!opened.ok) throw Error(opened.reason);
        expect(
          canonicalOoxmlFingerprint(opened.package.parts.get(opened.package.mainDocumentPart)!)
        ).toBe(action === 'accept' ? expected : before);
      } finally {
        editor.destroy();
        direct.destroy();
      }
    });

const marks: [string, OoxmlProperty][] = [
  ['bold', { localName: 'b' }],
  ['italic', { localName: 'i' }],
  ['underline', { localName: 'u', attributes: { val: 'single' } }],
  ['strike', { localName: 'strike' }],
  ['subscript', { localName: 'vertAlign', attributes: { val: 'subscript' } }],
  ['superscript', { localName: 'vertAlign', attributes: { val: 'superscript' } }],
  ['code', { localName: 'rStyle', attributes: { val: SCIOBOT_CODE_STYLE_ID } }],
];
for (const [name, property] of [...marks, ['heading', null] as const])
  for (const action of ['accept', 'reject'] as const)
    test(`native ${name} ${action} retains exact properties through history and reopen`, async () => {
      const read = readOoxmlPackage(listDocument(listParagraph('Before Item After')));
      if (!read.ok) throw Error(read.reason);
      const editor = propertyEditor(read.package),
        direct = propertyEditor(read.package);
      const plan = async (value: typeof editor) => {
        const session = value.surface!.session;
        const source = {
          part: session.part(),
          paragraphId: session.paragraphIds()[0]!,
          revision: readEditorRevision(value, context.documentId, context.epoch)!,
        };
        if (name === 'code') {
          const view = await readWorkingView(session, source.revision);
          if (!view.success) throw Error(view.error);
          const result = planTargetDocument(
            {
              revision: source.revision,
              pkg: session.currentPackage(),
              projection: view.view.projection,
            },
            'Before `Item` After'
          );
          if (!result.success) throw Error(result.error);
          return result.plan;
        }
        const result = property
          ? planRunFormat(
              { ...source, start: 7, end: 11, displayMode: 'proposed' },
              { properties: [property] }
            )
          : planParagraphFormat(source, {
              properties: [{ localName: 'pStyle', attributes: { val: 'Heading2' } }],
            });
        if (!result.success) throw Error(result.error);
        return result.plan;
      };
      const fingerprint = (value: typeof editor) =>
        canonicalOoxmlFingerprint(value.surface!.session.part());
      try {
        const before = fingerprint(editor);
        // Native range formatting partitions a run at the selected boundaries. The partitions
        // keep source provenance, so a rejection restores the one original run exactly,
        // without an empty property container.
        const rejected = before;
        const applied = applyDirectPlan(direct, await plan(direct), context);
        if (!applied.success) throw Error(applied.error);
        const expected = fingerprint(direct);
        const tracked = trackFormatPlan(editor, await plan(editor), context);
        if (!tracked.success) throw Error(tracked.error);
        const pending = fingerprint(editor);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(pending);
        expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        expect(fingerprint(editor)).toBe(action === 'accept' ? expected : rejected);
        const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!opened.ok) throw Error(opened.reason);
        expect(
          canonicalOoxmlFingerprint(opened.package.parts.get(opened.package.mainDocumentPart)!)
        ).toBe(action === 'accept' ? expected : rejected);
      } finally {
        editor.destroy();
        direct.destroy();
      }
    });
