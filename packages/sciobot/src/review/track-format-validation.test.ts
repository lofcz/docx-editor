import { expect, test } from 'bun:test';
import { readOoxmlPackage } from '@docx-editor.dev/core/store';
import { propertyEditor } from '../../test/support/property-execution';
import { listDocument, listParagraph } from '../../test/support/list-document';
import type { CanonicalEditPlan } from '../diff/plan';
import { trackFormatPlan } from './track-format';
const context = {
  documentId: 'format-validation',
  epoch: 'mount',
  author: 'Alice',
  date: '2026-10-23T00:00:00Z',
};
const cases: [string, unknown][] = [
  ['null', null],
  ['undefined', undefined],
  ['missing operations', {}],
  ['null operations', { operations: null }],
  ['null operation', { operations: [null] }],
];
for (const [name, input] of cases)
  test(`malformed ${name} returns a failure without document or history writes`, () => {
    const read = readOoxmlPackage(listDocument(listParagraph('Original')));
    if (!read.ok) throw Error(read.reason);
    const editor = propertyEditor(read.package);
    try {
      const session = editor.surface!.session,
        before = session.currentPackage(),
        revision = session.packageRevision();
      expect(trackFormatPlan(editor, input as CanonicalEditPlan, context)).toMatchObject({
        success: false,
        code: 'invalid-input',
      });
      expect(session.currentPackage()).toBe(before);
      expect(session.packageRevision()).toBe(revision);
      expect(editor.exec({ type: 'undo' }).ok).toBe(false);
    } finally {
      editor.destroy();
    }
  });
