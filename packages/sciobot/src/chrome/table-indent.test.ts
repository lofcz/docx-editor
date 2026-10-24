import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  validateTreeOp,
} from '@docx-editor.dev/core/store';
import {
  tableIndentDocument,
  indentPart,
  indentTable,
  indentLeaf,
  indentValue,
} from '../../test/support/table-indent';
import { savedTableFingerprint } from '../../test/support/manual-table-controls';
import { setManualTableIndent, manualTableIndent, createTableIndentToolbar } from './table-indent';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';

for (const merged of [false, true])
  test(`table indentation preserves complete imported structure, exact geometry, reset and history (merge=${merged})`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: tableIndentDocument(0, merged),
    });
    const toolbar = createTableIndentToolbar(editor);
    try {
      const s = editor.surface!,
        point = { paragraphId: s.session.paragraphIds().at(-2)!, offset: 2 };
      s.setSelection({ anchor: point, head: point });
      const before = canonicalOoxmlFingerprint(s.session.part());
      const table = () =>
        s
          .layout()
          .pages.flatMap((p) => p.fragments)
          .find((f) => f.kind === 'table')!;
      const x = table().box.x,
        width = table().box.width;
      const originalIds = s.session.paragraphIds();
      for (const pixels of [24, 0, -12]) {
        expect(setManualTableIndent(editor, pixels).ok).toBe(true);
        expect(manualTableIndent(editor)).toBe(pixels);
        expect(indentValue(indentLeaf(s.session.part()), 'w')).toBe(String(pixels * 15));
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(
          canonicalOoxmlFingerprint(indentPart(pixels * 15, merged))
        );
        expect(table().box.x - x).toBeCloseTo(pixels * 0.75, 6);
        expect(table().box.width).toBe(width);
        expect(s.session.paragraphIds()).toEqual(originalIds);
        expect(s.state().selection).toEqual({ anchor: point, head: point });
        expect(await savedTableFingerprint(editor)).toEqual(
          canonicalOoxmlFingerprint(s.session.part())
        );
      }
      for (const twips of [0, 360, 0]) {
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(
          canonicalOoxmlFingerprint(indentPart(twips, merged))
        );
      }
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      for (const twips of [360, 0, -180]) {
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(
          canonicalOoxmlFingerprint(indentPart(twips, merged))
        );
      }
    } finally {
      toolbar.dispose();
      editor.destroy();
    }
  });
test('indent validation refuses fractional twips, invalid bounds, non-table selection and view mode without writes', () => {
  for (const mode of ['edit', 'view'] as const) {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: tableIndentDocument(0),
      mode,
    });
    try {
      const s = editor.surface!,
        point = { paragraphId: s.session.paragraphIds()[0]!, offset: 0 };
      s.setSelection({ anchor: point, head: point });
      const before = canonicalOoxmlFingerprint(s.session.part());
      for (const pixels of [NaN, Infinity, 2113, -2113, 0.1])
        expect(setManualTableIndent(editor, pixels).ok).toBe(false);
      if (mode === 'view') expect(setManualTableIndent(editor, 24).ok).toBe(false);
      const tail = { paragraphId: s.session.paragraphIds().at(-1)!, offset: 0 };
      s.setSelection({ anchor: tail, head: tail });
      expect(setManualTableIndent(editor, 24).ok).toBe(false);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      for (const indentTwips of [NaN, Infinity, 0.5, 31681, -31681])
        expect(
          validateTreeOp(s.session.part(), {
            op: 'setTableProperties',
            tableId: indentTable(s.session.part()).id,
            indentTwips,
          })
        ).toBe('invalidArgs');
    } finally {
      editor.destroy();
    }
  }
});
test('an omitted indent is created without rebuilding table contents', () => {
  const part = indentPart(),
    table = indentTable(part);
  const result = applyTreeOp(part, {
    op: 'setTableProperties',
    tableId: table.id,
    indentTwips: 360,
  });
  if (!result.ok) throw new Error(result.reason);
  expect(indentValue(indentLeaf(result.part), 'w')).toBe('360');
  expect(indentValue(indentLeaf(result.part), 'type')).toBe('dxa');
  if (table.kind === 'textValue') throw new Error('Missing table');
  const after = indentTable(result.part);
  if (after.kind === 'textValue') throw new Error('Missing table');
  expect(after.children.filter((n) => n.kind !== 'tableProperties')).toEqual(
    table.children.filter((n) => n.kind !== 'tableProperties')
  );
});
for (const decision of ['accept', 'reject'] as const)
  test(`suggested indentation uses native tracked property ${decision} and exact Word history`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: tableIndentDocument(0),
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      author: 'Reviewer',
    });
    try {
      expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      const s = editor.surface!,
        point = { paragraphId: s.session.paragraphIds()[0]!, offset: 1 };
      s.setSelection({ anchor: point, head: point });
      expect(setManualTableIndent(editor, 24).ok).toBe(true);
      expect(collectRevisionSites(s.session.part()).length).toBeGreaterThan(0);
      const pending = canonicalOoxmlFingerprint(s.session.part());
      expect(
        editor.exec({ type: 'resolveAllReviewChanges', action: decision, scope: 'document' }).ok
      ).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(
        canonicalOoxmlFingerprint(indentPart(decision === 'accept' ? 360 : 0))
      );
      expect(await savedTableFingerprint(editor)).toEqual(
        canonicalOoxmlFingerprint(s.session.part())
      );
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(pending);
    } finally {
      editor.destroy();
    }
  });
