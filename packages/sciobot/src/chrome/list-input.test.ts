import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, paragraphTextOf } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { applyManualListInput, handleManualListKey } from './list-input';
import {
  manualListMarker,
  manualListState,
  setManualListChecked,
  adjustManualListLevel,
} from './lists';
const context = { documentId: 'manual-list-input', epoch: 'mount' };
for (const prefix of ['-', '*', '1.', '1)', '[]', '[x]'])
  test(`captured ${prefix} prefix becomes a native list with exact one-step history`, () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(`<w:p><w:r><w:t>${prefix}</w:t></w:r></w:p>`),
    });
    try {
      const s = editor.surface!,
        id = s.session.paragraphIds()[0]!,
        point = { paragraphId: id, offset: prefix.length };
      s.setSelection({ anchor: point, head: point });
      const before = canonicalOoxmlFingerprint(s.session.part());
      expect(applyManualListInput(editor, context)).toBe(true);
      const checkbox = prefix.startsWith('['),
        state = manualListState(editor, context)!;
      expect(state.checkbox?.checked ?? false).toBe(prefix === '[x]');
      expect(paragraphTextOf(s.session.part(), id)).toBe(
        checkbox ? (prefix === '[x]' ? '☒' : '☐') : ''
      );
      if (!checkbox)
        expect(manualListMarker(editor, id)!.text).toBe(
          prefix === '-' || prefix === '*' ? '•' : '1.'
        );
      const after = canonicalOoxmlFingerprint(s.session.part());
      expect(s.state().selection.head.offset).toBe(checkbox ? 1 : 0);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
      expect(s.state().selection.head.offset).toBe(checkbox ? 1 : 0);
    } finally {
      editor.destroy();
    }
  });
test('checklist Enter retains nesting, creates an unchecked item, outdents empty items and undoes as one gesture', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(
      '<w:p><w:r><w:t>[]</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
    ),
  });
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!,
      point = { paragraphId: id, offset: 2 };
    s.setSelection({ anchor: point, head: point });
    expect(applyManualListInput(editor, context)).toBe(true);
    expect(setManualListChecked(editor, true, context).success).toBe(true);
    expect(adjustManualListLevel(editor, 'increase', context).success).toBe(true);
    s.insertPlainText('Task');
    s.flushPendingInput();
    const before = canonicalOoxmlFingerprint(s.session.part());
    expect(handleManualListKey(editor, 'Enter', context)).toBe(true);
    const next = s.state().selection.head.paragraphId;
    expect(next).not.toBe(id);
    expect(manualListState(editor, context)!.checkbox).toMatchObject({ checked: false, level: 1 });
    expect(s.state().selection.head.offset).toBe(1);
    const after = canonicalOoxmlFingerprint(s.session.part());
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
    expect(handleManualListKey(editor, 'Enter', context)).toBe(true);
    expect(manualListState(editor, context)!.checkbox?.level).toBe(0);
    expect(handleManualListKey(editor, 'Backspace', context)).toBe(true);
    expect(manualListState(editor, context)!.checkbox).toBeNull();
    expect(paragraphTextOf(s.session.part(), next)).toBe('');
  } finally {
    editor.destroy();
  }
});
