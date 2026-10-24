import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor';
import { blankDocumentBytes } from '../blank-document';
import { canonicalOoxmlFingerprint } from '../../store/index';
test('a guarded native write captures its explicit post-edit selection for paint, undo and redo', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: blankDocumentBytes(),
  });
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!,
      point = { paragraphId: id, offset: 0 };
    s.setSelection({ anchor: point, head: point });
    const before = canonicalOoxmlFingerprint(s.session.part());
    const after = { anchor: { paragraphId: id, offset: 1 }, head: { paragraphId: id, offset: 2 } };
    expect(
      s.commitReviewOps(
        () =>
          s.session.applyTreeOps([{ op: 'insertText', paragraphId: id, offset: 0, text: 'abc' }]),
        'revision-resolve',
        () => after
      )
    ).toBeNull();
    expect(s.state().selection).toEqual(after);
    const written = canonicalOoxmlFingerprint(s.session.part());
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    expect(s.state().selection).toEqual({ anchor: point, head: point });
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(written);
    expect(s.state().selection).toEqual(after);
  } finally {
    editor.destroy();
  }
});
test('viewing refuses a guarded write before calling its native writer or selection resolver', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: blankDocumentBytes(),
    mode: 'view',
  });
  let calls = 0;
  try {
    const s = editor.surface!,
      before = canonicalOoxmlFingerprint(s.session.part());
    expect(
      s.commitReviewOps(
        () => {
          calls++;
          return { committed: true };
        },
        'revision-resolve',
        () => {
          calls++;
          return null;
        }
      )
    ).not.toBeNull();
    expect(calls).toBe(0);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
  } finally {
    editor.destroy();
  }
});
