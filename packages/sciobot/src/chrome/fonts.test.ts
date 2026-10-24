import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { fontLinkDocument, selectFontLink } from '../../test/support/font-link-document';
import {
  manualFontState,
  setManualFontFamily,
  setManualFontSize,
  FONT_SIZE_PIXELS,
  createFontToolbar,
} from './fonts';
test('all captured CSS font sizes retain exact native values, complete history and Word output', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fontLinkDocument(),
  });
  try {
    const s = editor.surface!;
    selectFontLink(editor, 0, 5);
    const before = canonicalOoxmlFingerprint(s.session.part());
    for (const size of FONT_SIZE_PIXELS) {
      expect(setManualFontSize(editor, size).ok).toBe(true);
      expect(manualFontState(editor).sizePixels).toBe(size);
      const after = canonicalOoxmlFingerprint(s.session.part()),
        saved = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!saved.ok) throw new Error(saved.reason);
      expect(
        canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
      ).toEqual(after);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    }
    selectFontLink(editor, 0, 5);
    expect(setManualFontSize(editor, 9).ok).toBe(true);
    expect(serializeOoxmlPart(s.session.part())).toContain('6.75pt');
  } finally {
    editor.destroy();
  }
});
test('family, size, mixed values, clear and caret typing retain native bold and text', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fontLinkDocument(),
  });
  const toolbar = createFontToolbar(editor);
  try {
    const s = editor.surface!;
    selectFontLink(editor, 0, 10);
    expect(manualFontState(editor).family).toBeNull();
    selectFontLink(editor, 0, 5);
    expect(setManualFontFamily(editor, 'Arial').ok).toBe(true);
    expect(manualFontState(editor).family).toBe('Arial');
    expect(setManualFontSize(editor, 17).ok).toBe(true);
    expect(manualFontState(editor).sizePixels).toBe(17);
    expect(s.formatting().bold).toBe(true);
    selectFontLink(editor, 0, 10);
    expect(manualFontState(editor).sizePixels).toBeNull();
    selectFontLink(editor, 0, 5);
    expect(setManualFontFamily(editor, null).ok).toBe(true);
    expect(setManualFontSize(editor, null).ok).toBe(true);
    expect(manualFontState(editor)).toEqual({ family: 'Calibri', sizePixels: 22 / 1.5 });
    selectFontLink(editor, 2);
    expect(setManualFontSize(editor, 17).ok).toBe(true);
    expect(manualFontState(editor).sizePixels).toBe(17);
    s.insertPlainText('X');
    s.flushPendingInput();
    expect(paragraphTextOf(s.session.part(), s.session.paragraphIds()[0]!)).toBe('AlXpha beta');
    expect(manualFontState(editor).sizePixels).toBe(17);
  } finally {
    toolbar.dispose();
    editor.destroy();
  }
});
test('invalid sizes and families refuse without native writes', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fontLinkDocument(),
  });
  try {
    selectFontLink(editor, 0, 5);
    const s = editor.surface!,
      before = canonicalOoxmlFingerprint(s.session.part());
    for (const size of [-1, 0, NaN, Infinity])
      expect(setManualFontSize(editor, size).ok).toBe(false);
    expect(setManualFontFamily(editor, 'bad"font').ok).toBe(false);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
  } finally {
    editor.destroy();
  }
});

test('clearing a caret font and size exposes the native inherited face before typing', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fontLinkDocument(),
  });
  try {
    selectFontLink(editor, 0, 5);
    expect(setManualFontFamily(editor, 'Arial').ok).toBe(true);
    expect(setManualFontSize(editor, 17).ok).toBe(true);
    selectFontLink(editor, 2);
    expect(setManualFontFamily(editor, null).ok).toBe(true);
    expect(setManualFontSize(editor, null).ok).toBe(true);
    expect(manualFontState(editor)).toEqual({ family: 'Calibri', sizePixels: 22 / 1.5 });
    const s = editor.surface!;
    s.insertPlainText('X');
    s.flushPendingInput();
    expect(manualFontState(editor)).toEqual({ family: 'Calibri', sizePixels: 22 / 1.5 });
  } finally {
    editor.destroy();
  }
});

test('range and caret font clears retain inherited character-style fonts and exact fractional sizes', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: fontLinkDocument(true),
  });
  try {
    selectFontLink(editor, 0, 5);
    expect(setManualFontFamily(editor, null).ok).toBe(true);
    expect(setManualFontSize(editor, null).ok).toBe(true);
    expect(manualFontState(editor)).toEqual({ family: 'Cambria', sizePixels: 9 });
    expect(setManualFontFamily(editor, 'Arial').ok).toBe(true);
    expect(setManualFontSize(editor, 17).ok).toBe(true);
    selectFontLink(editor, 2);
    expect(setManualFontFamily(editor, null).ok).toBe(true);
    expect(setManualFontSize(editor, null).ok).toBe(true);
    expect(manualFontState(editor)).toEqual({ family: 'Cambria', sizePixels: 9 });
    const s = editor.surface!;
    s.insertPlainText('X');
    s.flushPendingInput();
    expect(manualFontState(editor)).toEqual({ family: 'Cambria', sizePixels: 9 });
  } finally {
    editor.destroy();
  }
});
