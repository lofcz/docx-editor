import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  directParagraphProperties,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import reference from '../__tests__/fixtures/native-commands.json';
import {
  LINE_HEIGHT_CHOICES,
  lineHeightState,
  setLineHeight,
  createLineHeightToolbar,
} from './line-height';

const body =
  '<w:p><w:pPr><w:spacing w:line="251" w:lineRule="exact" w:before="123" w:after="456"/><w:keepNext/></w:pPr><w:r><w:t>Imported</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:spacing w:before="80" w:after="100"/><w:ind w:left="321"/></w:pPr><w:r><w:t>Alpha beta</w:t></w:r></w:p>';
for (const value of LINE_HEIGHT_CHOICES)
  test(`line height ${value} matches captured choice, preserves exact neighbor, selection, undo, and reopen`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(body),
    });
    try {
      const s = editor.surface!,
        ids = s.session.paragraphIds();
      const original = canonicalOoxmlFingerprint(s.session.part());
      const imported = directParagraphProperties(s.session.part(), ids[0]!);
      const fixture = reference.cases.find((c) => c.id === `lineHeight.${value}`)!;
      expect(fixture).toBeDefined();
      // Plate removes its default 1.5 property; Word must write the effective multiple.
      const captured = (fixture.after.value[0] as { lineHeight?: number }).lineHeight;
      if (value === 1.5) expect(captured).toBeUndefined();
      expect(captured ?? 1.5).toBe(value);
      const selection = {
        anchor: { paragraphId: ids[1]!, offset: 5 },
        head: { paragraphId: ids[1]!, offset: 0 },
      };
      s.setSelection(selection);
      expect(setLineHeight(editor, value).ok).toBe(true);
      expect(s.state().selection).toEqual(selection);
      expect(lineHeightState(editor)).toEqual({ rule: 'multiple', value });
      const props = directParagraphProperties(s.session.part(), ids[1]!);
      expect(props.find((p) => p.localName === 'spacing')?.attributes).toEqual({
        before: '80',
        after: '100',
        line: String(value * 240),
        lineRule: 'auto',
      });
      expect(props.find((p) => p.localName === 'ind')?.attributes?.left).toBe('321');
      expect(directParagraphProperties(s.session.part(), ids[0]!)).toEqual(imported);
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      const part = reopened.package.parts.get(reopened.package.mainDocumentPart)!;
      expect(canonicalOoxmlFingerprint(part)).toEqual(canonicalOoxmlFingerprint(s.session.part()));
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(original);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(lineHeightState(editor)).toEqual({ rule: 'multiple', value });
    } finally {
      editor.destroy();
    }
  });

test('line-height toolbar reads imported exact and mixed selections without normalizing either', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(body),
  });
  const toolbar = createLineHeightToolbar(editor);
  try {
    const s = editor.surface!,
      ids = s.session.paragraphIds();
    const before = canonicalOoxmlFingerprint(s.session.part());
    s.setSelection({
      anchor: { paragraphId: ids[0]!, offset: 0 },
      head: { paragraphId: ids[0]!, offset: 2 },
    });
    expect(lineHeightState(editor)).toEqual({ rule: 'exact', value: 12.55 });
    expect(toolbar.element.querySelector('output')!.textContent).toBe('Exactly 12.55 pt');
    s.setSelection({
      anchor: { paragraphId: ids[0]!, offset: 0 },
      head: { paragraphId: ids[1]!, offset: 2 },
    });
    expect(toolbar.element.querySelector('output')!.textContent).toBe('Mixed');
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    expect(setLineHeight(editor, 4 as 1).ok).toBe(false);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
  } finally {
    toolbar.dispose();
    editor.destroy();
  }
});
