import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { prepareNativeDocument } from '../native/prepare-document';
import reference from '../__tests__/fixtures/native-commands.json';
import { BASIC_MARKS, basicMarkState, toggleBasicMark } from './marks';
import { createMarksToolbar, installMarkShortcuts } from './marks-toolbar';
import { selectCellRectangle } from '../../../core/src/editor/__tests__/paginated-surface-fixtures';

function markDocument(body = '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p>') {
  const prepared = prepareNativeDocument(docxFixture(body));
  if (!prepared.success) throw new Error(prepared.error);
  return prepared.document;
}

for (const mark of BASIC_MARKS) {
  test(`${mark} matches captured on/off/caret marks, preserves selection, and saves native properties`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: markDocument(),
    });
    const surface = editor.surface!;
    const id = surface.session.paragraphIds()[0]!;
    const selection = {
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 5 },
    };
    try {
      const original = canonicalOoxmlFingerprint(surface.session.part());
      surface.setSelection(selection);
      for (const phase of ['on', 'off'] as const) {
        const fixture = reference.cases.find((c) => c.id === `mark.${mark}.${phase}`)!;
        expect(fixture).toBeDefined();
        expect(toggleBasicMark(editor, mark).ok).toBe(true);
        expect(basicMarkState(editor, mark)).toBe(
          fixture.after.marks?.[mark as 'bold'] ? 'on' : 'off'
        );
        expect(surface.state().selection).toEqual(selection);
        if (phase === 'on') {
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          const xml = serializeOoxmlPart(saved.package.parts.get(saved.package.mainDocumentPart)!);
          const expected = {
            bold: '<w:b',
            italic: '<w:i',
            underline: '<w:u',
            strikethrough: '<w:strike',
            subscript: 'w:val="subscript"',
            superscript: 'w:val="superscript"',
            code: 'w:val="SciobotCode"',
          }[mark];
          expect(xml).toContain(expected);
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(canonicalOoxmlFingerprint(surface.session.part())).toEqual(original);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          surface.setSelection({ ...selection, head: { paragraphId: id, offset: 10 } });
          expect(basicMarkState(editor, mark)).toBe('mixed');
          surface.setSelection(selection);
        }
      }
      const caret = { paragraphId: id, offset: 10 };
      surface.setSelection({ anchor: caret, head: caret });
      expect(toggleBasicMark(editor, mark).ok).toBe(true);
      expect(basicMarkState(editor, mark)).toBe('on');
      surface.insertPlainText('!');
      surface.flushPendingInput();
      surface.setSelection({ anchor: caret, head: { ...caret, offset: 11 } });
      expect(basicMarkState(editor, mark)).toBe('on');
    } finally {
      editor.destroy();
    }
  });
}

test('code applies across paragraphs in one undo and caret toggles return to ordinary text', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: markDocument(
      '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>Beta</w:t></w:r></w:p>'
    ),
  });
  try {
    const s = editor.surface!,
      ids = s.session.paragraphIds();
    const selection = {
      anchor: { paragraphId: ids[1]!, offset: 4 },
      head: { paragraphId: ids[0]!, offset: 0 },
    };
    const before = canonicalOoxmlFingerprint(s.session.part());
    s.setSelection(selection);
    expect(toggleBasicMark(editor, 'code').ok).toBe(true);
    expect(basicMarkState(editor, 'code')).toBe('on');
    expect(s.state().selection).toEqual(selection);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    const at = { paragraphId: ids[0]!, offset: 0 };
    s.setSelection({ anchor: at, head: at });
    expect(toggleBasicMark(editor, 'code').ok).toBe(true);
    expect(basicMarkState(editor, 'code')).toBe('on');
    expect(toggleBasicMark(editor, 'code').ok).toBe(true);
    expect(basicMarkState(editor, 'code')).toBe('off');
    s.insertPlainText('Plain');
    s.flushPendingInput();
    expect(serializeOoxmlPart(s.session.part())).not.toContain('w:val="SciobotCode"');
  } finally {
    editor.destroy();
  }
});

test('caret code removes and restores an imported style without changing unrelated run properties', () => {
  const source = prepareNativeDocument(
    docxFixture(
      '<w:p><w:r><w:rPr><w:rStyle w:val="Imported"/><w:lang w:val="cs-CZ"/></w:rPr><w:t>Alpha</w:t></w:r></w:p>',
      {
        styles:
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="character" w:styleId="Imported"><w:name w:val="Imported"/><w:rPr><w:b/></w:rPr></w:style></w:styles>',
      }
    )
  );
  if (!source.success) throw new Error(source.error);
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: source.document,
  });
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!;
    s.setSelection({
      anchor: { paragraphId: id, offset: 5 },
      head: { paragraphId: id, offset: 5 },
    });
    const before = canonicalOoxmlFingerprint(s.session.part());
    expect(toggleBasicMark(editor, 'code').ok).toBe(true);
    expect<unknown>(s.state().pendingFormat).toEqual([
      { localName: 'rStyle', attributes: { val: 'SciobotCode__Imported' } },
    ]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(s.state().pendingFormat).toBeNull();
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(basicMarkState(editor, 'code')).toBe('on');
    expect(toggleBasicMark(editor, 'code').ok).toBe(true);
    expect<unknown>(s.state().pendingFormat).toEqual([
      { localName: 'rStyle', attributes: { val: 'Imported' } },
    ]);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    s.insertPlainText('!');
    s.flushPendingInput();
    const xml = serializeOoxmlPart(s.session.part());
    expect(xml).toContain('w:val="Imported"');
    expect(xml).toContain('w:val="cs-CZ"');
    expect(xml).not.toContain('w:val="SciobotCode__Imported"');
    expect(s.formatting().bold).toBe(true);
  } finally {
    editor.destroy();
  }
});

test('rectangular code selection changes selected cells without formatting cells between them', () => {
  const body =
    '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    ['A', 'B']
      .map(
        (row) =>
          '<w:tr>' +
          [1, 2]
            .map(
              (col) =>
                `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${row}${col}</w:t></w:r></w:p></w:tc>`
            )
            .join('') +
          '</w:tr>'
      )
      .join('') +
    '</w:tbl><w:p/>';
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: markDocument(body),
  });
  try {
    const s = editor.surface!,
      ids = s.session.paragraphIds();
    selectCellRectangle(s, { row: 0, column: 0 }, { row: 1, column: 0 });
    expect(toggleBasicMark(editor, 'code').ok).toBe(true);
    for (const [index, expected] of ['on', 'off', 'on', 'off'].entries()) {
      s.setSelection({
        anchor: { paragraphId: ids[index]!, offset: 0 },
        head: { paragraphId: ids[index]!, offset: 2 },
      });
      expect(basicMarkState(editor, 'code')).toBe(expected as 'on' | 'off');
    }
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(serializeOoxmlPart(s.session.part())).not.toContain('w:val="SciobotCode"');
  } finally {
    editor.destroy();
  }
});

test('toolbar reports mixed state and shortcuts dispose without affecting foreign controls', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: markDocument(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Alpha</w:t></w:r><w:r><w:t> beta</w:t></w:r></w:p>'
    ),
  });
  const toolbar = createMarksToolbar(editor),
    stop = installMarkShortcuts(editor);
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!;
    s.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 10 },
    });
    expect(toolbar.element.querySelector('button')!.getAttribute('aria-pressed')).toBe('mixed');
    const editable = s.hostElement().querySelector('[contenteditable="true"]')!;
    const key = () =>
      new KeyboardEvent('keydown', { key: 'e', ctrlKey: true, bubbles: true, cancelable: true });
    const event = key();
    editable.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);
    expect(basicMarkState(editor, 'code')).toBe('on');
    stop();
    const after = key();
    editable.dispatchEvent(after);
    expect(basicMarkState(editor, 'code')).toBe('on');
  } finally {
    stop();
    toolbar.dispose();
    editor.destroy();
    container.remove();
  }
});
