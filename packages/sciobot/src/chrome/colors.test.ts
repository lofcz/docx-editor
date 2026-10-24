import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import reference from '../__tests__/fixtures/native-commands.json';
import { setTextColor, textColorState, createColorToolbar } from './colors';
import { spansInSelection } from '@docx-editor.dev/core/layout';

for (const channel of ['foreground', 'background'] as const)
  for (const color of ['#CC0000', '#FFFF00', '#17A2B8'])
    test(`${channel} ${color} applies and clears independently, preserves native selection and themes, and saves`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: docxFixture(
          '<w:p><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1" w:themeTint="33"/><w:shd w:val="clear" w:fill="DDEEFF" w:themeFill="accent2" w:themeFillShade="66"/></w:rPr><w:t>Alpha beta</w:t></w:r></w:p>'
        ),
      });
      try {
        const s = editor.surface!,
          id = s.session.paragraphIds()[0]!;
        const selection = {
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 5 },
        };
        s.setSelection(selection);
        const original = canonicalOoxmlFingerprint(s.session.part());
        const other = textColorState(
          editor,
          channel === 'foreground' ? 'background' : 'foreground'
        );
        expect(setTextColor(editor, channel, color).ok).toBe(true);
        expect(s.state().selection).toEqual(selection);
        expect(textColorState(editor, channel)).toEqual({ value: color, mixed: false });
        expect(
          textColorState(editor, channel === 'foreground' ? 'background' : 'foreground')
        ).toEqual(other);
        expect(serializeOoxmlPart(s.session.part())).toContain('w:themeColor="accent1"');
        expect(serializeOoxmlPart(s.session.part())).toContain('w:themeFill="accent2"');
        const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!saved.ok) throw new Error(saved.reason);
        expect(
          canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
        ).toEqual(canonicalOoxmlFingerprint(s.session.part()));
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(original);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        s.setSelection({ ...selection, head: { paragraphId: id, offset: 10 } });
        expect(textColorState(editor, channel).mixed).toBe(true);
        s.setSelection(selection);
        expect(setTextColor(editor, channel, null).ok).toBe(true);
        expect(
          textColorState(editor, channel === 'foreground' ? 'background' : 'foreground')
        ).toEqual(other);
        const property = channel === 'foreground' ? 'color' : 'shd';
        const xml = serializeOoxmlPart(s.session.part());
        expect(xml).not.toContain(`<w:${property}/>`);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(textColorState(editor, channel).value).toBe(color);
      } finally {
        editor.destroy();
      }
    });

for (const channel of ['foreground', 'background'] as const)
  test(`${channel} caret colors and clear preferences survive typing and local undo/redo`, () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture('<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>'),
    });
    try {
      const s = editor.surface!,
        id = s.session.paragraphIds()[0]!,
        at = { paragraphId: id, offset: 5 };
      s.setSelection({ anchor: at, head: at });
      const original = canonicalOoxmlFingerprint(s.session.part());
      expect(setTextColor(editor, channel, '#17A2B8').ok).toBe(true);
      expect(textColorState(editor, channel).value).toBe('#17A2B8');
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(original);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(s.state().pendingFormat).toBeNull();
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      s.insertPlainText('!');
      s.flushPendingInput();
      expect(textColorState(editor, channel).value).toBe('#17A2B8');
      expect(setTextColor(editor, channel, null).ok).toBe(true);
      s.insertPlainText('?');
      s.flushPendingInput();
      s.setSelection({
        anchor: { paragraphId: id, offset: 6 },
        head: { paragraphId: id, offset: 7 },
      });
      const cleared = textColorState(editor, channel);
      expect(cleared.value).not.toBe('#17A2B8');
      expect(serializeOoxmlPart(s.session.part())).not.toContain(
        channel === 'foreground' ? '<w:color/>' : '<w:shd/>'
      );
    } finally {
      editor.destroy();
    }
  });

test('captured palette set/clear states, fixed highlights, and invalid inputs keep native semantics', () => {
  for (const id of [
    'font.color.set',
    'font.color.clear',
    'font.backgroundColor.set',
    'font.backgroundColor.clear',
    'font.color.inherit',
    'font.background.transparent',
  ])
    expect(reference.cases.find((c) => c.id === id)).toBeDefined();
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(
      '<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Alpha</w:t></w:r></w:p>'
    ),
  });
  const toolbar = createColorToolbar(editor, ['#17A2B8']);
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!;
    s.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 5 },
    });
    expect(textColorState(editor, 'background').value).toBe('#FFFF00');
    expect(setTextColor(editor, 'background', '#17A2B8').ok).toBe(true);
    expect(textColorState(editor, 'background').value).toBe('#17A2B8');
    expect(toolbar.element.querySelector('[aria-label="Current background"]')!.textContent).toBe(
      '#17A2B8'
    );
    const before = canonicalOoxmlFingerprint(s.session.part());
    expect(setTextColor(editor, 'background', 'url(https://example.com)').ok).toBe(false);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
  } finally {
    toolbar.dispose();
    editor.destroy();
  }
});

test('Enter carries imported shading without overriding an inherited highlight; an explicit caret picker still overrides it', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(
      '<w:p><w:r><w:rPr><w:shd w:val="clear" w:fill="17A2B8"/></w:rPr><w:t>Alpha</w:t></w:r></w:p>',
      {
        styles:
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:highlight w:val="yellow"/></w:rPr></w:style></w:styles>',
      }
    ),
  });
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!;
    s.setSelection({
      anchor: { paragraphId: id, offset: 5 },
      head: { paragraphId: id, offset: 5 },
    });
    expect(textColorState(editor, 'background').value).toBe('#FFFF00');
    s.splitParagraph();
    s.insertPlainText('After');
    s.flushPendingInput();
    expect(textColorState(editor, 'background').value).toBe('#FFFF00');
    expect(serializeOoxmlPart(s.session.part())).not.toContain('<w:highlight');
    const second = s.session.paragraphIds()[1]!;
    expect(setTextColor(editor, 'background', '#CC0000').ok).toBe(true);
    s.insertPlainText('!');
    s.flushPendingInput();
    const selection = {
      anchor: { paragraphId: second, offset: 5 },
      head: { paragraphId: second, offset: 6 },
    };
    s.setSelection(selection);
    const span = spansInSelection(s.layout(), selection, s.session.paragraphIds())[0]!;
    expect(span.style.highlight).toBeNull();
    expect(span.style.shading).toBe('CC0000');
  } finally {
    editor.destroy();
  }
});
