import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  directParagraphProperties,
  paragraphTextOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import {
  readOoxmlPackage,
  writeOoxmlPackage,
  stylesPartOf,
  findNode,
} from '@docx-editor.dev/core/store';
import { isHorizontalRuleParagraph } from '../native/horizontal-rule';
import { withBlockTypeStyles } from './block-styles';
import { insertBlock } from './insert-block';
import reference from '../__tests__/fixtures/native-commands.json';
import { docxFixture, W } from '../../test/support/document';
import { listNumbering } from '../../test/support/list-document';
import { setBlockType, type TextBlockType } from './block-type';

const styles = `<w:styles xmlns:w="${W}">${['Normal', 'Heading1', 'Heading2', 'Heading3', 'Quote'].map((id) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/>${id.startsWith('Heading') ? `<w:pPr><w:outlineLvl w:val="${Number(id.at(-1)) - 1}"/></w:pPr>` : ''}</w:style>`).join('')}</w:styles>`;
const context = { documentId: 'blocks', epoch: 'mount' };
for (const [type, expectedStyle] of [
  ['p', 'Normal'],
  ['h1', 'Heading1'],
  ['h2', 'Heading2'],
  ['h3', 'Heading3'],
  ['blockquote', 'Quote'],
] as const) {
  for (const listed of [false, true]) {
    test(`${type} selected block preserves text/caret and imported metadata, removes lists, and undoes once (${listed ? 'list' : 'plain'})`, () => {
      const fixture = reference.cases.find(
        (entry) => entry.id === `block.${type}${listed ? '.from-list' : ''}`
      );
      expect(fixture).toBeDefined();
      expect(fixture!.before.selection?.anchor.offset).toBe(0);
      expect(fixture!.before.selection?.focus.offset).toBe(5);
      if (type === 'blockquote') expect(fixture!.after.selection).toBeNull();
      else expect(fixture!.after.selection).toEqual(fixture!.before.selection);
      const body = `<w:p><w:pPr><w:spacing w:after="123"/><w:customKeep w:val="yes"/>${listed ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>' : ''}</w:pPr><w:r><w:t>Alpha beta</w:t></w:r></w:p>`;
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: docxFixture(body, { styles, numbering: listNumbering }),
      });
      try {
        const surface = editor.surface!,
          session = surface.session;
        const id = session.paragraphIds()[0]!;
        const selection = {
          anchor: { paragraphId: id, offset: 0 },
          head: { paragraphId: id, offset: 5 },
        };
        surface.setSelection(selection);
        const before = directParagraphProperties(session.part(), id);
        expect(setBlockType(editor, type as TextBlockType, context).success).toBe(true);
        const after = directParagraphProperties(session.part(), id);
        expect(paragraphTextOf(session.part(), id)).toBe('Alpha beta');
        expect(after.find((p) => p.localName === 'pStyle')?.attributes?.val).toBe(expectedStyle);
        expect(after.find((p) => p.localName === 'spacing')?.attributes?.after).toBe('123');
        expect(serializeOoxmlPart(session.part())).toContain('<w:customKeep w:val="yes"/>');
        expect(surface.state().selection).toEqual(selection);
        expect(
          surface
            .layout()
            .pages.flatMap((p) => p.fragments)
            .filter((p) => p.kind === 'paragraph')
            .every((p) => !p.marker)
        ).toBe(true);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(directParagraphProperties(session.part(), id)).toEqual(before);
        expect(paragraphTextOf(session.part(), id)).toBe('Alpha beta');
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(
          directParagraphProperties(session.part(), id).find((p) => p.localName === 'pStyle')
            ?.attributes?.val
        ).toBe(expectedStyle);
      } finally {
        editor.destroy();
      }
    });
  }
}

test('preparation retains imported styles and body identity while adding absent native block definitions', () => {
  const loaded = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Keep</w:t></w:r></w:p>', { styles }));
  if (!loaded.ok) throw new Error(loaded.reason);
  const original = loaded.package;
  const prepared = withBlockTypeStyles(original)!;
  expect(prepared).not.toBeNull();
  expect(prepared.parts.get(original.mainDocumentPart)).toBe(
    original.parts.get(original.mainDocumentPart)
  );
  expect(stylesPartOf(prepared)!.root.children.slice(0, 5)).toEqual([
    ...stylesPartOf(original)!.root.children,
  ]);
  expect(withBlockTypeStyles(prepared)).toBe(prepared);
});

for (const type of ['p', 'h1', 'h2', 'h3', 'blockquote', 'hr'] as const) {
  test(`insert ${type} keeps selected source text, selects the new paragraph, and undoes once`, () => {
    const loaded = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p>'));
    if (!loaded.ok) throw new Error(loaded.reason);
    const pkg = withBlockTypeStyles(loaded.package)!;
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(pkg),
    });
    try {
      const surface = editor.surface!,
        session = surface.session;
      const original = session.paragraphIds()[0]!;
      surface.setSelection({
        anchor: { paragraphId: original, offset: 0 },
        head: { paragraphId: original, offset: 5 },
      });
      const result = insertBlock(editor, type);
      if (!result.success) throw new Error(result.error);
      expect(result.success).toBe(true);
      expect(session.paragraphIds()).toEqual([original, result.paragraphId]);
      expect(paragraphTextOf(session.part(), original)).toBe('Alpha beta');
      expect(paragraphTextOf(session.part(), result.paragraphId)).toBe('');
      expect(surface.state().selection).toEqual({
        anchor: { paragraphId: result.paragraphId, offset: 0 },
        head: { paragraphId: result.paragraphId, offset: 0 },
      });
      if (type === 'hr')
        expect(isHorizontalRuleParagraph(findNode(session.part(), result.paragraphId)!)).toBe(true);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(session.paragraphIds()).toEqual([original]);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(session.paragraphIds()).toHaveLength(2);
      expect(paragraphTextOf(session.part(), original)).toBe('Alpha beta');
    } finally {
      editor.destroy();
    }
  });
}
