import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { listParagraph, listNumbering } from '../../test/support/list-document';
import { setManualList, manualListMarker, manualListState } from './lists';
const context = { documentId: 'removed-list-sequence', epoch: 'mount' };
for (const custom of [false, true])
  test(`removing the middle list item restarts its tail without changing imported counter formatting: custom=${custom}`, async () => {
    const numbering = custom
      ? listNumbering
          .replace('<w:lvlText w:val="%1."/>', '<w:lvlText w:val="§ %1)"/>')
          .replace(
            '</w:pPr>',
            '</w:pPr><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/></w:rPr>'
          )
      : listNumbering;
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(
        listParagraph('One', '1') +
          listParagraph('Middle', '1') +
          listParagraph('Tail', '1') +
          listParagraph('Last', '1'),
        { numbering }
      ),
    });
    try {
      const s = editor.surface!,
        ids = s.session.paragraphIds(),
        point = { paragraphId: ids[1]!, offset: 2 };
      s.setSelection({ anchor: point, head: point });
      const original = s.session.part(),
        before = canonicalOoxmlFingerprint(original),
        definitions = s.session
          .currentPackage()
          .parts.get('/word/numbering.xml')!
          .root.children.map((node) => canonicalOoxmlFingerprint(node));
      const result = setManualList(editor, 'none', context);
      if (!result.success) throw new Error(result.error);
      expect(ids.map((id) => manualListMarker(editor, id)?.text ?? null)).toEqual(
        custom ? ['§ 1)', null, '§ 1)', '§ 2)'] : ['1.', null, '1.', '2.']
      );
      const tail = manualListState(editor, context, ids[2]!)!.numbering!,
        last = manualListState(editor, context, ids[3]!)!.numbering!;
      expect(tail.numId).not.toBe('1');
      expect(last.numId).toBe(tail.numId);
      const after = canonicalOoxmlFingerprint(s.session.part());
      const next = s.session.currentPackage().parts.get('/word/numbering.xml')!.root.children;
      expect(
        next.slice(0, definitions.length).map((node) => canonicalOoxmlFingerprint(node))
      ).toEqual(definitions);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      expect(
        canonicalOoxmlFingerprint(reopened.package.parts.get(reopened.package.mainDocumentPart)!)
      ).toEqual(after);
    } finally {
      editor.destroy();
    }
  });
