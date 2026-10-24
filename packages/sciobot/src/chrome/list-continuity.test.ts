import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { listDocument, listParagraph } from '../../test/support/list-document';
import { setManualList, manualListState, manualListMarker, LIST_STYLES } from './lists';
const context = { documentId: 'manual-lists', epoch: 'mount' };
for (const style of LIST_STYLES)
  for (const reverse of [false, true])
    test(`${style} joins a compatible ${reverse ? 'following' : 'preceding'} counter without changing its definition`, () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: listDocument(
          listParagraph('One') +
            listParagraph('Two') +
            listParagraph('Gap') +
            listParagraph('Other')
        ),
      });
      try {
        const s = editor.surface!,
          ids = s.session.paragraphIds();
        const select = (index: number) => {
          const point = { paragraphId: ids[index]!, offset: 1 };
          s.setSelection({ anchor: point, head: point });
        };
        select(reverse ? 1 : 0);
        expect(setManualList(editor, style, context).success).toBe(true);
        const original = manualListState(editor, context)!.numbering!,
          numbering = canonicalOoxmlFingerprint(
            s.session.currentPackage().parts.get('/word/numbering.xml')!
          );
        select(reverse ? 0 : 1);
        const before = canonicalOoxmlFingerprint(s.session.part());
        expect(setManualList(editor, style, context).success).toBe(true);
        expect(manualListState(editor, context)!.numbering).toEqual(original);
        expect(
          canonicalOoxmlFingerprint(s.session.currentPackage().parts.get('/word/numbering.xml')!)
        ).toEqual(numbering);
        if (style === 'decimal') expect(manualListMarker(editor, ids[1]!)?.text).toBe('2.');
        const after = canonicalOoxmlFingerprint(s.session.part());
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
        // An intervening plain paragraph starts an independent sequence.
        select(3);
        expect(setManualList(editor, style, context).success).toBe(true);
        expect(manualListState(editor, context)!.numbering?.numId).not.toBe(original.numId);
      } finally {
        editor.destroy();
      }
    });
