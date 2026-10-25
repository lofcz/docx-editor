import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { listParagraph, listNumbering } from '../../test/support/list-document';
import { applyManualListInput } from './list-input';
import { manualListMarker, manualListState } from './lists';
const context = { documentId: 'list-input-restart', epoch: 'mount' };
for (const prior of [false, true])
  test(`typed ordered prefix uses a polite restart: prior=${prior}`, () => {
    const body = (prior ? listParagraph('Prior', '1') : '') + '<w:p><w:r><w:t>1.</w:t></w:r></w:p>';
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(body, { numbering: listNumbering }),
    });
    try {
      const s = editor.surface!,
        ids = s.session.paragraphIds(),
        id = ids.at(-1)!;
      s.setSelection({
        anchor: { paragraphId: id, offset: 2 },
        head: { paragraphId: id, offset: 2 },
      });
      const before = canonicalOoxmlFingerprint(s.session.part());
      expect(applyManualListInput(editor, context)).toBe(true);
      expect(manualListMarker(editor, id)?.text).toBe(prior ? '2.' : '1.');
      const state = manualListState(editor, context, id)!;
      expect(state.numbering).not.toBeNull();
      if (!prior) expect(state.numbering!.numId).toBeDefined();
      expect(canonicalOoxmlFingerprint(s.session.part())).not.toEqual(before);
    } finally {
      editor.destroy();
    }
  });
