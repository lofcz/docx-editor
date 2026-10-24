import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  directParagraphProperties,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { adjustManualListLevel, manualListState, setManualList } from './lists';
const context = { documentId: 'manual-lists', epoch: 'mount' };
for (const style of ['decimal', 'todo'] as const)
  test(`mixed ${style} and plain indentation uses one exact native history unit`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(
        '<w:p><w:r><w:t>List</w:t></w:r></w:p><w:p><w:pPr><w:ind w:left="120" w:firstLine="80"/></w:pPr><w:r><w:t>Plain</w:t></w:r></w:p>'
      ),
    });
    try {
      const s = editor.surface!,
        [list, plain] = s.session.paragraphIds();
      const point = { paragraphId: list!, offset: 0 };
      s.setSelection({ anchor: point, head: point });
      expect(setManualList(editor, style, context).success).toBe(true);
      s.setSelection({
        anchor: { paragraphId: plain!, offset: 5 },
        head: { paragraphId: list!, offset: style === 'todo' ? 1 : 0 },
      });
      const before = canonicalOoxmlFingerprint(s.session.part()),
        selection = s.state().selection;
      expect(adjustManualListLevel(editor, 'increase', context).success).toBe(true);
      const state = manualListState(editor, context, list!);
      expect(state?.checkbox?.level ?? state?.numbering?.level).toBe(1);
      expect(
        directParagraphProperties(s.session.part(), plain!).find((p) => p.localName === 'ind')
          ?.attributes
      ).toMatchObject({ left: '480', firstLine: '80' });
      const after = canonicalOoxmlFingerprint(s.session.part());
      expect(s.state().selection).toEqual(selection);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      expect(s.state().selection).toEqual(selection);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
      const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!saved.ok) throw new Error(saved.reason);
      expect(
        canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
      ).toEqual(after);
      expect(adjustManualListLevel(editor, 'decrease', context).success).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    } finally {
      editor.destroy();
    }
  });

for (const style of ['decimal', 'todo'] as const)
  test(`mixed ${style} indentation keeps one history unit and restores each reviewed native property`, () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(
        '<w:p><w:r><w:t>List</w:t></w:r></w:p><w:p><w:pPr><w:customKeep w:val="yes"/></w:pPr><w:r><w:t>Plain</w:t></w:r></w:p>'
      ),
      author: 'User',
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
    });
    try {
      const s = editor.surface!,
        [list, plain] = s.session.paragraphIds(),
        point = { paragraphId: list!, offset: 0 };
      s.setSelection({ anchor: point, head: point });
      expect(setManualList(editor, style, context).success).toBe(true);
      s.setSelection({
        anchor: { paragraphId: list!, offset: style === 'todo' ? 1 : 0 },
        head: { paragraphId: plain!, offset: 5 },
      });
      expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      const before = canonicalOoxmlFingerprint(s.session.part());
      expect(adjustManualListLevel(editor, 'increase', context).success).toBe(true);
      const after = canonicalOoxmlFingerprint(s.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
      const groups = readReviewGroups(s.session.currentPackage());
      if (!groups.success) throw new Error(groups.error);
      expect(groups.groups).toHaveLength(2);
      for (const group of groups.groups)
        expect(
          rejectReviewGroup(
            editor,
            group.id,
            readEditorRevision(editor, context.documentId, context.epoch)!,
            context
          ).success
        ).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    } finally {
      editor.destroy();
    }
  });
