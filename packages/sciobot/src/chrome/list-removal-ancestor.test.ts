import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  readOoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture, W } from '../../test/support/document';
import { listParagraph, listNumbering } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { setManualList, manualListMarker, manualListState } from './lists';
const context = { documentId: 'removed-nested-counter', epoch: 'mount' };
const numbering = listNumbering.replace('<w:lvlText w:val="%2."/>', '<w:lvlText w:val="%1.%2."/>');
for (const decision of ['direct', 'accept', 'reject'] as const)
  test(`restarting a nested tail preserves the imported parent counter value: ${decision}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      author: 'User',
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      document: docxFixture(
        listParagraph('Root one', '1') +
          listParagraph('Root two', '1') +
          listParagraph('Removed child', '1', 1) +
          listParagraph('Tail child', '1', 1),
        { numbering }
      ),
    });
    try {
      const s = editor.surface!,
        ids = s.session.paragraphIds();
      expect(ids.map((id) => manualListMarker(editor, id)?.text)).toEqual([
        '1.',
        '2.',
        '2.1.',
        '2.2.',
      ]);
      const point = { paragraphId: ids[2]!, offset: 1 };
      s.setSelection({ anchor: point, head: point });
      const original = s.session.part(),
        before = canonicalOoxmlFingerprint(original);
      if (decision !== 'direct')
        expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      const result = setManualList(editor, 'none', context);
      if (!result.success) throw new Error(result.error);
      expect(ids.map((id) => manualListMarker(editor, id)?.text ?? null)).toEqual([
        '1.',
        '2.',
        null,
        '2.1.',
      ]);
      const counter = manualListState(editor, context, ids[3]!)!.numbering!.numId;
      let expected = original;
      for (const [paragraphId, numId, level] of [
        [ids[2]!, '0', 0],
        [ids[3]!, counter, 1],
      ] as const) {
        const applied = applyTreeOp(expected, {
          op: 'setListNumbering',
          paragraphId,
          numId,
          level,
        });
        if (!applied.ok) throw new Error(applied.reason);
        expected = applied.part;
      }
      const numberingPart = s.session.currentPackage().parts.get('/word/numbering.xml')!;
      const descriptor = numberingPart.root.children.find(
        (node) =>
          node.kind !== 'textValue' &&
          node.localName === 'num' &&
          node.attributes.some(
            (attr) =>
              attr.namespaceUri === W && attr.localName === 'numId' && attr.value === counter
          )
      );
      const word = readOoxmlPart(
        `<w:numbering xmlns:w="${W}"><w:num w:numId="${counter}"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="2"/></w:lvlOverride><w:lvlOverride w:ilvl="1"><w:startOverride w:val="1"/></w:lvlOverride></w:num></w:numbering>`,
        { name: numberingPart.name, contentType: numberingPart.contentType }
      );
      if (!word.ok) throw new Error(word.reason);
      expect(descriptor).toBeDefined();
      expect(canonicalOoxmlFingerprint(descriptor!)).toEqual(
        canonicalOoxmlFingerprint(word.part.root.children[0]!)
      );
      const pending = canonicalOoxmlFingerprint(s.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(pending);
      if (decision !== 'direct') {
        const groups = readReviewGroups(s.session.currentPackage());
        if (!groups.success) throw new Error(groups.error);
        expect(groups.groups).toHaveLength(1);
        const decided = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
          editor,
          groups.groups[0]!.id,
          readEditorRevision(editor, context.documentId, context.epoch)!,
          context
        );
        if (!decided.success) throw new Error(decided.error);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(pending);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      }
      const final = decision === 'reject' ? before : canonicalOoxmlFingerprint(expected);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(final);
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      expect(
        canonicalOoxmlFingerprint(reopened.package.parts.get(reopened.package.mainDocumentPart)!)
      ).toEqual(final);
    } finally {
      editor.destroy();
    }
  });
