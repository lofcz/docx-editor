import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  replaceNode,
  readOoxmlPackage,
  type OoxmlNode,
  type OoxmlElement,
} from '@docx-editor.dev/core/store';
import { docxFixture, W } from '../../test/support/document';
import { listParagraph, listNumbering } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { toggleManualList } from './lists';
const context = { documentId: 'toggle-opaque-indent', epoch: 'mount' };
for (const decision of ['direct', 'accept', 'reject'] as const)
  test(`toggle removal keeps an opaque indentation holder without trailing Word attributes: ${decision}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      author: 'User',
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      document: docxFixture(
        listParagraph('Item', '1').replace(
          '<w:customKeep',
          '<w:ind xmlns:u="urn:opaque" w:left="720" w:hanging="360" u:keep="yes"><u:child/></w:ind><w:customKeep'
        ),
        { numbering: listNumbering }
      ),
    });
    try {
      const s = editor.surface!,
        id = s.session.paragraphIds()[0]!,
        point = { paragraphId: id, offset: 2 };
      s.setSelection({ anchor: point, head: point });
      const original = s.session.part(),
        before = canonicalOoxmlFingerprint(original);
      const removed = applyTreeOp(original, {
        op: 'setListNumbering',
        paragraphId: id,
        numId: '0',
      });
      if (!removed.ok) throw new Error(removed.reason);
      const p = findNode(removed.part, id)!;
      if (p.kind !== 'paragraph') throw new Error('Missing paragraph.');
      const properties = p.children[0]!;
      if (properties.localName !== 'pPr') throw new Error('Missing properties.');
      const expected = replaceNode(removed.part, properties.id, {
        ...properties,
        children: properties.children.map((child) =>
          child.kind !== 'textValue' && child.namespaceUri === W && child.localName === 'ind'
            ? ({
                ...child,
                attributes: child.attributes.filter((attr) => attr.namespaceUri !== W),
              } as OoxmlElement)
            : child
        ),
      } as OoxmlNode);
      if (!expected.ok) throw new Error('Invalid oracle.');
      if (decision !== 'direct')
        expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      const changed = toggleManualList(editor, 'decimal', context);
      if (!changed.success) throw new Error(changed.error);
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
      }
      const fingerprint = decision === 'reject' ? before : canonicalOoxmlFingerprint(expected.part);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(fingerprint);
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      expect(
        canonicalOoxmlFingerprint(reopened.package.parts.get(reopened.package.mainDocumentPart)!)
      ).toEqual(fingerprint);
    } finally {
      editor.destroy();
    }
  });
