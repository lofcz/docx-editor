import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  replaceNode,
  type OoxmlElement,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { docxFixture, W } from '../../test/support/document';
import { listNumbering, listParagraph } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { toggleManualList, manualListMarker, manualListState } from './lists';
const context = { documentId: 'source-list-toggle', epoch: 'mount' };
for (const expanded of [false, true])
  for (const decision of ['direct', 'accept', 'reject'] as const)
    test(`source list style toggle changes only the required original siblings: expanded=${expanded}/${decision}`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        author: 'User',
        modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
        document: docxFixture(
          listParagraph('One', '1') +
            listParagraph('Nested', '2', 1) +
            listParagraph('Two', '1') +
            listParagraph('Boundary', '2') +
            listParagraph('Tail', '1'),
          { numbering: listNumbering }
        ),
      });
      try {
        const s = editor.surface!,
          ids = s.session.paragraphIds(),
          id = ids[2]!;
        s.setSelection({
          anchor: { paragraphId: id, offset: 1 },
          head: { paragraphId: id, offset: expanded ? 2 : 1 },
        });
        const original = s.session.part(),
          before = canonicalOoxmlFingerprint(original),
          selection = s.state().selection;
        if (decision !== 'direct')
          expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
        const result = toggleManualList(editor, 'lower-roman', context);
        if (!result.success) throw new Error(result.error);
        const membership = manualListState(editor, context, id)!.numbering!;
        const changed = expanded ? [2] : [0, 2];
        let expected = original;
        for (const index of changed) {
          const applied = applyTreeOp(expected, {
            op: 'setListNumbering',
            paragraphId: ids[index]!,
            numId: membership.numId,
            level: 0,
          });
          if (!applied.ok) throw new Error(applied.reason);
          expected = applied.part;
        }
        expect(ids.map((id) => manualListMarker(editor, id)?.text ?? null)).toEqual(
          expanded ? ['1.', '•', 'i.', '•', '2.'] : ['i.', '•', 'ii.', '•', '1.']
        );
        expect(s.state().selection).toEqual(selection);
        const pending = canonicalOoxmlFingerprint(s.session.part());
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(pending);
        if (decision !== 'direct') {
          const groups = readReviewGroups(s.session.currentPackage());
          if (!groups.success) throw new Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
            editor,
            groups.groups[0]!.id,
            readEditorRevision(editor, context.documentId, context.epoch)!,
            context
          );
          if (!result.success) throw new Error(result.error);
          expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(
            decision === 'accept' ? canonicalOoxmlFingerprint(expected) : before
          );
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(pending);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        } else expect(pending).toEqual(canonicalOoxmlFingerprint(expected));
        const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!loaded.ok) throw new Error(loaded.reason);
        expect(
          canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
        ).toEqual(decision === 'reject' ? before : canonicalOoxmlFingerprint(expected));
      } finally {
        editor.destroy();
      }
    });
for (const depth of [0, 2])
  for (const decision of ['direct', 'accept', 'reject'] as const)
    test(`source list toggle removal retains reduced plain indentation and opaque properties: depth=${depth}/${decision}`, async () => {
      const body = listParagraph('Item', '1', depth).replace(
        '<w:customKeep',
        '<w:ind xmlns:u="urn:opaque" w:left="2160" w:hanging="360" w:right="240" w:rightChars="2" u:keep="yes"/><w:customKeep'
      );
      const editor = createDocxEditor({
        container: document.createElement('div'),
        author: 'User',
        modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
        document: docxFixture(body, { numbering: listNumbering }),
      });
      try {
        const s = editor.surface!,
          id = s.session.paragraphIds()[0]!,
          point = { paragraphId: id, offset: 2 };
        s.setSelection({ anchor: point, head: point });
        const original = s.session.part(),
          before = canonicalOoxmlFingerprint(original);
        const unlisted = applyTreeOp(original, {
          op: 'setListNumbering',
          paragraphId: id,
          numId: '0',
        });
        if (!unlisted.ok) throw new Error(unlisted.reason);
        const paragraph = findNode(unlisted.part, id)!;
        if (paragraph.kind !== 'paragraph') throw new Error('Missing paragraph.');
        const leading = new Set([
          'left',
          'start',
          'hanging',
          'firstLine',
          'leftChars',
          'startChars',
          'hangingChars',
          'firstLineChars',
        ]);
        const children = paragraph.children.map((node) => {
          if (node.namespaceUri !== W || node.localName !== 'pPr') return node;
          return {
            ...node,
            children: node.children.map((child) => {
              if (
                child.kind === 'textValue' ||
                child.namespaceUri !== W ||
                child.localName !== 'ind'
              )
                return child;
              const attrs = child.attributes.filter(
                (attr) => attr.namespaceUri !== W || !leading.has(attr.localName)
              );
              if (depth)
                attrs.push({
                  kind: 'genericExtension',
                  namespaceUri: W,
                  prefix: 'w',
                  localName: 'left',
                  value: String(depth * 360),
                });
              return { ...child, attributes: attrs } as OoxmlElement;
            }),
          } as OoxmlElement;
        });
        const expected = replaceNode(unlisted.part, id, { ...paragraph, children } as OoxmlNode);
        if (!expected.ok) throw new Error('Invalid expected paragraph.');
        if (decision !== 'direct')
          expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
        const result = toggleManualList(editor, 'decimal', context);
        if (!result.success) throw new Error(result.error);
        expect(manualListState(editor, context, id)!.numbering).toBeNull();
        const pending = canonicalOoxmlFingerprint(s.session.part());
        if (decision !== 'direct') {
          const groups = readReviewGroups(s.session.currentPackage());
          if (!groups.success) throw new Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
            editor,
            groups.groups[0]!.id,
            readEditorRevision(editor, context.documentId, context.epoch)!,
            context
          );
          if (!result.success) throw new Error(result.error);
        }
        const final = decision === 'reject' ? before : canonicalOoxmlFingerprint(expected.part);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(final);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(
          decision === 'direct' ? before : pending
        );
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(final);
        const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!loaded.ok) throw new Error(loaded.reason);
        expect(
          canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
        ).toEqual(final);
      } finally {
        editor.destroy();
      }
    });
