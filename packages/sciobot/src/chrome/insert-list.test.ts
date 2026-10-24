import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { rejectReviewGroup } from '../review/reject-group';
import { acceptReviewGroup } from '../review/accept-group';
import { readEditorRevision } from '../projection/revision-cache';
import { manualListState, manualListMarker } from './lists';
import { insertManualList } from './insert-list';
const context = { documentId: 'insert-lists', epoch: 'mount' };
for (const style of ['disc', 'decimal', 'todo'] as const)
  for (const offsets of [
    [0, 5],
    [3, 3],
    [10, 10],
    [0, 10],
  ] as const)
    test(`${style} insertion ${offsets} replaces selection, keeps Word sides and records one history unit`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: docxFixture('<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p>'),
      });
      try {
        const s = editor.surface!,
          id = s.session.paragraphIds()[0]!;
        s.setSelection({
          anchor: { paragraphId: id, offset: offsets[0] },
          head: { paragraphId: id, offset: offsets[1] },
        });
        const before = canonicalOoxmlFingerprint(s.session.part()),
          selected = s.state().selection;
        const result = insertManualList(editor, style, context);
        if (!result.success) throw new Error(result.error);
        const ids = s.session.paragraphIds(),
          marker = style === 'todo' ? '☐' : '';
        expect(ids.map((p) => paragraphTextOf(s.session.part(), p))).toEqual(
          offsets[0] === 0
            ? offsets[1] === 10
              ? [marker]
              : [marker, ' beta']
            : offsets[0] === 3
              ? ['Alp', marker, 'ha beta']
              : ['Alpha beta', marker]
        );
        expect(s.state().selection.head).toEqual({
          paragraphId: result.paragraphId,
          offset: marker.length,
        });
        expect(
          style === 'todo'
            ? manualListState(editor, context)?.checkbox?.checked
            : manualListMarker(editor, result.paragraphId)?.text
        ).toBe(style === 'todo' ? false : style === 'disc' ? '•' : '1.');
        const after = canonicalOoxmlFingerprint(s.session.part()),
          selection = s.state().selection;
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
        expect(s.state().selection).toEqual(selected);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
        expect(s.state().selection).toEqual(selection);
        const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!saved.ok) throw new Error(saved.reason);
        expect(
          canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
        ).toEqual(after);
      } finally {
        editor.destroy();
      }
    });
for (const style of ['disc', 'decimal', 'todo'] as const)
  for (const offsets of [
    [0, 5],
    [3, 3],
    [10, 10],
    [0, 10],
    [5, 0],
  ] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`${style} tracked insertion ${offsets} groups deleted text, paragraph boundaries and marker: ${decision}`, () => {
        const editor = createDocxEditor({
          container: document.createElement('div'),
          document: docxFixture(
            '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
          ),
          author: 'User',
          modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
        });
        try {
          const s = editor.surface!,
            id = s.session.paragraphIds()[0]!;
          s.setSelection({
            anchor: { paragraphId: id, offset: offsets[0] },
            head: { paragraphId: id, offset: offsets[1] },
          });
          expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
          const before = canonicalOoxmlFingerprint(s.session.part());
          const result = insertManualList(editor, style, context);
          if (!result.success) throw new Error(result.error);
          const after = canonicalOoxmlFingerprint(s.session.part());
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
          const groups = readReviewGroups(s.session.currentPackage());
          if (!groups.success) throw new Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          const decide = decision === 'accept' ? acceptReviewGroup : rejectReviewGroup;
          expect(
            decide(
              editor,
              groups.groups[0]!.id,
              readEditorRevision(editor, context.documentId, context.epoch)!,
              context
            ).success
          ).toBe(true);
          expect(s.session.paragraphIds().map((p) => paragraphTextOf(s.session.part(), p))).toEqual(
            decision === 'accept'
              ? Math.min(...offsets) === 0
                ? Math.max(...offsets) === 10
                  ? [style === 'todo' ? '☐' : '', 'Tail']
                  : [style === 'todo' ? '☐' : '', ' beta', 'Tail']
                : offsets[0] === 3
                  ? ['Alp', style === 'todo' ? '☐' : '', 'ha beta', 'Tail']
                  : ['Alpha beta', style === 'todo' ? '☐' : '', 'Tail']
              : ['Alpha beta', 'Tail']
          );
        } finally {
          editor.destroy();
        }
      });

for (const style of ['decimal', 'todo'] as const)
  for (const tracked of [false, true])
    test(`${style} insertion preserves both unselected sides of a reversed paragraph-spanning selection: tracked=${tracked}`, () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: docxFixture(
          '<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p><w:p><w:r><w:t>Gamma delta</w:t></w:r></w:p>'
        ),
        author: 'User',
        modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      });
      try {
        const s = editor.surface!,
          [first, last] = s.session.paragraphIds();
        s.setSelection({
          anchor: { paragraphId: last!, offset: 5 },
          head: { paragraphId: first!, offset: 2 },
        });
        if (tracked)
          expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
        const before = canonicalOoxmlFingerprint(s.session.part());
        const result = insertManualList(editor, style, context);
        if (!result.success) throw new Error(result.error);
        const after = canonicalOoxmlFingerprint(s.session.part());
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
        if (tracked) {
          const groups = readReviewGroups(s.session.currentPackage());
          if (!groups.success) throw new Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          expect(
            acceptReviewGroup(
              editor,
              groups.groups[0]!.id,
              readEditorRevision(editor, context.documentId, context.epoch)!,
              context
            ).success
          ).toBe(true);
        }
        expect(s.session.paragraphIds().map((id) => paragraphTextOf(s.session.part(), id))).toEqual(
          ['Al', style === 'todo' ? '☐' : '', ' delta']
        );
      } finally {
        editor.destroy();
      }
    });
