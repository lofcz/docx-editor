import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { formattingFingerprint } from '../../test/support/formatting-fingerprint';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { insertManualList } from './insert-list';
import { manualListMarker, manualListState, setManualList, setManualListChecked } from './lists';

const context = { documentId: 'empty-list-insert', epoch: 'mount' };
const insertionType = (source: string) =>
  source.includes('todo') ? 'todo' : source === 'decimal' ? 'decimal' : 'disc';
for (const source of [
  'disc',
  'circle',
  'square',
  'decimal',
  'lower-alpha',
  'upper-alpha',
  'lower-roman',
  'upper-roman',
  'todo',
  'checked-todo',
] as const)
  for (const target of ['disc', 'decimal', 'todo'] as const)
    for (const content of ['', 'Task'])
      for (const decision of ['direct', 'accept', 'reject'] as const)
        if (!content || insertionType(source) === target)
          test(`${content ? 'fully selected' : 'empty'} ${source} inserts ${target} with captured replacement rules and native history: ${decision}`, async () => {
            const editor = createDocxEditor({
              container: document.createElement('div'),
              document: docxFixture('<w:p/>'),
              author: 'User',
              modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
            });
            try {
              const style = source === 'checked-todo' ? 'todo' : source;
              const sameKind = insertionType(source) === target;
              expect(setManualList(editor, style, context).success).toBe(true);
              if (source === 'checked-todo')
                expect(setManualListChecked(editor, true, context).success).toBe(true);
              const s = editor.surface!;
              if (content) {
                s.insertPlainText(content);
                s.flushPendingInput();
                const paragraphId = s.session.paragraphIds()[0]!;
                s.setSelection({
                  anchor: { paragraphId, offset: style === 'todo' ? 1 : 0 },
                  head: { paragraphId, offset: (style === 'todo' ? 1 : 0) + content.length },
                });
              }
              const before = canonicalOoxmlFingerprint(s.session.part());
              const beforeWord = formattingFingerprint(s.session.part());
              const originalId = s.session.paragraphIds()[0]!;
              const originalList = manualListState(editor, context, originalId)!;
              const originalMarker = manualListMarker(editor, originalId);
              const originalSelection = s.state().selection;
              if (decision !== 'direct')
                expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
              const result = insertManualList(editor, target, context);
              if (!result.success) throw new Error(result.error);
              const after = canonicalOoxmlFingerprint(s.session.part());
              const selected = s.state().selection;
              expect(editor.exec({ type: 'undo' }).ok).toBe(true);
              expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
              expect(s.state().selection).toEqual(originalSelection);
              expect(editor.exec({ type: 'redo' }).ok).toBe(true);
              expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
              expect(s.state().selection).toEqual(selected);
              if (decision !== 'direct') {
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
              }
              const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
              if (!saved.ok) throw new Error(saved.reason);
              expect(
                canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
              ).toEqual(canonicalOoxmlFingerprint(s.session.part()));
              if (decision === 'reject') {
                expect(formattingFingerprint(s.session.part())).toEqual(beforeWord);
                return;
              }
              const ids = s.session.paragraphIds();
              expect(ids).toHaveLength(sameKind ? 2 : 1);
              expect(ids[0]).toBe(originalId);
              if (sameKind) {
                expect(manualListMarker(editor, originalId)).toEqual(originalMarker);
                expect(manualListState(editor, context, originalId)).toEqual(originalList);
                if (target !== 'todo' && style === target)
                  expect(manualListState(editor, context, ids[1])!.numbering?.numId).toBe(
                    originalList.numbering!.numId
                  );
              }
              expect(paragraphTextOf(s.session.part(), originalId)).toBe(
                sameKind && style === 'todo' && source === 'checked-todo'
                  ? '('
                  : target === 'todo'
                    ? '☐'
                    : ''
              );
              if (source === 'checked-todo' && sameKind)
                expect(serializeOoxmlPart(s.session.part())).toMatch(/<w:sym\b[^>]*w:char="2612"/);
              const inserted = ids.at(-1)!;
              expect(paragraphTextOf(s.session.part(), inserted)).toBe(
                target === 'todo' ? '☐' : ''
              );
              expect(
                target === 'todo'
                  ? manualListState(editor, context, inserted)!.checkbox?.checked
                  : manualListMarker(editor, inserted)?.text
              ).toBe(
                target === 'todo' ? false : target === 'disc' ? '•' : style === target ? '2.' : '1.'
              );
            } finally {
              editor.destroy();
            }
          });
