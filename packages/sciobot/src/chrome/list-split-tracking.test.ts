import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  replaceNode,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { formattingFingerprint } from '../../test/support/formatting-fingerprint';
import { listDocument, checklistParagraph, listParagraph } from '../../test/support/list-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';
import { readEditorRevision } from '../projection/revision-cache';
import { handleManualListKey } from './list-input';
import { manualListState } from './lists';
const context = { documentId: 'manual-lists', epoch: 'mount' };
for (const selection of [
  [5, 5],
  [3, 3],
  [2, 4],
  [4, 2],
] as const)
  for (const decision of ['accept', 'reject'] as const)
    test(`tracked checklist Enter ${selection} is one native transaction, review group and history unit: ${decision}`, async () => {
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: listDocument(checklistParagraph(true, 1) + listParagraph('Tail')),
        author: 'User',
        modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      });
      try {
        const s = editor.surface!,
          id = s.session.paragraphIds()[0]!;
        s.setSelection({
          anchor: { paragraphId: id, offset: selection[0] },
          head: { paragraphId: id, offset: selection[1] },
        });
        expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
        const original = s.session.part(),
          before = canonicalOoxmlFingerprint(original),
          originalSelection = s.state().selection;
        const packageVersion = s.session.packageRevision();
        expect(handleManualListKey(editor, 'Enter', context)).toBe(true);
        expect(s.session.packageRevision()).toBe(packageVersion + 1);
        const next = s.state().selection.head.paragraphId;
        expect(next).not.toBe(id);
        expect(manualListState(editor, context, next)?.checkbox).toMatchObject({
          checked: selection[0] !== 5,
          level: 1,
        });
        const after = canonicalOoxmlFingerprint(s.session.part()),
          afterSelection = s.state().selection;
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
        expect(s.state().selection).toEqual(originalSelection);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
        expect(s.state().selection).toEqual(afterSelection);
        const groups = readReviewGroups(s.session.currentPackage());
        if (!groups.success) throw new Error(groups.error);
        expect(groups.groups).toHaveLength(1);
        const decide = decision === 'accept' ? acceptReviewGroup : rejectReviewGroup;
        const result = decide(
          editor,
          groups.groups[0]!.id,
          readEditorRevision(editor, context.documentId, context.epoch)!,
          context
        );
        expect(result.success).toBe(true);
        if (decision === 'reject') {
          // Word removes the first paragraph mark. The second paragraph keeps its durable IDs.
          const source = findNode(original, id)!,
            survivor = findNode(s.session.part(), next)!;
          if (source.kind !== 'paragraph' || survivor.kind !== 'paragraph')
            throw new Error('Missing paragraph');
          const identity = new Set(['paraId', 'textId']);
          const attributes = source.attributes.map((attribute) =>
            attribute.namespaceUri === 'http://schemas.microsoft.com/office/word/2010/wordml' &&
            identity.has(attribute.localName)
              ? survivor.attributes.find(
                  (value) =>
                    value.namespaceUri === attribute.namespaceUri &&
                    value.localName === attribute.localName
                )!
              : attribute
          );
          const expected = replaceNode(original, id, { ...source, attributes });
          if (!expected.ok) throw new Error(JSON.stringify(expected.issues));
          expect(formattingFingerprint(s.session.part())).toEqual(
            formattingFingerprint(expected.part)
          );
          expect(findNode(s.session.part(), id)).toBeNull();
        } else {
          expect(s.session.paragraphIds()).toHaveLength(3);
          expect(paragraphTextOf(s.session.part(), id)).toBe(
            selection[0] === selection[1] ? '☒Task'.slice(0, selection[0]) : '☒T'
          );
          expect(manualListState(editor, context, next)?.checkbox).toMatchObject({
            checked: selection[0] !== 5,
            level: 1,
          });
        }
        const resolved = canonicalOoxmlFingerprint(s.session.part());
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(after);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(resolved);
        const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!saved.ok) throw new Error(saved.reason);
        expect(
          canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
        ).toEqual(resolved);
      } finally {
        editor.destroy();
      }
    });

for (const decision of ['accept', 'reject'] as const)
  test(`splitting pending AI text preserves its separate review ownership: ${decision}`, () => {
    const source = checklistParagraph(true, 1).replace(
      '<w:r><w:t>Task</w:t></w:r>',
      '<w:ins w:id="17" w:author="Bot" w:date="2026-10-22T10:00:00Z"><w:r><w:t>Task</w:t></w:r></w:ins>'
    );
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: listDocument(source + listParagraph('Tail')),
      author: 'User',
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
    });
    try {
      const s = editor.surface!,
        id = s.session.paragraphIds()[0]!;
      const point = { paragraphId: id, offset: 3 };
      s.setSelection({ anchor: point, head: point });
      expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      const original = readReviewGroups(s.session.currentPackage());
      if (!original.success) throw new Error(original.error);
      expect(original.groups).toHaveLength(1);
      const bot = original.groups[0]!;
      expect(handleManualListKey(editor, 'Enter', context)).toBe(true);
      const groups = readReviewGroups(s.session.currentPackage());
      if (!groups.success) throw new Error(groups.error);
      expect(groups.groups).toHaveLength(2);
      expect(
        groups.groups
          .find((group) => group.id === bot.id)
          ?.native.every((member) => member.address.author === 'Bot')
      ).toBe(true);
      const manual = groups.groups.find((group) => group.id !== bot.id)!;
      expect(manual.native.every((member) => member.address.author === 'User')).toBe(true);
      const decide = decision === 'accept' ? acceptReviewGroup : rejectReviewGroup;
      expect(
        decide(
          editor,
          manual.id,
          readEditorRevision(editor, context.documentId, context.epoch)!,
          context
        ).success
      ).toBe(true);
      const remaining = readReviewGroups(s.session.currentPackage());
      if (!remaining.success) throw new Error(remaining.error);
      expect(remaining.groups).toHaveLength(1);
      expect(remaining.groups[0]!.id).toBe(bot.id);
      expect(remaining.groups[0]!.native.every((member) => member.address.author === 'Bot')).toBe(
        true
      );
    } finally {
      editor.destroy();
    }
  });

test('a refused tracked checklist continuation consumes Enter without a partial split or fallback', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: listDocument(checklistParagraph(true, 1) + listParagraph('Tail')),
    author: 'User',
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
  try {
    const s = editor.surface!,
      point = { paragraphId: s.session.paragraphIds()[0]!, offset: 5 };
    s.setSelection({ anchor: point, head: point });
    expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
    s.setAuthor(undefined);
    const before = canonicalOoxmlFingerprint(s.session.part()),
      revision = s.session.packageRevision();
    expect(handleManualListKey(editor, 'Enter', context)).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    expect(s.session.packageRevision()).toBe(revision);
    expect(s.state().lastRejection).toBeTruthy();
    expect(editor.exec({ type: 'undo' }).ok).toBe(false);
  } finally {
    editor.destroy();
  }
});
