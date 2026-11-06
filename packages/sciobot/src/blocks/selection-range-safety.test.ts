import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { bodyStoryRoot, canonicalOoxmlFingerprint, findNode } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { blockClipboardDocument } from '../../test/support/react-block-document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { buildBlockIndex } from './block-index';
import { duplicateBlock, duplicateBlockSelection } from './duplicate';
import { moveBlockSelection } from './selection-move';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const boundaries = [
  ['bookmark', '<w:bookmarkStart w:id="1" w:name="Range"/>', '<w:bookmarkEnd w:id="1"/>'],
  [
    'comment',
    '<w:commentRangeStart w:id="1"/>',
    '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>',
  ],
  [
    'field',
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF Range </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
  ],
] as const;

for (const [kind, start, end] of boundaries)
  test(`${kind} range movement refuses split boundaries and interior extraction, but moves the complete range`, () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(
        p('Before') +
          `<w:p>${start}<w:r><w:t>Start</w:t></w:r></w:p>` +
          p('Inside') +
          `<w:p><w:r><w:t>End</w:t></w:r>${end}</w:p>` +
          p('After')
      ),
    });
    try {
      const surface = editor.surface!;
      const before = canonicalOoxmlFingerprint(surface.session.part());
      const ids = surface.session.paragraphIds();
      const parentId = bodyStoryRoot(surface.session.part())!.id;
      for (const nodeId of [ids[1]!, ids[2]!, ids[3]!]) {
        expect(
          moveBlockSelection(editor, { nodeIds: [nodeId], parentId, beforeId: null })
        ).toMatchObject({ success: false, code: 'unsupported-operation' });
        expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
        expect(surface.session.canUndo()).toBe(false);
      }
      for (const nodeIds of [
        [ids[0]!, ids[2]!],
        [ids[1]!, ids[3]!],
      ]) {
        expect(moveBlockSelection(editor, { nodeIds, parentId, beforeId: null })).toMatchObject({
          success: false,
          code: 'unsupported-operation',
        });
        expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
        expect(surface.session.canUndo()).toBe(false);
      }
      for (const nodeIds of [[ids[1]!], [ids[3]!], ids.slice(1, 4)]) {
        expect(
          duplicateBlockSelection(editor, { nodeIds, allocateBlockId: () => 'unused' })
        ).toMatchObject({ success: false, code: 'unsupported-operation' });
        expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
        expect(surface.session.canUndo()).toBe(false);
      }
      expect(
        moveBlockSelection(editor, { nodeIds: ids.slice(1, 4), parentId, beforeId: null })
      ).toMatchObject({ success: true, changed: true });
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
    } finally {
      editor.destroy();
    }
  });

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`registered copies retain exact identity fallback without optional provider resolution (${clientIds})`, async () => {
    const pair = await createTwoEditorFixture({ document: blockClipboardDocument(), clientIds });
    try {
      const editor = pair.editors[0]!;
      Object.defineProperty(pair.rooms[0]!.session, 'resolveNodeId', {
        value: undefined,
        configurable: true,
      });
      expect(editor.surface!.collaborationSession()?.resolveNodeId).toBeUndefined();
      const block = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(
        'activity-1'
      )!;
      const copied = duplicateBlock(editor, {
        controlId: block.controlId,
        allocateBlockId: (type) => `fallback-${type}`,
      });
      if (!copied.success) throw Error(copied.error);
      expect(findNode(editor.surface!.session.part(), copied.controlId)?.kind).toBe(
        'contentControl'
      );
      await pair.settle();
      expect(
        buildBlockIndex(editor.surface!.session.currentPackage()).byId.has(
          `fallback-${block.payload.type}`
        )
      ).toBe(true);
      const paragraphId = editor.surface!.session.paragraphIds().at(-1)!;
      const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
      const revision = editor.surface!.session.packageRevision();
      expect(
        duplicateBlockSelection(editor, { nodeIds: [paragraphId], allocateBlockId: () => 'unused' })
      ).toMatchObject({ success: false, code: 'unsupported-operation' });
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
      expect(editor.surface!.session.packageRevision()).toBe(revision);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
