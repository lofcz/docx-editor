import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  findNode,
  parentNodeOf,
  paragraphTextOf,
  readOoxmlPackage,
  serializeOoxmlPart,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { hostControl, hostParagraph } from '../../test/support/react-block-document';
import { columnTableXml } from './column-geometry';
import { buildBlockIndex } from './block-index';
import { moveColumnContent } from './move-column-content';
import { withPayloadRecord } from './payload-opc-codec';

const boundaries = [
  ['bookmark', '<w:bookmarkStart w:id="1" w:name="Range"/>', '<w:bookmarkEnd w:id="1"/>'],
  ['comment', '<w:commentRangeStart w:id="1"/>', '<w:commentRangeEnd w:id="1"/>'],
  [
    'field',
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF Range </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
  ],
] as const;

function columns(left: string): Uint8Array {
  const read = readOoxmlPackage(
    docxFixture(
      hostControl(
        'group',
        columnTableXml([4680, 4680], (index) =>
          hostControl(
            index === 0 ? 'left' : 'right',
            index === 0 ? left : hostParagraph('Destination')
          )
        )
      )
    )
  );
  if (!read.ok) throw Error(read.reason);
  let pkg = read.package;
  for (const [id, type] of [
    ['group', 'column_group'],
    ['left', 'column'],
    ['right', 'column'],
  ] as const) {
    const written = withPayloadRecord(
      pkg,
      { id, type, version: 1, attributes: type === 'column' ? { width: '50%' } : {} },
      id
    );
    if (!written.ok) throw Error(written.reason);
    pkg = written.pkg;
  }
  return writeOoxmlPackage(pkg);
}

const texts = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));
const idOf = (editor: DocxEditorInstance, text: string) =>
  editor
    .surface!.session.paragraphIds()
    .find((id) => paragraphTextOf(editor.surface!.session.part(), id) === text)!;
const move = (editor: DocxEditorInstance, nodeId: string) =>
  moveColumnContent(editor, {
    groupId: 'group',
    nodeId,
    destinationColumnId: 'right',
    beforeId: null,
    expectedRevision: editor.surface!.session.packageRevision(),
  });
function expectOwner(editor: DocxEditorInstance, nodeId: string, column: string) {
  const surface = editor.surface!;
  const block = buildBlockIndex(surface.session.currentPackage()).byId.get(column)!;
  const content = contentControlContentNodeOf(findNode(surface.session.part(), block.controlId)!)!;
  expect(parentNodeOf(surface.session.part(), nodeId)?.id).toBe(content.id);
}

for (const [kind, start, end] of boundaries)
  test(`column movement preserves ${kind} range membership before any mutation`, () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: columns(
        `<w:p>${start}<w:r><w:t>Start</w:t></w:r></w:p>` +
          hostParagraph('Inside') +
          `<w:p><w:r><w:t>End</w:t></w:r>${end}</w:p>`
      ),
    });
    try {
      const surface = editor.surface!;
      const before = canonicalOoxmlFingerprint(surface.session.part());
      const revision = surface.session.packageRevision();
      for (const text of ['Start', 'Inside', 'End']) {
        const nodeId = surface.session
          .paragraphIds()
          .find((id) => paragraphTextOf(surface.session.part(), id) === text)!;
        expect(
          moveColumnContent(editor, {
            groupId: 'group',
            nodeId,
            destinationColumnId: 'right',
            beforeId: null,
            expectedRevision: revision,
          })
        ).toMatchObject({ success: false, code: 'unsupported-operation' });
        expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
        expect(surface.session.packageRevision()).toBe(revision);
        expect(surface.session.canUndo()).toBe(false);
      }
    } finally {
      editor.destroy();
    }
  });

for (const [kind, start, end] of boundaries)
  test(`a complete ${kind} range moves with its paragraph and exact history`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: columns(
        `<w:p>${start}<w:r><w:t>Whole</w:t></w:r>${end}</w:p>` + hostParagraph('Tail')
      ),
    });
    try {
      const surface = editor.surface!;
      const before = canonicalOoxmlFingerprint(surface.session.part());
      // Closed REF fields own one canonical offset unit. Their saved result remains in XML.
      const value = kind === 'field' ? '\uFFFC' : 'Whole';
      const nodeId = idOf(editor, value);
      expect(move(editor, nodeId).success).toBe(true);
      expectOwner(editor, nodeId, 'right');
      const moved = canonicalOoxmlFingerprint(surface.session.part());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
      expectOwner(editor, nodeId, 'left');
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(moved);
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        document: new Uint8Array(await editor.save()),
      });
      try {
        expect(texts(reopened)).toEqual(['Tail', 'Destination', value]);
        expectOwner(reopened, idOf(reopened, value), 'right');
        expect(serializeOoxmlPart(reopened.surface!.session.part())).toContain('<w:t>Whole</w:t>');
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  });

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`column range movement retains concurrent text and deletion history (${clientIds})`, async () => {
    const [, start, end] = boundaries[0];
    const pair = await createTwoEditorFixture({
      document: columns(
        `<w:p>${start}<w:r><w:t>Whole</w:t></w:r>${end}</w:p>` + hostParagraph('Tail')
      ),
      clientIds,
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const nodeId = idOf(alice, 'Whole');
      pair.pause();
      expect(move(alice, nodeId).success).toBe(true);
      bob.surface!.setSelection({
        anchor: { paragraphId: nodeId, offset: 5 },
        head: { paragraphId: nodeId, offset: 5 },
      });
      bob.surface!.type(' peer');
      bob.surface!.flushPendingInput();
      pair.reconnect(clientIds[0] === 41 ? 0 : 1);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(texts(editor)).toEqual(['Tail', 'Destination', 'Whole peer']);
        expectOwner(editor, nodeId, 'right');
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(texts(editor)).toEqual(['Whole peer', 'Tail', 'Destination']);
        expectOwner(editor, nodeId, 'left');
      }
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      const reopened = createDocxEditor({
        container: document.createElement('div'),
        document: new Uint8Array(await alice.save()),
      });
      try {
        expect(texts(reopened)).toEqual(['Tail', 'Destination', 'Whole peer']);
        expectOwner(reopened, idOf(reopened, 'Whole peer'), 'right');
      } finally {
        reopened.destroy();
      }
      const deleted = bob.surface!.applyAutomationOps(
        () => [{ op: 'deleteBlock', blockId: nodeId }],
        { kind: 'body' }
      );
      expect(deleted.rejected).toBe(false);
      await pair.settle();
      const before = canonicalOoxmlFingerprint(alice.surface!.session.part());
      const revision = alice.surface!.session.packageRevision();
      expect(move(alice, nodeId).success).toBe(false);
      expect(canonicalOoxmlFingerprint(alice.surface!.session.part())).toBe(before);
      expect(alice.surface!.session.packageRevision()).toBe(revision);
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(texts(editor)).toEqual(['Tail', 'Destination', 'Whole peer']);
        expectOwner(editor, nodeId, 'right');
      }
      const after = await pair.savedDigests();
      expect(after[0]).toEqual(after[1]);
    } finally {
      pair.destroy();
    }
  });
