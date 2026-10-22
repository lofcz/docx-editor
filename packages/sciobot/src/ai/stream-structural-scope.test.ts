import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { hostParagraph } from '../../test/support/react-block-document';
import { beginStreamInsertion } from './stream-insert';

const fixture = () =>
  docxFixture(
    hostParagraph('Intro') +
      '<w:tbl><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc>' +
      hostParagraph('Cell') +
      '</w:tc></w:tr></w:tbl>' +
      '<w:sdt><w:sdtPr><w:tag w:val="foreign-control"/></w:sdtPr><w:sdtContent>' +
      hostParagraph('Foreign') +
      '</w:sdtContent></w:sdt>' +
      hostParagraph('Tail')
  );
const bodyOf = (editor: DocxEditorInstance) => {
  const body = bodyStoryRoot(editor.surface!.session.part());
  if (!body || body.kind === 'textValue') throw new Error('Missing Word body.');
  return body;
};
const root = (editor: DocxEditorInstance, kind: 'table' | 'contentControl') =>
  bodyOf(editor).children.find((node) => node.kind === kind)!;
const idOf = (editor: DocxEditorInstance, text: string) =>
  editor
    .surface!.session.paragraphIds()
    .find((id) => paragraphTextOf(editor.surface!.session.part(), id) === text)!;
const append = (text: string) => ({
  inserted: [{ type: 'paragraph', children: [{ type: 'text', value: text }] }],
});

for (const scope of ['table', 'contentControl', 'caret'] as const)
  for (const receiver of [0, 1] as const)
    test(`generation after ${scope} preserves peer structure, history and deletion; receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ document: fixture() });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const kind = scope === 'contentControl' ? 'contentControl' : 'table';
        if (scope === 'caret') {
          const paragraphId = idOf(alice, 'Cell');
          alice.surface!.setSelection({
            anchor: { paragraphId, offset: 2 },
            head: { paragraphId, offset: 2 },
          });
        }
        const started = beginStreamInsertion(
          alice,
          { documentId: 'structural', epoch: 'mount' },
          scope === 'caret' ? [] : [root(alice, kind).id]
        );
        if (!started.success) throw new Error(started.error);
        pair.pause();
        expect(started.insertion.append(append('First')).success).toBe(true);
        const paragraphId = idOf(bob, kind === 'table' ? 'Cell' : 'Foreign');
        const old = paragraphTextOf(bob.surface!.session.part(), paragraphId)!;
        const ops: TreeDocOp[] = [
          { op: 'insertText', paragraphId, offset: old.length, text: ' Bob' },
        ];
        const table = root(bob, 'table');
        if (table.kind !== 'table') throw new Error('Missing native table.');
        const row = table.children.find((node) => node.kind === 'tableRow')!;
        ops.push({ op: 'insertTableRow', tableId: table.id, rowId: row.id, where: 'below' });
        expect(bob.surface!.applyAutomationOps(() => ops).rejected).toBe(false);
        const peer = readOoxmlPackage(new Uint8Array(await bob.save()));
        if (!peer.ok) throw new Error(peer.reason);
        const peerOnly = semanticDigest(peer.package.parts.values());
        pair.reconnect(receiver);
        await pair.settle();
        expect(started.insertion.state).toBe('active');
        expect(started.insertion.append(append('Second')).success).toBe(true);
        await pair.settle();
        started.insertion.finish();
        for (const editor of pair.editors) {
          const body = bodyOf(editor);
          const index = body.children.findIndex((node) => node.id === root(editor, kind).id);
          expect(
            body.children
              .slice(index + 1, index + 3)
              .map((node) => paragraphTextOf(editor.surface!.session.part(), node.id))
          ).toEqual(['First', 'Second']);
        }
        const landed = await pair.savedDigests();
        expect(landed[0]).toEqual(landed[1]);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(await pair.savedDigests()).toEqual([peerOnly, peerOnly]);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(await pair.savedDigests()).toEqual(landed);

        const next = beginStreamInsertion(alice, { documentId: 'structural', epoch: 'mount' }, [
          root(alice, kind).id,
        ]);
        if (!next.success) throw new Error(next.error);
        pair.pause();
        const node = root(bob, kind);
        const removal: TreeDocOp =
          kind === 'table'
            ? { op: 'deleteBlock', blockId: node.id }
            : { op: 'removeContentControl', controlId: node.id, keepContent: false };
        expect(bob.surface!.applyAutomationOps(() => [removal]).rejected).toBe(false);
        const edited = idOf(alice, old + ' Bob');
        expect(
          alice.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: edited, offset: old.length, text: ' Alice' },
          ]).rejected
        ).toBe(false);
        pair.reconnect(receiver);
        await pair.settle();
        const beforeLate = await pair.savedDigests();
        expect(beforeLate[0]).toEqual(beforeLate[1]);
        expect(next.insertion.state).toBe('ended');
        expect(next.insertion.append(append('Late')).success).toBe(false);
        expect(await pair.savedDigests()).toEqual(beforeLate);
      } finally {
        pair.destroy();
      }
    });
