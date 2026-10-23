import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import {
  nodeChildren,
  namedChild,
  rowCells,
  tableRows,
} from '../../test/support/native-table-oracle';
import { prepareNativeDocument } from '../native/prepare-document';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { planTargetDocument } from './plan-target';
import { applyDirectPlan } from '../ai/apply-direct';
const context = { documentId: 'concurrent-born-columns', epoch: 'native' };
const body =
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Keep</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Peer</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
const table = (part: OoxmlPart) =>
  nodeChildren(bodyStoryRoot(part)!).find((node) => node.kind === 'table')!;
const grid = (part: OoxmlPart) => nodeChildren(namedChild(table(part), 'tblGrid')!);
const partOf = (editor: DocxEditorInstance) => editor.surface!.session.part();
const word = (editor: DocxEditorInstance) => canonicalOoxmlFingerprint(partOf(editor));
async function insert(editor: DocxEditorInstance, label: string, math: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const source = await readWorkingView(editor.surface!.session, revision);
  if (!source.success) throw Error(source.error);
  const td = (value: string, span = 1) =>
    `<td${span > 1 ? ` colspan="${span}"` : ''}>\n\n${value}\n\n</td>`;
  const planned = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: source.view.projection },
    `<table>\n<tr>\n${td('Merged', 3)}\n</tr>\n<tr>\n${td('Keep')}\n${td(`${label} $${math}$`)}\n${td('Peer')}\n</tr>\n</table>`
  );
  if (!planned.success) throw Error(planned.error);
  const applied = applyDirectPlan(editor, planned.plan, context);
  if (!applied.success) throw Error(applied.error);
}
function headerWord(node: OoxmlNode, span: number): OoxmlNode {
  if (node.kind === 'textValue') return node;
  return {
    ...node,
    attributes:
      node.localName === 'gridSpan'
        ? node.attributes.map((attr) =>
            attr.localName === 'val' ? { ...attr, value: String(span) } : attr
          )
        : node.localName === 'tcW'
          ? node.attributes.map((attr) =>
              attr.localName === 'w' ? { ...attr, value: String(span * 2400) } : attr
            )
          : node.attributes,
    children: node.children.map((child) => headerWord(child, span)),
  } as OoxmlNode;
}
for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`independent column births join merged-header coverage and retain peer-owned cells: ${clientIds}/${receiver}`, async () => {
      const prepared = prepareNativeDocument(docxFixture(body));
      if (!prepared.success) throw Error(prepared.error);
      const fixture = { document: prepared.document, clientIds };
      const pair = await createTwoEditorFixture(fixture);
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const initial = partOf(alice);
        const initialGrid = grid(initial).map((node) => node.id);
        const rows = tableRows(initial, table(initial).id);
        const originals = rowCells(rows[1]!).map((cell) => ({
          id: cell.id,
          word: canonicalOoxmlFingerprint(cell),
        }));
        pair.pause();
        await insert(alice, 'TailA', 'x');
        const isolatedAlice = word(alice);
        const nativeInsertedHeader = rowCells(
          tableRows(partOf(alice), table(partOf(alice)).id)[0]!
        )[0]!;
        await insert(bob, 'TailB', 'y');
        const births = pair.editors.map((editor) => {
          const part = partOf(editor);
          const column = grid(part).find((node) => !initialGrid.includes(node.id))!.id;
          const cell = rowCells(tableRows(part, table(part).id)[1]!)[1]!;
          return { column, cell: cell.id, word: canonicalOoxmlFingerprint(cell) };
        });
        pair.reconnect(receiver);
        await pair.settle();
        const check = (a: string | null, b: string | null) => {
          for (const editor of pair.editors) {
            const part = partOf(editor);
            const columns = grid(part).map((node) => node.id);
            expect([...columns].sort()).toEqual(
              [
                ...initialGrid,
                ...(a === null ? [] : [births[0]!.column]),
                ...(b === null ? [] : [births[1]!.column]),
              ].sort()
            );
            const nativeRows = tableRows(part, table(part).id);
            const header = rowCells(nativeRows[0]!)[0]!;
            expect(canonicalOoxmlFingerprint(header)).toBe(
              canonicalOoxmlFingerprint(headerWord(nativeInsertedHeader, columns.length))
            );
            const expected = new Map([
              [initialGrid[0]!, { id: originals[0]!.id, text: 'Keep' }],
              [initialGrid[1]!, { id: originals[1]!.id, text: 'Peer' }],
              [births[0]!.column, { id: births[0]!.cell, text: a ?? '' }],
              [births[1]!.column, { id: births[1]!.cell, text: b ?? '' }],
            ]);
            expect(
              rowCells(nativeRows[1]!).map((cell) => ({
                id: cell.id,
                text: nodeChildren(cell)
                  .filter((node) => node.kind === 'paragraph')
                  .map((p) => paragraphTextOf(part, p.id))
                  .join('\n'),
              }))
            ).toEqual(columns.map((id) => expected.get(id)!));
            for (const cell of originals)
              expect(canonicalOoxmlFingerprint(findNode(part, cell.id)!)).toBe(cell.word);
          }
          expect(word(alice)).toBe(word(bob));
        };
        check('TailA x', 'TailB y');
        const combined = word(alice);
        for (const editor of pair.editors)
          for (const birth of births)
            expect(canonicalOoxmlFingerprint(findNode(partOf(editor), birth.cell)!)).toBe(
              birth.word
            );
        const paragraph = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(partOf(bob), id)?.startsWith('TailA'))!;
        expect(
          bob.surface!.applyAutomationOps(
            () => [
              {
                op: 'insertText',
                paragraphId: paragraph,
                offset: paragraphTextOf(partOf(bob), paragraph)!.length,
                text: ' peerA',
              },
            ],
            { kind: 'body' },
            [],
            false,
            true
          ).rejected
        ).toBe(false);
        await pair.settle();
        const authored = word(alice);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(' peerA', 'TailB y');
        const foreign = word(alice);
        const beforeReads = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
        for (const room of [false, true]) {
          const reopened = await createTwoEditorFixture({
            ...fixture,
            clientIds: undefined,
            document: new Uint8Array(await alice.save()),
            ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
          });
          try {
            await reopened.settle();
            for (const editor of reopened.editors) expect(word(editor)).toBe(foreign);
          } finally {
            reopened.destroy();
          }
        }
        pair.documents.forEach((doc, at) =>
          expect(Y.encodeStateAsUpdate(doc)).toEqual(beforeReads[at]!)
        );
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(null, 'TailB y');
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(' peerA', 'TailB y');
        expect(word(alice)).toBe(foreign);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(word(alice)).toBe(authored);
        expect(word(bob)).toBe(authored);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check('TailA x', null);
        expect(word(alice)).toBe(isolatedAlice);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(word(alice)).toBe(combined);
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
      } finally {
        pair.destroy();
      }
    }, 60_000);
