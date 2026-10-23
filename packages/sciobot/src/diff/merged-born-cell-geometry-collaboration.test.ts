import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import {
  namedChild,
  nodeChildren,
  rowCells,
  tableRows,
} from '../../test/support/native-table-oracle';
import { prepareNativeDocument } from '../native/prepare-document';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';
const context = { documentId: 'merged-birth-history', epoch: 'native' };
const initialXml =
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Keep</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Peer</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
function topology(part: OoxmlPart) {
  const table = nodeChildren(bodyStoryRoot(part)!).find((node) => node.kind === 'table')!;
  return {
    grid: nodeChildren(namedChild(table, 'tblGrid')!).length,
    rows: tableRows(part, table.id).map((row) =>
      rowCells(row).map((cell) => {
        const props = namedChild(cell, 'tcPr');
        const span = props ? namedChild(props, 'gridSpan') : null;
        return {
          text: nodeChildren(cell)
            .filter((node) => node.kind === 'paragraph')
            .map((p) => paragraphTextOf(part, p.id))
            .join('\n'),
          span: Number(
            span && span.kind !== 'textValue'
              ? span.attributes.find((a) => a.localName === 'val')?.value
              : 1
          ),
        };
      })
    ),
  };
}
for (const kind of ['column', 'split'] as const)
  for (const late of [false, true])
    test(`foreign text retains exact ${kind} coverage below a merged header, late=${late}`, async () => {
      const prepared = prepareNativeDocument(docxFixture(initialXml));
      if (!prepared.success) throw Error(prepared.error);
      const fixture = { document: prepared.document, clientIds: [41, 73] as const };
      const pair = await createTwoEditorFixture(fixture);
      try {
        const [alice, bob] = pair.editors;
        const initial = canonicalOoxmlFingerprint(alice!.surface!.session.part());
        const revision = readEditorRevision(alice!, context.documentId, context.epoch)!;
        const source = await readWorkingView(alice!.surface!.session, revision);
        if (!source.success) throw Error(source.error);
        const td = (text: string, span = 1) =>
          `<td${span > 1 ? ` colspan="${span}"` : ''}>\n\n${text}\n\n</td>`;
        const target =
          kind === 'column'
            ? `<table>\n<tr>\n${td('Merged', 3)}\n</tr>\n<tr>\n${td('Keep')}\n${td('Tail $x$')}\n${td('Peer')}\n</tr>\n</table>`
            : `<table>\n<tr>\n${td('Merged')}\n${td('Tail $x$')}\n</tr>\n<tr>\n${td('Keep')}\n${td('Peer')}\n</tr>\n</table>`;
        const planned = planTargetDocument(
          {
            revision,
            pkg: alice!.surface!.session.currentPackage(),
            projection: source.view.projection,
          },
          target
        );
        if (!planned.success) throw Error(planned.error);
        const applied = applyDirectPlan(alice!, planned.plan, context);
        if (!applied.success) throw Error(applied.error);
        await pair.settle();
        const born = bob!
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob!.surface!.session.part(), id)?.startsWith('Tail'))!;
        expect(born).toBeDefined();
        if (late) pair.pause();
        expect(
          bob!.surface!.applyAutomationOps(
            () => [
              {
                op: 'insertText',
                paragraphId: born,
                offset: paragraphTextOf(bob!.surface!.session.part(), born)!.length,
                text: ' peerborn',
              },
            ],
            { kind: 'body' },
            [],
            false,
            true
          ).rejected
        ).toBe(false);
        if (!late) await pair.settle();
        const authored = canonicalOoxmlFingerprint(bob!.surface!.session.part());
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        if (late) {
          expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(initial);
          pair.reconnect(0);
        }
        await pair.settle();
        const expected =
          kind === 'column'
            ? {
                grid: 3,
                rows: [
                  [{ text: 'Merged', span: 3 }],
                  [
                    { text: 'Keep', span: 1 },
                    { text: ' peerborn', span: 1 },
                    { text: 'Peer', span: 1 },
                  ],
                ],
              }
            : {
                grid: 2,
                rows: [
                  [
                    { text: 'Merged', span: 1 },
                    { text: ' peerborn', span: 1 },
                  ],
                  [
                    { text: 'Keep', span: 1 },
                    { text: 'Peer', span: 1 },
                  ],
                ],
              };
        const foreign = canonicalOoxmlFingerprint(alice!.surface!.session.part());
        for (const editor of pair.editors)
          expect(topology(editor.surface!.session.part())).toEqual(expected);
        expect(canonicalOoxmlFingerprint(bob!.surface!.session.part())).toBe(foreign);
        const unchanged = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        pair.documents.forEach((doc, index) =>
          expect(Y.encodeStateAsUpdate(doc)).toEqual(unchanged[index]!)
        );
        for (const room of [false, true]) {
          const reopened = await createTwoEditorFixture({
            ...fixture,
            clientIds: undefined,
            document: new Uint8Array(await alice!.save()),
            ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
          });
          try {
            await reopened.settle();
            for (const editor of reopened.editors) {
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(foreign);
              expect(topology(editor.surface!.session.part())).toEqual(expected);
            }
          } finally {
            reopened.destroy();
          }
        }
        if (kind === 'column') {
          const part = bob!.surface!.session.part();
          const table = nodeChildren(bodyStoryRoot(part)!).find((node) => node.kind === 'table')!;
          const head = rowCells(tableRows(part, table.id)[0]!)[0]!;
          expect(
            bob!.surface!.applyAutomationOps(
              () => [
                { op: 'splitTableCell', tableId: table.id, cellId: head.id, rows: 1, columns: 2 },
              ],
              { kind: 'body' },
              [],
              false,
              true
            ).rejected
          ).toBe(false);
          await pair.settle();
          for (const editor of pair.editors)
            expect(topology(editor.surface!.session.part())).toEqual({
              grid: 4,
              rows: [
                [
                  { text: 'Merged', span: 2 },
                  { text: '', span: 2 },
                ],
                [
                  { text: 'Keep', span: 1 },
                  { text: ' peerborn', span: 2 },
                  { text: 'Peer', span: 1 },
                ],
              ],
            });
          const splitWord = canonicalOoxmlFingerprint(bob!.surface!.session.part());
          for (const room of [false, true]) {
            const reopened = await createTwoEditorFixture({
              ...fixture,
              clientIds: undefined,
              document: new Uint8Array(await bob!.save()),
              ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[1]!) } : {}),
            });
            try {
              await reopened.settle();
              for (const editor of reopened.editors)
                expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(splitWord);
            } finally {
              reopened.destroy();
            }
          }
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(foreign);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(splitWord);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(foreign);
        }
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(initial);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(foreign);
          expect(topology(editor.surface!.session.part())).toEqual(expected);
        }
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(authored);
      } finally {
        pair.destroy();
      }
    }, 60_000);
