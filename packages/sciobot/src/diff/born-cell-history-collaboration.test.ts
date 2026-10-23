import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  storyParagraphs,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { table } from '../../test/support/target-table-fixture';
import {
  nodeChildren,
  namedChild,
  rowCells,
  tableRows,
} from '../../test/support/native-table-oracle';
import { prepareNativeDocument } from '../native/prepare-document';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from '../ai/apply-direct';
import { buildBlockIndex } from '../blocks/block-index';
import { planTargetDocument } from './plan-target';
const context = { documentId: 'born-cell-history', epoch: 'native' };
const nativeTable = (part: OoxmlPart) =>
  nodeChildren(bodyStoryRoot(part)!).find((node) => node.kind === 'table')!;
const topologyOf = (part: OoxmlPart) => {
  const native = nativeTable(part);
  const grid = namedChild(native, 'tblGrid')!;
  return {
    rows: tableRows(part, native.id).map((row) =>
      rowCells(row).map((cell) =>
        storyParagraphs(cell)
          .map((p) => paragraphTextOf(part, p.id))
          .join('\n')
      )
    ),
    widths: nodeChildren(grid)
      .filter((node) => node.localName === 'gridCol')
      .map((node) =>
        node.kind === 'textValue'
          ? null
          : node.attributes.find((attr) => attr.localName === 'w')?.value
      ),
  };
};
for (const kind of ['row', 'column'] as const)
  for (const clientIds of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      for (const late of [false, true])
        test(`born ${kind} keeps foreign text in its original cell through both histories and cold reads: ${clientIds}/${receiver}/late=${late}`, async () => {
          const prepared = prepareNativeDocument(
            docxFixture(
              table([
                ['A', 'B'],
                ['Keep', 'Peer'],
              ])
            )
          );
          if (!prepared.success) throw Error(prepared.error);
          const fixture = { document: prepared.document, clientIds };
          const pair = await createTwoEditorFixture(fixture);
          try {
            const [alice, bob] = pair.editors;
            const initial = canonicalOoxmlFingerprint(alice!.surface!.session.part());
            const initialTopology = topologyOf(alice!.surface!.session.part());
            const originalCells = tableRows(
              alice!.surface!.session.part(),
              nativeTable(alice!.surface!.session.part()).id
            )
              .flatMap(rowCells)
              .map((cell) => ({ id: cell.id, word: canonicalOoxmlFingerprint(cell) }));
            const revision = readEditorRevision(alice!, context.documentId, context.epoch)!;
            const source = await readWorkingView(alice!.surface!.session, revision);
            if (!source.success) throw Error(source.error);
            const target =
              kind === 'row'
                ? '| A | B |\n| - | - |\n| Keep | Peer |\n| $x$ | Tail |'
                : '| A | New | B |\n| - | - | - |\n| Keep | Tail $x$ | Peer |';
            const plan = planTargetDocument(
              {
                revision,
                pkg: alice!.surface!.session.currentPackage(),
                projection: source.view.projection,
              },
              target
            );
            if (!plan.success) throw Error(plan.error);
            const applied = applyDirectPlan(alice!, plan.plan, context);
            if (!applied.success) throw Error(applied.error);
            await pair.settle();
            const born = bob!
              .surface!.session.paragraphIds()
              .find((id) => paragraphTextOf(bob!.surface!.session.part(), id)?.startsWith('Tail'))!;
            expect(born).toBeDefined();
            const createdWidths = topologyOf(alice!.surface!.session.part()).widths;
            if (late) pair.pause();
            const offset = paragraphTextOf(bob!.surface!.session.part(), born)!.length;
            expect(
              bob!.surface!.applyAutomationOps(
                () => [{ op: 'insertText', paragraphId: born, offset, text: ' peerborn' }],
                { kind: 'body' },
                [],
                false,
                true
              ).rejected
            ).toBe(false);
            if (!late) await pair.settle();
            const beforeUndo = canonicalOoxmlFingerprint(bob!.surface!.session.part());
            expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
            if (late) {
              expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(initial);
              pair.reconnect(receiver);
            }
            await pair.settle();
            const checkForeign = async () => {
              const updates = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
              for (const editor of pair.editors) {
                const part = editor.surface!.session.part();
                const expected =
                  kind === 'row'
                    ? [
                        ['A', 'B'],
                        ['Keep', 'Peer'],
                        ['', ' peerborn'],
                      ]
                    : [
                        ['A', '', 'B'],
                        ['Keep', ' peerborn', 'Peer'],
                      ];
                expect(topologyOf(part)).toEqual({ rows: expected, widths: createdWidths });
                expect(paragraphTextOf(part, born)).toBe(' peerborn');
                expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(0);
                for (const original of originalCells)
                  expect(canonicalOoxmlFingerprint(findNode(part, original.id)!)).toBe(
                    original.word
                  );
              }
              pair.documents.forEach((doc, index) =>
                expect(Y.encodeStateAsUpdate(doc)).toEqual(updates[index]!)
              );
              expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(
                canonicalOoxmlFingerprint(bob!.surface!.session.part())
              );
              const digests = await pair.savedDigests();
              expect(digests[0]).toEqual(digests[1]);
            };
            await checkForeign();
            const foreignWord = canonicalOoxmlFingerprint(alice!.surface!.session.part());
            bob!.surface!.setSelection({
              anchor: { paragraphId: born, offset: 9 },
              head: { paragraphId: born, offset: 9 },
            });
            await pair.settle();
            const remote = pair.rooms[0]!.session.remoteSelections().find(
              (selection) => selection.actorId === 'bob'
            );
            expect(remote?.head.nodeId).toBe(born);
            expect(remote?.head.offset).toBe(9);
            const carets = pair.editors.map((editor) =>
              caretAt(editor.surface!.layout(), { paragraphId: born, offset: 9 })
            );
            expect(carets.every(Boolean)).toBe(true);
            expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({
              x: carets[1]!.x,
              y: carets[1]!.y,
            });
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
                  expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
                    foreignWord
                  );
                  expect(topologyOf(editor.surface!.session.part())).toEqual(
                    topologyOf(alice!.surface!.session.part())
                  );
                }
              } finally {
                reopened.destroy();
              }
            }
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors) {
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(initial);
              expect(topologyOf(editor.surface!.session.part())).toEqual(initialTopology);
            }
            expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await checkForeign();
            expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors) {
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(beforeUndo);
              expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(1);
            }
          } finally {
            pair.destroy();
          }
        }, 60_000);
