import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { seedReviewCapacity } from '../../test/support/review-capacity';
import { groups } from '../../test/support/review-groups';
import {
  cellReviewDocument,
  cellScopeFor,
  childNamed,
  CELL_CONTEXT,
} from '../../test/support/table-review';
import { REVIEW_FLAG_INTENTS, REVIEW_FLAG_TARGETS } from '../collaboration/review-flag-codec';
import { planTableFormat } from '../diff/table-format';
import { readEditorRevision } from '../projection/revision-cache';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableFormatPlan } from './track-table-format';

const editorOptions = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
const pkg = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
const children = (node: OoxmlNode | undefined) =>
  node && node.kind !== 'textValue' ? node.children : [];
const word = (editor: DocxEditorInstance, text: string) =>
  editor
    .surface!.session.paragraphIds()
    .find((id) => paragraphTextOf(editor.surface!.session.part(), id) === text)!;
function type(editor: DocxEditorInstance, text: string, value: string) {
  const surface = editor.surface!,
    paragraphId = word(editor, text),
    offset = text.length;
  surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
  surface.insertPlainText(value);
  surface.flushPendingInput();
}
function snapshotNodeCount(node: OoxmlNode): number {
  return 1 + children(node).reduce((sum, child) => sum + snapshotNodeCount(child), 0);
}
function widths(editor: DocxEditorInstance) {
  const table = children(childNamed(editor.surface!.session.part().root, 'body')).find(
    (n) => n.kind === 'table'
  )!;
  return children(childNamed(table, 'tblGrid'))
    .filter((n) => n.kind !== 'textValue' && n.localName === 'gridCol')
    .map((n) => n.kind !== 'textValue' && n.attributes.find((a) => a.localName === 'w')?.value);
}
async function fixture(clients: readonly [number, number], action: 'accept' | 'reject') {
  const pair = await createTwoEditorFixture({
    document: cellReviewDocument(),
    clientIds: clients,
    ownedAttributes: true,
    editor: editorOptions,
  });
  try {
    const alice = pair.editors[0]!,
      bob = pair.editors[1]!,
      scope = cellScopeFor(alice),
      planned = planTableFormat(
        {
          part: alice.surface!.session.part(),
          revision: readEditorRevision(alice, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
          tableId: scope.tableId,
        },
        { op: 'setTableProperties', columnWidthsTwips: [3500, 2500] }
      );
    if (!planned.success) throw new Error(planned.error);
    expect(trackTableFormatPlan(alice, planned.plan, CELL_CONTEXT).success).toBe(true);
    await pair.settle();
    type(bob, 'Two', ' peer');
    await pair.settle();
    const part = alice.surface!.session.part(),
      sites = collectRevisionSites(part);
    let footprint = sites.reduce((sum, site) => sum + snapshotNodeCount(site.node), 0);
    if (action === 'reject') {
      const preview = applyTreeOp(
        { ...part, root: { ...part.root } },
        { op: 'rejectAllRevisions', siteNodeIds: sites.map((site) => site.node.id) }
      );
      if (!preview.ok) throw new Error(preview.reason);
      const collectIds = (node: OoxmlNode, ids = new Set<string>()): Set<string> => {
        ids.add(node.id);
        for (const child of children(node)) collectIds(child, ids);
        return ids;
      };
      const after = collectIds(preview.part.root);
      footprint = [...collectIds(part.root)].filter((id) => !after.has(id)).length;
    }
    expect(sites).toHaveLength(6);
    expect(footprint).toBeGreaterThan(1);
    return { pair, alice, bob, id: groups(alice)[0]!.id, footprint };
  } catch (cause) {
    pair.destroy();
    throw cause;
  }
}
function decide(editor: DocxEditorInstance, id: string, action: 'accept' | 'reject') {
  return (action === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
    editor,
    id,
    readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
    CELL_CONTEXT
  );
}

for (const action of ['accept', 'reject'] as const)
  for (const kind of ['targets', 'intents'] as const)
    for (const clients of [
      [41, 73],
      [73, 41],
    ] as const)
      for (const receiver of [0, 1] as const)
        test(`${action} native group ${kind} capacity refusal preserves Word XML, vectors, peer state, and prior redo; ${clients}/${receiver}`, async () => {
          const { pair, alice, bob, id } = await fixture(clients, action);
          try {
            type(alice, 'Tail', ' note');
            await pair.settle();
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            seedReviewCapacity(pair.documents[0]!, kind, 1);
            await pair.settle();
            const before = pair.editors.map((editor) => ({
                bytes: writeOoxmlPackage(pkg(editor)),
                revision: editor.surface!.session.packageRevision(),
              })),
              vectors = pair.documents.map((doc) => Y.encodeStateVector(doc)),
              counts = pair.documents.map((doc) => [
                doc.getMap(REVIEW_FLAG_TARGETS).size,
                doc.getMap(REVIEW_FLAG_INTENTS).size,
              ]),
              history = pair.rooms.map((room) => [room.session.canUndo(), room.session.canRedo()]);
            expect(history[0]![1]).toBe(true);
            pair.pause();
            for (const editor of pair.editors)
              expect(decide(editor, id, action)).toMatchObject({
                success: false,
                error: expect.stringContaining('record limit'),
              });
            pair.reconnect(receiver);
            await pair.settle();
            pair.editors.forEach((editor, index) => {
              expect(writeOoxmlPackage(pkg(editor))).toEqual(before[index]!.bytes);
              expect(editor.surface!.session.packageRevision()).toBe(before[index]!.revision);
              expect(groups(editor).map((group) => group.id)).toEqual([id]);
              expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(6);
              expect(widths(editor)).toEqual(['3500', '2500']);
              expect(word(editor, 'Two peer')).toBeDefined();
              expect(Y.encodeStateVector(pair.documents[index]!)).toEqual(vectors[index]);
              expect([
                pair.documents[index]!.getMap(REVIEW_FLAG_TARGETS).size,
                pair.documents[index]!.getMap(REVIEW_FLAG_INTENTS).size,
              ]).toEqual(counts[index]);
            });
            expect(
              pair.rooms.map((room) => [room.session.canUndo(), room.session.canRedo()])
            ).toEqual(history);
            const remote = pair.rooms[0]!.session.remoteSelections().find(
              (row) => row.actorId === 'bob'
            );
            expect(remote).toBeDefined();
            expect(
              caretAt(alice.surface!.layout(), {
                paragraphId: remote!.head.nodeId,
                offset: remote!.head.offset,
              })
            ).not.toBeNull();
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors) expect(word(editor, 'Tail note')).toBeDefined();
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors) {
              expect(word(editor, 'Two')).toBeDefined();
              expect(groups(editor).map((group) => group.id)).toEqual([id]);
            }
            const digests = await pair.savedDigests();
            expect(digests[0]).toEqual(digests[1]);
            for (const editor of pair.editors) {
              const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
              if (!saved.ok) throw new Error(saved.reason);
              expect(
                collectRevisionSites(saved.package.parts.get(saved.package.mainDocumentPart)!)
              ).toHaveLength(6);
            }
          } finally {
            pair.destroy();
          }
        }, 30000);

for (const action of ['accept', 'reject'] as const)
  for (const kind of ['targets', 'intents'] as const)
    for (const clients of [
      [41, 73],
      [73, 41],
    ] as const)
      for (const receiver of [0, 1] as const)
        test(`${action} native group succeeds at its exact ${kind} admission budget and retains complete history/reopening; ${clients}/${receiver}`, async () => {
          const { pair, alice, bob, id, footprint } = await fixture(clients, action);
          try {
            seedReviewCapacity(pair.documents[0]!, kind, footprint);
            await pair.settle();
            pair.pause();
            expect(decide(alice, id, action).success).toBe(true);
            pair.reconnect(receiver);
            await pair.settle();
            const assertState = (pending: boolean) => {
              for (const editor of pair.editors) {
                expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(
                  pending ? 6 : 0
                );
                expect(groups(editor).map((group) => group.id)).toEqual(pending ? [id] : []);
                expect(widths(editor)).toEqual(
                  action === 'reject' && !pending ? ['3000', '3000'] : ['3500', '2500']
                );
                expect(word(editor, 'Two peer')).toBeDefined();
              }
            };
            assertState(false);
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            assertState(true);
            if (kind === 'intents') {
              const before = pair.documents.map((doc) => Y.encodeStateVector(doc));
              expect(decide(bob, id, action)).toMatchObject({
                success: false,
                error: expect.stringContaining('record limit'),
              });
              pair.documents.forEach((doc, index) =>
                expect(Y.encodeStateVector(doc)).toEqual(before[index])
              );
              expect(pair.rooms[0]!.session.canRedo()).toBe(true);
              assertState(true);
            }
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            assertState(false);
            const digests = await pair.savedDigests();
            expect(digests[0]).toEqual(digests[1]);
            const docx = new Uint8Array(await alice.save()),
              snapshot = Y.encodeStateAsUpdate(pair.documents[receiver]!);
            for (const options of [
              { document: docx },
              { document: new Uint8Array(), roomSnapshot: snapshot },
            ]) {
              const reopened = await createTwoEditorFixture({
                ...options,
                ownedAttributes: true,
                editor: editorOptions,
              });
              try {
                await reopened.settle();
                for (const editor of reopened.editors) {
                  expect(collectRevisionSites(editor.surface!.session.part())).toEqual([]);
                  expect(groups(editor)).toEqual([]);
                  expect(widths(editor)).toEqual(
                    action === 'reject' ? ['3000', '3000'] : ['3500', '2500']
                  );
                  expect(word(editor, 'Two peer')).toBeDefined();
                }
              } finally {
                reopened.destroy();
              }
            }
          } finally {
            pair.destroy();
          }
        }, 30000);
