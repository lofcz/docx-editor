import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { groups } from '../../test/support/review-groups';
import {
  cellReviewDocument,
  cellScopeFor,
  childNamed,
  CELL_CONTEXT,
} from '../../test/support/table-review';
import { planTableFormat } from '../diff/table-format';
import { readEditorRevision } from '../projection/revision-cache';
import { acceptReviewGroup } from './accept-group';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableFormatPlan } from './track-table-format';

const editorOptions = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
const children = (node: OoxmlNode | undefined) =>
  node && node.kind !== 'textValue' ? node.children : [];
const attribute = (node: OoxmlNode | undefined, name: string) =>
  node && node.kind !== 'textValue'
    ? node.attributes.find((a) => a.localName === name)?.value
    : undefined;
function table(part: OoxmlPart) {
  return children(childNamed(part.root, 'body')).find((node) => node.kind === 'table')!;
}
/** Read the actual Word property snapshots, independently of application review projections. */
function wordWidths(part: OoxmlPart, original: boolean) {
  const grid = childNamed(table(part), 'tblGrid'),
    snapshot = childNamed(childNamed(grid, 'tblGridChange'), 'tblGrid');
  return children(original && snapshot ? snapshot : grid)
    .filter((node) => node.kind !== 'textValue' && node.localName === 'gridCol')
    .map((node) => attribute(node, 'w'));
}
function members(part: OoxmlPart) {
  return collectRevisionSites(part)
    .map(({ node }) => [
      node.localName,
      ...['id', 'author', 'date'].map((key) => attribute(node, key)),
    ])
    .sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
}
function hasParagraphText(part: OoxmlPart, text: string) {
  const visit = (node: OoxmlNode): boolean =>
    (node.kind === 'paragraph' && paragraphTextOf(part, node.id) === text) ||
    children(node).some(visit);
  return visit(part.root);
}
function assertWord(part: OoxmlPart, pending: boolean, membership: ReturnType<typeof members>) {
  expect(wordWidths(part, false)).toEqual(['3500', '2500']);
  expect(wordWidths(part, true)).toEqual(pending ? ['3000', '3000'] : ['3500', '2500']);
  expect(members(part)).toEqual(pending ? membership : []);
  const rows = children(table(part)).filter((node) => node.kind === 'tableRow');
  expect(rows).toHaveLength(2);
  for (const row of rows) {
    const cells = children(row).filter((node) => node.kind === 'tableCell');
    cells.forEach((cell, index) => {
      const properties = childNamed(cell, 'tcPr'),
        before = childNamed(childNamed(properties, 'tcPrChange'), 'tcPr');
      expect(attribute(childNamed(properties, 'tcW'), 'w')).toBe(index ? '2500' : '3500');
      expect(attribute(childNamed(properties, 'shd'), 'fill')).toBe('EEEEEE');
      if (pending) {
        expect(before).toBeDefined();
        expect(childNamed(before, 'tcW')).toBeUndefined();
        expect(attribute(childNamed(before, 'shd'), 'fill')).toBe('EEEEEE');
      }
    });
  }
}
function propose(editor: DocxEditorInstance) {
  const scope = cellScopeFor(editor),
    planned = planTableFormat(
      {
        part: editor.surface!.session.part(),
        revision: readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
        tableId: scope.tableId,
      },
      { op: 'setTableProperties', columnWidthsTwips: [3500, 2500] }
    );
  if (!planned.success) throw new Error(planned.error);
  expect(trackTableFormatPlan(editor, planned.plan, CELL_CONTEXT).success).toBe(true);
}
function accept(editor: DocxEditorInstance, id: string) {
  expect(
    acceptReviewGroup(
      editor,
      id,
      readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
      CELL_CONTEXT
    ).success
  ).toBe(true);
}
function type(editor: DocxEditorInstance, text: string) {
  const surface = editor.surface!,
    paragraphId = cellScopeFor(editor).paragraphId;
  surface.setSelection({ anchor: { paragraphId, offset: 0 }, head: { paragraphId, offset: 0 } });
  surface.insertPlainText(text);
  surface.flushPendingInput();
}
async function assertReopen(
  pair: Awaited<ReturnType<typeof createTwoEditorFixture>>,
  id: string,
  pending: boolean,
  membership: ReturnType<typeof members>,
  expectedText = 'Peer: Two'
) {
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  for (const editor of pair.editors) {
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    const part = saved.package.parts.get(saved.package.mainDocumentPart)!;
    assertWord(part, pending, membership);
    expect(hasParagraphText(part, expectedText)).toBe(true);
  }
  const docx = new Uint8Array(await pair.editors[0]!.save()),
    snapshot = Y.encodeStateAsUpdate(pair.documents[1]!);
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
        assertWord(editor.surface!.session.part(), pending, membership);
        expect(groups(editor).map((group) => group.id)).toEqual(pending ? [id] : []);
        expect(
          editor
            .surface!.session.paragraphIds()
            .some((pid) => paragraphTextOf(editor.surface!.session.part(), pid) === expectedText)
        ).toBe(true);
      }
    } finally {
      reopened.destroy();
    }
  }
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const firstUndo of [0, 1] as const)
      for (const firstRedo of [0, 1] as const)
        test(`native property group redo resolves every Word member with independent peer history; ${clients}/${receiver}/undo=${firstUndo}/redo=${firstRedo}`, async () => {
          const pair = await createTwoEditorFixture({
            document: cellReviewDocument(),
            clientIds: clients,
            ownedAttributes: true,
            editor: editorOptions,
          });
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
            propose(alice);
            await pair.settle();
            const id = groups(alice)[0]!.id,
              membership = members(alice.surface!.session.part());
            expect(membership.map((row) => row[0])).toEqual([
              'tblGridChange',
              'tblPrChange',
              'tcPrChange',
              'tcPrChange',
              'tcPrChange',
              'tcPrChange',
            ]);
            pair.pause();
            accept(alice, id);
            type(bob, 'Peer: ');
            accept(bob, id);
            pair.reconnect(receiver);
            await pair.settle();
            const assertState = (pending: boolean, text = 'Peer: Two') => {
              for (const editor of pair.editors) {
                assertWord(editor.surface!.session.part(), pending, membership);
                expect(groups(editor).map((group) => group.id)).toEqual(pending ? [id] : []);
                expect(
                  editor
                    .surface!.session.paragraphIds()
                    .some((pid) => paragraphTextOf(editor.surface!.session.part(), pid) === text)
                ).toBe(true);
              }
            };
            const step = async (actor: number, action: 'undo' | 'redo', pending: boolean) => {
              expect(pair.editors[actor]!.exec({ type: action }).ok).toBe(true);
              await pair.settle();
              assertState(pending);
            };
            assertState(false);
            await step(firstUndo, 'undo', false);
            await step(1 - firstUndo, 'undo', true);
            await assertReopen(pair, id, true, membership);
            await step(firstRedo, 'redo', false);
            await step(1 - firstRedo, 'redo', false);
            await step(firstRedo, 'undo', false);
            await step(1 - firstRedo, 'undo', true);
            // The next history frames belong to prior typing and the proposal, with no repair step.
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            assertState(true, 'Two');
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            assertState(true);
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors) {
              expect(members(editor.surface!.session.part())).toEqual([]);
              expect(wordWidths(editor.surface!.session.part(), false)).toEqual(['3000', '3000']);
              expect(wordWidths(editor.surface!.session.part(), true)).toEqual(['3000', '3000']);
            }
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            assertState(true);
            await step(0, 'redo', false);
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
            await assertReopen(pair, id, false, membership);
          } finally {
            pair.destroy();
          }
        }, 30000);

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`native property redo retains peer ownership during row deletion and peer cell input; ${clients}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: cellReviewDocument(),
        clientIds: clients,
        ownedAttributes: true,
        editor: editorOptions,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        propose(alice);
        await pair.settle();
        const id = groups(alice)[0]!.id,
          membership = members(alice.surface!.session.part());
        pair.pause();
        accept(alice, id);
        accept(bob, id);
        pair.reconnect(receiver);
        await pair.settle();
        pair.pause();
        const scope = cellScopeFor(alice);
        expect(
          alice.surface!.applyAutomationOps(
            () => [{ op: 'deleteTableRow', tableId: scope.tableId, rowId: scope.rowId }],
            { kind: 'body' },
            undefined,
            false,
            true
          ).committed
        ).toBe(true);
        type(bob, 'Peer: ');
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(
            children(table(editor.surface!.session.part())).filter((n) => n.kind === 'tableRow')
          ).toHaveLength(1);
          expect(groups(editor)).toEqual([]);
        }
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          assertWord(editor.surface!.session.part(), false, membership);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          assertWord(editor.surface!.session.part(), false, membership);
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          assertWord(editor.surface!.session.part(), true, membership);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          assertWord(editor.surface!.session.part(), false, membership);
        await assertReopen(pair, id, false, membership, 'Two');
      } finally {
        pair.destroy();
      }
    }, 30000);

for (const actor of [0, 1] as const)
  test(`native property redo converges when its normalization update arrives before its ownership intents; actor=${actor}`, async () => {
    const pair = await createTwoEditorFixture({
      document: cellReviewDocument(),
      clientIds: [41, 73],
      ownedAttributes: true,
      editor: editorOptions,
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      propose(alice);
      await pair.settle();
      const id = groups(alice)[0]!.id,
        membership = members(alice.surface!.session.part());
      pair.pause();
      accept(alice, id);
      type(bob, 'Peer: ');
      accept(bob, id);
      pair.reconnect();
      await pair.settle();
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      pair.pause();
      const updates: Uint8Array[] = [],
        collect = (update: Uint8Array) => updates.push(update);
      pair.documents[actor]!.on('update', collect);
      try {
        expect(pair.editors[actor]!.exec({ type: 'redo' }).ok).toBe(true);
        pair.flush();
      } finally {
        pair.documents[actor]!.off('update', collect);
      }
      if (actor === 0) expect(updates.length).toBeGreaterThan(1);
      for (const update of [...updates].reverse())
        Y.applyUpdate(pair.documents[1 - actor]!, update, 'relay');
      pair.reconnect(actor);
      await pair.settle();
      for (const editor of pair.editors)
        assertWord(editor.surface!.session.part(), false, membership);
      await assertReopen(pair, id, false, membership);
    } finally {
      pair.destroy();
    }
  }, 30000);
