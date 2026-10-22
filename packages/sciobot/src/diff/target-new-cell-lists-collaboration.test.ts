import { nativeGridColumns } from './inserted-column-native';
import { expect, test } from 'bun:test';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput, targetParagraph } from '../../test/support/target-document';
import { table, nativeTable, part } from '../../test/support/target-table-fixture';
import { rowCells, tableRows } from '../../test/support/native-table-oracle';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';

const context = { documentId: 'new-cell-list-session', epoch: '1' };
const lists = [
  '- Bullet one\n- Bullet two',
  '3. Ordered three\n4. Ordered four',
  '- [ ] Task open\n- [x] Task done',
];
const html = (rows: readonly (readonly string[])[]) =>
  '<table>\n' +
  rows
    .map(
      (row) => '<tr>\n' + row.map((cell) => '<td>\n\n' + cell + '\n\n</td>').join('\n') + '\n</tr>'
    )
    .join('\n') +
  '\n</table>\n\nTail';
for (const kind of ['row', 'column'] as const)
  for (const clientIds of [
    [43, 79],
    [79, 43],
  ] as const)
    test(`new ${kind} lists converge with offline peer text, undo/redo, and reopen: ${clientIds}`, async () => {
      const initial = kind === 'row' ? [['A', 'B', 'C']] : [['A'], ['B'], ['C']];
      const input = await targetDocumentInput(table(initial) + targetParagraph('Tail'));
      const pair = await createTwoEditorFixture({
        document: writeOoxmlPackage(input.pkg),
        clientIds,
      });
      try {
        const [alice, bob] = pair.editors;
        const session = alice!.surface!.session;
        const token = readEditorRevision(alice!, context.documentId, context.epoch)!;
        const view = await readWorkingView(session, token, { tableSyntax: 'html' });
        if (!view.success) throw new Error(view.error);
        const target = html(
          kind === 'row' ? [...initial, lists] : initial.map((row, i) => [...row, lists[i]!])
        );
        const plan = planTargetDocument(
          { pkg: session.currentPackage(), revision: token, projection: view.view.projection },
          target
        );
        if (!plan.success) throw new Error(plan.error);
        const tail = session
          .paragraphIds()
          .find((id) => paragraphTextOf(session.part(), id) === 'Tail')!;
        bob!.surface!.setSelection({
          anchor: { paragraphId: tail, offset: 4 },
          head: { paragraphId: tail, offset: 4 },
        });
        await pair.settle();
        pair.pause();
        const result = applyDirectPlan(alice!, plan.plan, context);
        if (!result.success) throw new Error(result.error);
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: tail, offset: 4, text: ' Peer' },
          ]).rejected
        ).toBe(false);
        pair.reconnect(clientIds[0] === 43 ? 0 : 1);
        await pair.settle();
        const texts = async () =>
          Promise.all(
            pair.editors.map(async (editor) => {
              const read = await readWorkingView(
                editor.surface!.session,
                readEditorRevision(editor, context.documentId, context.epoch)!,
                { tableSyntax: 'html' }
              );
              if (!read.success) throw new Error(read.error);
              return read.view.projection.text;
            })
          );
        const landed = await texts();
        expect(landed[0]).toBe(landed[1]);
        expect(landed[0]).toContain('- Bullet one');
        expect(landed[0]).toContain('3. Ordered three');
        expect(landed[0]).toContain('- [x] Task done');
        expect(landed[0]).toEndWith('Tail Peer');
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (item) => item.actorId === 'bob'
        );
        expect(remote).toBeDefined();
        const position = { paragraphId: tail, offset: remote!.head.offset };
        const a = caretAt(alice!.surface!.layout(), position),
          b = caretAt(bob!.surface!.layout(), position);
        expect(a).not.toBeNull();
        expect(b).not.toBeNull();
        expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const undone = await texts();
        expect(undone[0]).toBe(undone[1]);
        expect(undone[0]).not.toContain('Bullet one');
        expect(undone[0]).toEndWith('Tail Peer');
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(await texts()).toEqual(landed);
        const reopened = await Promise.all(
          pair.editors.map(async (editor) => {
            const parsed = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!parsed.ok) throw new Error(parsed.reason);
            return semanticDigest(parsed.package.parts.values());
          })
        );
        expect(reopened[0]).toEqual(reopened[1]);
        // Deleting the inserted owner while another peer edits its paragraph must remove both.
        const current = { ...input, pkg: session.currentPackage() };
        const rows = tableRows(part(current), nativeTable(current).id);
        const owner = kind === 'row' ? rows.at(-1)! : rowCells(rows[0]!).at(-1)!;
        const paragraph = storyParagraphs(kind === 'row' ? rowCells(owner)[0]! : owner)[0]!;
        pair.pause();
        expect(
          alice!.surface!.applyAutomationOps(() => [
            kind === 'row'
              ? { op: 'deleteTableRow', tableId: nativeTable(current).id, rowId: owner.id }
              : {
                  op: 'deleteTableColumn',
                  tableId: nativeTable(current).id,
                  gridColumnId: nativeGridColumns(part(current), nativeTable(current).id).at(-1)!
                    .id,
                },
          ]).rejected
        ).toBe(false);
        expect(
          bob!.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: paragraph.id, offset: 0, text: 'Peer ' },
          ]).rejected
        ).toBe(false);
        pair.reconnect();
        await pair.settle();
        for (const editor of pair.editors)
          expect(findNode(editor.surface!.session.part(), paragraph.id)).toBeNull();
        const removed = await texts();
        expect(removed[0]).toBe(removed[1]);
      } finally {
        pair.destroy();
      }
    });
