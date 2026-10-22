import { directParagraphProperties } from '@docx-editor.dev/core/store';
import { namedChild } from '../../test/support/native-table-oracle';
import { formatAttr } from '../../test/support/table-format-fixture';
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

const context = { documentId: 'new-cell-property-session', epoch: '1' };
const lists = ['New one', 'New two', 'New three'];
const splitXml =
  '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="3"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
const html = (rows: readonly (readonly string[])[]) =>
  '<table>\n' +
  rows
    .map(
      (row) =>
        '<tr>\n' +
        row
          .map(
            (cell) =>
              '<td' +
              (cell.startsWith('New')
                ? ' style="background-color: #123456; border: 1pt solid red; text-align: center; vertical-align: bottom"'
                : '') +
              '>\n\n' +
              cell +
              '\n\n</td>'
          )
          .join('\n') +
        '\n</tr>'
    )
    .join('\n') +
  '\n</table>\n\nTail';
for (const kind of ['row', 'column', 'split'] as const)
  for (const clientIds of [
    [43, 79],
    [79, 43],
  ] as const)
    test(`new ${kind} properties converge with offline peer text, undo/redo, and reopen: ${clientIds}`, async () => {
      const initial = kind === 'row' ? [['A', 'B', 'C']] : [['A'], ['B'], ['C']];
      const input = await targetDocumentInput(
        (kind === 'split' ? splitXml : table(initial)) + targetParagraph('Tail')
      );
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
          kind === 'split'
            ? [['A', 'New one', 'New two']]
            : kind === 'row'
              ? [...initial, lists]
              : initial.map((row, i) => [...row, lists[i]!])
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
        expect(landed[0]).toContain('New one');
        for (const editor of pair.editors) {
          const source = { ...input, pkg: editor.surface!.session.currentPackage() };
          const rows = tableRows(part(source), nativeTable(source).id);
          const cells =
            kind === 'row'
              ? rowCells(rows.at(-1)!)
              : kind === 'split'
                ? rowCells(rows[0]!).slice(1)
                : rows.map((row) => rowCells(row).at(-1)!);
          for (const cell of cells) {
            const props = namedChild(cell, 'tcPr')!;
            expect(formatAttr(namedChild(props, 'shd')!, 'fill')).toBe('123456');
            expect(formatAttr(namedChild(props, 'vAlign')!, 'val')).toBe('bottom');
            expect(formatAttr(namedChild(namedChild(props, 'tcBorders')!, 'top')!, 'color')).toBe(
              'FF0000'
            );
            expect(
              storyParagraphs(cell).map(
                (p) =>
                  directParagraphProperties(part(source), p.id).find(
                    (value) => value.localName === 'jc'
                  )?.attributes?.val
              )
            ).toEqual(['center']);
          }
        }
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
        expect(undone[0]).not.toContain('New one');
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
