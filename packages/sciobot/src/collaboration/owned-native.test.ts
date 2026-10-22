import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, propertyContainer, runsCovering } from '@docx-editor.dev/core/store';
import { strFromU8, unzipSync } from 'fflate';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  CELL_CONTEXT,
  CELL_ORIGINAL,
  cellPlan,
  cellReviewDocument,
  cellScopeFor,
  cellViews,
  scopeOnly,
} from '../../test/support/table-review';
import { directParagraphTextOps, planParagraphText } from '../diff/paragraph-text';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { trackTableCellPlan } from '../review/track-table-cells';

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    for (const marked of [false, true])
      for (const action of ['accept', 'reject'] as const)
        test(`owned native bold excludes disconnected ${marked ? 'italic' : 'plain'} suffix through ${action}; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
          const pair = await createTwoEditorFixture({
            document: cellReviewDocument(),
            clientIds,
            ownedAttributes: true,
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
          });
          const stops: (() => void)[] = [];
          try {
            const [alice, bob] = pair.editors;
            const scope = cellScopeFor(alice!);
            bob!.surface!.setSelection({
              anchor: { paragraphId: scope.paragraphId, offset: 3 },
              head: { paragraphId: scope.paragraphId, offset: 3 },
            });
            await pair.settle();
            const plan = await cellPlan(alice!, scope, [
              {
                paragraphId: scope.paragraphId,
                text: 'Two',
                formats: [{ start: 0, end: 3, target: { properties: [{ localName: 'b' }] } }],
              },
            ]);
            pair.pause();
            expect(trackTableCellPlan(alice!, plan, scopeOnly(scope), CELL_CONTEXT).success).toBe(
              true
            );
            const part = bob!.surface!.session.part();
            const direct = planParagraphText(
              {
                revision: readEditorRevision(bob!, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
                partName: part.name,
                paragraphId: scope.paragraphId,
                text: paragraphTextOf(part, scope.paragraphId)!,
              },
              'Two Bob'
            );
            if (!direct.success) throw new Error(direct.error);
            const ops = direct.plan.operations.flatMap((edit) =>
              edit.kind === 'text' ? directParagraphTextOps(edit) : []
            );
            expect(
              bob!.surface!.applyAutomationOps(() => [
                ...ops,
                ...(marked
                  ? [
                      {
                        op: 'setRunProperties' as const,
                        paragraphId: scope.paragraphId,
                        start: 4,
                        end: 7,
                        properties: [{ localName: 'i' }],
                      },
                    ]
                  : []),
              ]).rejected
            ).toBe(false);
            const foreign: string[] = [];
            let publications = 0;
            for (const editor of pair.editors)
              stops.push(
                editor.on('change', () => {
                  publications++;
                  const current = editor.surface!.session.part();
                  for (const run of runsCovering(current, scope.paragraphId, 4, 7)) {
                    const properties = propertyContainer(run, 'runProperties', 'rPr');
                    if (
                      properties &&
                      properties.kind !== 'textValue' &&
                      properties.children.some(
                        (child) =>
                          child.kind !== 'textValue' && ['b', 'rPrChange'].includes(child.localName)
                      )
                    )
                      foreign.push(run.id);
                  }
                })
              );
            pair.reconnect(first);
            await pair.settle();
            const peer = marked ? ' *Bob*' : ' Bob';
            const original = CELL_ORIGINAL.replace('Two', 'Two' + peer);
            const proposed = CELL_ORIGINAL.replace('Two', '**Two**' + peer);
            for (const editor of pair.editors)
              expect(await cellViews(editor)).toEqual([original, proposed]);
            const remote = pair.rooms[0]!.session.remoteSelections().find(
              (selection) => selection.actorId === 'bob'
            );
            expect(remote).toBeDefined();
            expect(remote!.head.offset).toBe(bob!.surface!.state().selection!.head.offset);
            const position = { paragraphId: scope.paragraphId, offset: remote!.head.offset };
            const leftCaret = caretAt(alice!.surface!.layout(), position),
              rightCaret = caretAt(bob!.surface!.layout(), position);
            expect(leftCaret).not.toBeNull();
            expect(rightCaret).not.toBeNull();
            expect({ x: leftCaret!.x, y: leftCaret!.y }).toEqual({
              x: rightCaret!.x,
              y: rightCaret!.y,
            });
            const saved = await pair.savedDigests();
            expect(saved[0]).toEqual(saved[1]);
            const reopened = await createTwoEditorFixture({
              document: cellReviewDocument(),
              ownedAttributes: true,
              roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
              editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
            });
            try {
              for (const editor of reopened.editors)
                expect(await cellViews(editor)).toEqual([original, proposed]);
              expect(
                reopened.editors[0]!.exec({ type: 'resolveAllReviewChanges', action }).ok
              ).toBe(true);
              await reopened.settle();
              const resolved = action === 'accept' ? proposed : original;
              for (const editor of reopened.editors)
                expect(await cellViews(editor)).toEqual([resolved, resolved]);
              expect(reopened.editors[0]!.exec({ type: 'undo' }).ok).toBe(true);
              await reopened.settle();
              for (const editor of reopened.editors)
                expect(await cellViews(editor)).toEqual([original, proposed]);
              reopened.ownership.forEach((handle) => handle.assertHealthy());
            } finally {
              reopened.destroy();
            }
            expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(await cellViews(editor)).toEqual([original, original]);
            expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(await cellViews(editor)).toEqual([original, proposed]);
            expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
            await pair.settle();
            const expected = action === 'accept' ? proposed : original;
            for (const editor of pair.editors) {
              expect(await cellViews(editor)).toEqual([expected, expected]);
              expect(editor.getReviewItems()).toHaveLength(0);
            }
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(await cellViews(editor)).toEqual([original, proposed]);
            expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(await cellViews(editor)).toEqual([expected, expected]);
            const finalSaved = await pair.savedDigests();
            expect(finalSaved[0]).toEqual(finalSaved[1]);
            for (const editor of pair.editors) {
              const bytes = new Uint8Array(await editor.save());
              for (const [name, content] of Object.entries(unzipSync(bytes)))
                if (name.endsWith('.xml'))
                  expect(strFromU8(content)).not.toContain('sciobot-owned');
              const docxReopen = await createTwoEditorFixture({
                document: bytes,
                ownedAttributes: true,
                editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
              });
              try {
                for (const item of docxReopen.editors)
                  expect(await cellViews(item)).toEqual([expected, expected]);
              } finally {
                docxReopen.destroy();
              }
            }
            expect(publications).toBeGreaterThan(0);
            expect(foreign).toEqual([]);
            pair.ownership.forEach((handle) => handle.assertHealthy());
          } finally {
            stops.forEach((stop) => stop());
            pair.destroy();
          }
        }, 30_000);

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const first of [0, 1] as const)
    test(`owned native format survives owner deletion history; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
      const pair = await createTwoEditorFixture({
        document: cellReviewDocument(),
        ownedAttributes: true,
        clientIds,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
      });
      try {
        const [alice, bob] = pair.editors,
          scope = cellScopeFor(bob!);
        const plan = await cellPlan(bob!, scope, [
          {
            paragraphId: scope.paragraphId,
            text: 'Two',
            formats: [{ start: 0, end: 3, target: { properties: [{ localName: 'b' }] } }],
          },
        ]);
        pair.pause();
        expect(trackTableCellPlan(bob!, plan, scopeOnly(scope), CELL_CONTEXT).success).toBe(true);
        expect(
          alice!.surface!.applyAutomationOps(() => [
            { op: 'deleteTableRow', tableId: scope.tableId, rowId: scope.rowId },
          ]).rejected
        ).toBe(false);
        pair.reconnect(first);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(paragraphTextOf(editor.surface!.session.part(), scope.paragraphId)).toBeNull();
          expect(editor.getReviewItems()).toHaveLength(0);
        }
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await cellViews(editor)).toEqual([
            CELL_ORIGINAL,
            CELL_ORIGINAL.replace('Two', '**Two**'),
          ]);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(paragraphTextOf(editor.surface!.session.part(), scope.paragraphId)).toBeNull();
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]);
        pair.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 30_000);
