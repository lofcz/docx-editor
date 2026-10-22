import { applyTreeOp, withPart, canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { stageTrackedPlan } from '../review/tracked-stage';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { trackTableRowPlan } from '../review/track-table-rows';
import { trackTableColumnPlan } from '../review/track-table-columns';
import { readNumPr } from '@docx-editor.dev/core/layout';
import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  propertyContainer,
  contentControlPropertiesOf,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { table, nativeTable, part, ast } from '../../test/support/target-table-fixture';
import { rowCells, tableRows } from '../../test/support/native-table-oracle';
import { planTargetDocument } from './plan-target';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { applyDirectPlan } from '../ai/apply-direct';
import { createEditPlan } from './plan';

// Word does not store tight/loose list spacing as Markdown list metadata.
const semanticAst = (markdown: string): unknown =>
  JSON.parse(JSON.stringify(ast(markdown), (key, value) => (key === 'spread' ? undefined : value)));

const context = { documentId: 'new-cell-lists', epoch: '1' };
const html = (rows: readonly (readonly string[])[]) =>
  '<table>\n' +
  rows
    .map(
      (row) => '<tr>\n' + row.map((cell) => '<td>\n\n' + cell + '\n\n</td>').join('\n') + '\n</tr>'
    )
    .join('\n') +
  '\n</table>';
const lists = [
  '- Bullet one\n- Bullet two',
  '3. Ordered three\n4. Ordered four',
  '- [ ] Task open\n- [x] Task done',
];

for (const mode of ['direct', 'tracked'] as const)
  for (const kind of ['row', 'column'] as const) {
    test(`${mode} new ${kind} list bindings resolve only after native allocation and survive reopen`, async () => {
      const initial = kind === 'row' ? [['A', 'B', 'C']] : [['A'], ['B'], ['C']];
      const input = await targetDocumentInput(table(initial));
      const editor = createDocxEditor({
        container: document.createElement('div'),
        document: writeOoxmlPackage(input.pkg),
        modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
      });
      try {
        const session = editor.surface!.session;
        const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
        const view = await readWorkingView(session, revision, { tableSyntax: 'html' });
        if (!view.success) throw new Error(view.error);
        const target = html(
          kind === 'row' ? [...initial, lists] : initial.map((row, i) => [...row, lists[i]!])
        );
        const before = writeOoxmlPackage(session.currentPackage());
        const originalMain = canonicalOoxmlFingerprint(session.part());
        const plan = planTargetDocument(
          { pkg: session.currentPackage(), revision, projection: view.view.projection },
          target
        );
        if (!plan.success) throw new Error(plan.error);
        expect(plan.plan.operations).toHaveLength(1);
        const edit = plan.plan.operations[0]!;
        expect(edit.kind).toBe(kind === 'row' ? 'insert-row' : 'insert-column');
        if (edit.kind !== 'insert-row' && edit.kind !== 'insert-column')
          throw new Error('Missing cell insertion');
        expect(edit.intent.lists?.map((binding) => binding.path)).toEqual([
          [0, 0],
          [0, 1],
          [1, 0],
          [1, 1],
          [2, 0],
          [2, 1],
        ]);
        expect(edit.intent.lists?.map((binding) => binding.target.kind)).toEqual([
          'bullet',
          'bullet',
          'ordered',
          'ordered',
          'checklist',
          'checklist',
        ]);
        expect(edit.intent.lists?.[2]?.target).toMatchObject({
          kind: 'ordered',
          level: 0,
          instance: { kind: 'new', start: 3 },
        });
        expect(edit.intent.lists?.[5]?.target).toEqual({
          kind: 'checklist',
          level: 0,
          checked: true,
        });
        expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
        const applied =
          mode === 'direct'
            ? applyDirectPlan(editor, plan.plan, context)
            : (kind === 'row' ? trackTableRowPlan : trackTableColumnPlan)(editor, plan.plan, {
                ...context,
                author: 'ScioBot',
                date: '2026-10-22T08:00:00Z',
              });
        if (!applied.success) throw new Error(applied.error);
        if (mode === 'tracked') {
          expect(editor.getReviewItems()).toHaveLength(kind === 'row' ? 1 : 4);
          expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
          expect(canonicalOoxmlFingerprint(session.part())).toBe(originalMain);
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
        }
        const pkg = session.currentPackage();
        const saved = readOoxmlPackage(writeOoxmlPackage(pkg));
        if (!saved.ok) throw new Error(saved.reason);
        for (const current of [pkg, saved.package]) {
          const readInput = { ...input, pkg: current };
          const rows = tableRows(part(readInput), nativeTable(readInput).id);
          const cells =
            kind === 'row' ? rowCells(rows.at(-1)!) : rows.map((row) => rowCells(row).at(-1)!);
          expect(
            cells.map((cell) =>
              storyParagraphs(cell).map((node) => paragraphTextOf(part(readInput), node.id))
            )
          ).toEqual([
            ['Bullet one', 'Bullet two'],
            ['Ordered three', 'Ordered four'],
            ['☐Task open', '☒Task done'],
          ]);
          expect(
            storyParagraphs(cells[2]!).map((node) => {
              if (node.kind === 'textValue') throw new Error('Missing task paragraph');
              const control = node.children.find((child) => child.kind === 'contentControl');
              if (!control) throw new Error('Missing task control');
              return contentControlPropertiesOf(control).checkbox?.checked;
            })
          ).toEqual([false, true]);
          expect(
            storyParagraphs(cells[0]!).map(
              (node) => readNumPr([propertyContainer(node, 'paragraphProperties', 'pPr')!])?.ilvl
            )
          ).toEqual([0, 0]);
          const ids = cells
            .slice(0, 2)
            .map((cell) =>
              storyParagraphs(cell).map(
                (node) => readNumPr([propertyContainer(node, 'paragraphProperties', 'pPr')!])!.numId
              )
            );
          expect(ids[0]![0]).toBe(ids[0]![1]);
          expect(ids[1]![0]).toBe(ids[1]![1]);
          expect(ids[0]![0]).not.toBe(ids[1]![0]);
        }
        const projected = await readWorkingView(
          session,
          readEditorRevision(editor, context.documentId, context.epoch)!,
          { tableSyntax: 'html' }
        );
        if (!projected.success) throw new Error(projected.error);
        expect(semanticAst(projected.view.projection.text)).toEqual(semanticAst(target));
        expect(projected.view.projection.text).toContain('- Bullet one');
        expect(projected.view.projection.text).toContain('3. Ordered three');
        expect(projected.view.projection.text).toContain('- [x] Task done');
        const invalid = {
          ...edit,
          intent: {
            ...edit.intent,
            lists: [
              {
                path: [99, 0],
                target: { kind: 'bullet', level: 0, instance: { kind: 'new', key: 'x' } },
              },
            ],
          },
        };
        const bad = createEditPlan(plan.plan.revision, [invalid] as unknown as Parameters<
          typeof createEditPlan
        >[1]);
        expect(bad.success).toBe(false);
      } finally {
        editor.destroy();
      }
    });
  }

for (const kind of ['row', 'column'] as const)
  test(`staged tracked ${kind} publishes new-cell lists into detached package only`, async () => {
    const initial = kind === 'row' ? [['A', 'B', 'C']] : [['A'], ['B'], ['C']];
    const input = await targetDocumentInput(table(initial));
    const target = html(
      kind === 'row' ? [...initial, lists] : initial.map((row, index) => [...row, lists[index]!])
    );
    const planned = planTargetDocument(input, target);
    if (!planned.success) throw new Error(planned.error);
    const before = writeOoxmlPackage(input.pkg);
    let pkg = input.pkg;
    const receipts = stageTrackedPlan(
      {
        currentPackage: () => pkg,
        apply: (partName, op) => {
          const applied = applyTreeOp(pkg.parts.get(partName)!, op);
          if (!applied.ok) throw new Error(applied.reason);
          pkg = withPart(pkg, applied.part);
        },
        edit: (change) => {
          pkg = change(pkg);
        },
      },
      planned.plan,
      { author: 'ScioBot', date: '2026-10-22T08:00:00Z' },
      true
    );
    expect(receipts.length).toBeGreaterThan(0);
    const projected = await projectSourcePackage(pkg);
    expect(projected.text).toContain('Bullet one');
    expect(projected.text).toContain('Ordered three');
    expect(projected.text).toContain('- [x] Task done');
    expect(writeOoxmlPackage(input.pkg)).toEqual(before);
  });
