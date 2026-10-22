import { applyTreeOp, canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { nodeChildren } from '../../test/support/native-table-oracle';
import { expect, test } from 'bun:test';
import { readOoxmlPackage, bodyStoryRoot, paragraphTextOf } from '@docx-editor.dev/core/store';
import { propertyEditor, propertyContext } from '../../test/support/property-execution';
import { document_ } from '../../test/support/compound-target-table';
import { groupToken, groups } from '../../test/support/review-groups';
import { createEditPlan, decodeEditPlan } from './plan';
import { applyDirectPlan } from '../ai/apply-direct';
import { replaceOwnedProposals } from '../review/replace-owned-proposals';
const revision = { documentId: 'doc', epoch: 'mount', revision: 0 };
const base = {
  kind: 'table-geometry',
  partName: '/word/document.xml',
  intent: { tableId: 'table', kind: 'merge', region: { row: 0, column: 0, rows: 1, columns: 2 } },
};
const malformed = [
  { ...base, intent: { ...base.intent, cellId: 'preview-cell' } },
  { ...base, intent: { ...base.intent, region: { row: -1, column: 0, rows: 1, columns: 2 } } },
  {
    ...base,
    intent: {
      ...base.intent,
      region: { row: Number.MAX_SAFE_INTEGER, column: 0, rows: 2, columns: 2 },
    },
  },
  { ...base, intent: { ...base.intent, region: { row: 0, column: 62, rows: 1, columns: 2 } } },
  { ...base, intent: { ...base.intent, kind: 'split', rows: 1, columns: 2, cells: [] } },
  {
    ...base,
    intent: {
      ...base.intent,
      kind: 'split',
      rows: 1,
      columns: 2,
      cells: [[{ type: 'paragraph', children: 'malformed' }]],
    },
  },
];
for (const [index, edit] of malformed.entries())
  test(`deferred table geometry rejects malformed region or preview identity ${index}`, () => {
    const decoded = decodeEditPlan(
      JSON.stringify({ schemaVersion: 1, revision, operations: [edit] })
    );
    expect(decoded.success).toBe(false);
  });
for (const mode of ['direct', 'tracked'] as const)
  test(`late ${mode} region refusal preserves package, prior redo, groups, and publication count`, () => {
    const read = readOoxmlPackage(document_());
    if (!read.ok) throw new Error(read.reason);
    const editor = propertyEditor(read.package);
    try {
      const session = editor.surface!.session,
        part = session.part(),
        table = nodeChildren(bodyStoryRoot(part)!).find((node) => node.kind === 'table')!;
      const row = table.children.find((node) => node.kind === 'tableRow')!;
      if (!row) throw new Error('Missing row');
      const tail = session.paragraphIds().find((id) => paragraphTextOf(part, id) === 'Tail')!;
      expect(
        editor.surface!.applyAutomationOps(
          () => [{ op: 'insertText', paragraphId: tail, offset: 4, text: ' redo' }],
          { kind: 'body' },
          [],
          false,
          true
        )
      ).toMatchObject({ committed: true, rejected: false });
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(session.canRedo()).toBe(true);
      const before = session.currentPackage(),
        beforeGroups = groups(editor),
        plan = createEditPlan(groupToken(editor), [
          {
            kind: 'insert-row',
            partName: part.name,
            intent: {
              tableId: table.id,
              rowId: row.id,
              where: 'below',
              cells: [[], [], []],
              properties: {},
              payloads: [],
            },
          },
          {
            kind: 'table-geometry',
            partName: part.name,
            intent: {
              tableId: table.id,
              kind: 'merge',
              region: { row: 20, column: 0, rows: 1, columns: 2 },
            },
          },
        ]);
      if (!plan.success) throw new Error(plan.error);
      let publications = 0;
      const stop = session.subscribe(() => publications++);
      const result =
        mode === 'direct'
          ? applyDirectPlan(editor, plan.plan, propertyContext)
          : replaceOwnedProposals(editor, [], plan.plan, propertyContext);
      stop();
      expect(result.success).toBe(false);
      expect(session.currentPackage()).toBe(before);
      expect(groups(editor)).toEqual(beforeGroups);
      expect(publications).toBe(0);
      expect(session.canRedo()).toBe(true);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(paragraphTextOf(session.part(), tail)).toBe('Tail redo');
    } finally {
      editor.destroy();
    }
  });

for (const mode of ['direct', 'tracked'] as const)
  test(`${mode} native fragment admission rejects an invalid text value after split allocation`, () => {
    const opened = readOoxmlPackage(document_(true));
    if (!opened.ok) throw new Error(opened.reason);
    const editor = propertyEditor(opened.package);
    try {
      const session = editor.surface!.session,
        table = nodeChildren(bodyStoryRoot(session.part())!).find((node) => node.kind === 'table')!,
        before = session.currentPackage();
      const plan = createEditPlan(groupToken(editor), [
        {
          kind: 'table-geometry',
          partName: session.part().name,
          intent: {
            tableId: table.id,
            kind: 'split',
            region: { row: 0, column: 0, rows: 1, columns: 2 },
            rows: 1,
            columns: 2,
            cells: [[{ type: 'paragraph', children: [{ type: 'text', value: 1 }] }]],
          },
        },
      ]);
      // Plain JSON/schema admission delegates semantic leaf validation to the native fragment stage.
      if (!plan.success) throw new Error(plan.error);
      expect(decodeEditPlan(JSON.stringify(plan.plan)).success).toBe(true);
      let publications = 0;
      const stop = session.subscribe(() => publications++);
      const result =
        mode === 'direct'
          ? applyDirectPlan(editor, plan.plan, propertyContext)
          : replaceOwnedProposals(editor, [], plan.plan, propertyContext);
      stop();
      expect(result.success).toBe(false);
      expect(session.currentPackage()).toBe(before);
      expect(publications).toBe(0);
      expect(groups(editor)).toEqual([]);
      expect(session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  });

for (const action of ['delete-earlier-row', 'format-covered-cell'] as const)
  test(`direct table commands retain their order after deferred geometry: ${action}`, () => {
    const opened = readOoxmlPackage(document_());
    if (!opened.ok) throw new Error(opened.reason);
    const editor = propertyEditor(opened.package);
    try {
      const session = editor.surface!.session,
        before = session.currentPackage(),
        original = session.part(),
        table = nodeChildren(bodyStoryRoot(original)!).find((node) => node.kind === 'table')!,
        rows = nodeChildren(table).filter((node) => node.kind === 'tableRow'),
        selected = action === 'delete-earlier-row' ? 1 : 0,
        cells = nodeChildren(rows[selected]!).filter((node) => node.kind === 'tableCell');
      const merge = {
        op: 'mergeTableCells' as const,
        tableId: table.id,
        cellIds: cells.slice(0, 2).map((cell) => cell.id),
      };
      const later =
        action === 'delete-earlier-row'
          ? { op: 'deleteTableRow' as const, tableId: table.id, rowId: rows[0]!.id }
          : {
              op: 'setTableCellFill' as const,
              tableId: table.id,
              cellIds: [cells[1]!.id],
              color: { kind: 'hex' as const, value: 'FF0000' },
            };
      const first = applyTreeOp(original, merge);
      if (!first.ok) throw new Error(first.reason);
      const golden = applyTreeOp(first.part, later);
      const plan = createEditPlan(groupToken(editor), [
        {
          kind: 'table-geometry',
          partName: original.name,
          intent: {
            tableId: table.id,
            kind: 'merge',
            region: { row: selected, column: 0, rows: 1, columns: 2 },
          },
        },
        { kind: 'table', partName: original.name, operation: later },
      ]);
      if (!plan.success) throw new Error(plan.error);
      let publications = 0;
      const stop = session.subscribe(() => publications++);
      const actual = applyDirectPlan(editor, plan.plan, propertyContext);
      stop();
      expect(actual.success).toBe(golden.ok);
      if (golden.ok) {
        expect(publications).toBe(1);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(
          canonicalOoxmlFingerprint(golden.part)
        );
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(session.part()).toBe(original);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(
          canonicalOoxmlFingerprint(golden.part)
        );
      } else {
        expect(publications).toBe(0);
        expect(session.currentPackage()).toBe(before);
        expect(session.canUndo()).toBe(false);
      }
    } finally {
      editor.destroy();
    }
  });
