import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import {
  CELL_CONTEXT,
  CELL_ORIGINAL,
  cellParagraph,
  cellPlan,
  cellReviewDocument,
  cellScopeFor,
  cellViews,
  childNamed,
  directCellResult,
  scopeOnly,
} from '../../test/support/table-review';
import { createEditPlan, type CanonicalEditPlan } from '../diff/plan';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableCellPlan } from './track-table-cells';

const mount = (bytes = cellReviewDocument(), review = true) =>
  createDocxEditor({
    container: document.createElement('div'),
    document: bytes,
    author: 'User',
    modules: review ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [],
  });
const bold = [{ start: 0, end: 3, target: { properties: [{ localName: 'b' }] } }];

for (const action of ['accept', 'reject'] as const)
  test(`C01 cell ${action} keeps native neighbors and restores pending state through undo/redo`, async () => {
    const editor = mount();
    try {
      const scope = cellScopeFor(editor),
        session = editor.surface!.session,
        original = session.part();
      const table = findNode(original, scope.tableId)!;
      if (table.kind !== 'table') throw new Error('Missing table');
      const neighbors = table.children
        .filter((node) => node.kind === 'tableRow')
        .flatMap((row) =>
          row.children.filter((node) => node.kind === 'tableCell' && node.id !== scope.cellId)
        );
      const plan = await cellPlan(editor, scope, [
        { paragraphId: scope.paragraphId, text: 'Changed' },
      ]);
      const result = trackTableCellPlan(editor, plan, scopeOnly(scope), CELL_CONTEXT);
      if (!result.success) throw new Error(`${result.code}: ${result.error}`);
      expect(result.receiptError).toBeUndefined();
      expect(result.proposals.map((item) => item.kind).sort()).toEqual(['delete', 'insert']);
      expect(
        result.proposals
          .flatMap((item) => item.addresses)
          .every(
            (address) => address.author === 'ScioBot' && address.date === '2026-10-19T09:00:00.000Z'
          )
      ).toBe(true);
      expect(editor.getReviewItems()).toHaveLength(1);
      expect(await cellViews(editor)).toEqual([
        CELL_ORIGINAL,
        CELL_ORIGINAL.replace('Two', 'Changed'),
      ]);
      const pending = canonicalOoxmlFingerprint(session.part());
      for (const neighbor of neighbors)
        expect(canonicalOoxmlFingerprint(findNode(session.part(), neighbor.id)!)).toBe(
          canonicalOoxmlFingerprint(neighbor)
        );
      for (const [id, name] of [
        [scope.tableId, 'tblGrid'],
        [scope.tableId, 'tblPr'],
        [scope.cellId, 'tcPr'],
        [scope.rowId, 'trPr'],
      ])
        expect(canonicalOoxmlFingerprint(childNamed(findNode(session.part(), id!), name!)!)).toBe(
          canonicalOoxmlFingerprint(childNamed(findNode(original, id!), name!)!)
        );
      expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
      expect(paragraphTextOf(session.part(), scope.paragraphId)).toBe(
        action === 'accept' ? 'Changed' : 'Two'
      );
      expect(findNode(session.part(), scope.cellId)?.kind).toBe('tableCell');
      expect(editor.getReviewItems()).toHaveLength(0);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(paragraphTextOf(session.part(), scope.paragraphId)).toBe(
        action === 'accept' ? 'Changed' : 'Two'
      );
    } finally {
      editor.destroy();
    }
  });

test('combined cell text/mark proposals use one history entry and match direct native acceptance', async () => {
  const editor = mount();
  try {
    const scope = cellScopeFor(editor),
      session = editor.surface!.session;
    const original = canonicalOoxmlFingerprint(session.part());
    const plan = await cellPlan(editor, scope, [
      { paragraphId: scope.paragraphId, text: 'Changed', formats: bold },
    ]);
    const expected = directCellResult(session.part(), plan);
    const result = trackTableCellPlan(editor, plan, scopeOnly(scope), CELL_CONTEXT);
    if (!result.success) throw new Error(result.error);
    expect(result.proposals.some((item) => item.kind === 'format')).toBe(true);
    expect(result.proposals.some((item) => item.formattingKind === 'rPrChange')).toBe(true);
    expect(serializeOoxmlPart(session.part())).toContain('w:rPrChange');
    expect((await cellViews(editor))[0]).toBe(CELL_ORIGINAL);
    expect((await cellViews(editor))[1]).toContain('**Changed**');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(original);
    expect(session.canUndo()).toBe(false);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const pending = canonicalOoxmlFingerprint(session.part());
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(canonicalOoxmlFingerprint(expected));
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(original);
  } finally {
    editor.destroy();
  }
});

test('mark-only cell proposals preserve text and return native formatting receipts', async () => {
  const editor = mount();
  try {
    const scope = cellScopeFor(editor),
      session = editor.surface!.session;
    const before = canonicalOoxmlFingerprint(session.part());
    const plan = await cellPlan(editor, scope, [{ paragraphId: scope.paragraphId, formats: bold }]);
    const result = trackTableCellPlan(editor, plan, scopeOnly(scope), CELL_CONTEXT);
    if (!result.success) throw new Error(result.error);
    expect(result.proposals).toHaveLength(1);
    expect(result.proposals[0]).toMatchObject({
      kind: 'format',
      formattingChanges: [{ property: 'bold', value: 'true' }],
    });
    expect(paragraphTextOf(session.part(), scope.paragraphId)).toBe('Two');
    expect(await cellViews(editor)).toEqual([
      CELL_ORIGINAL,
      CELL_ORIGINAL.replace('Two', '**Two**'),
    ]);
    expect(editor.rejectReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
  } finally {
    editor.destroy();
  }
});

test('pending cell text/marks and native revision addresses survive save/reopen', async () => {
  const editor = mount();
  let reopened: DocxEditorInstance | undefined;
  try {
    const scope = cellScopeFor(editor);
    const result = trackTableCellPlan(
      editor,
      await cellPlan(editor, scope, [
        { paragraphId: scope.paragraphId, text: 'Changed', formats: bold },
      ]),
      scopeOnly(scope),
      CELL_CONTEXT
    );
    if (!result.success) throw new Error(result.error);
    const addresses = result.proposals
      .flatMap((item) => item.addresses)
      .map((address) => JSON.stringify(address))
      .sort();
    reopened = mount(new Uint8Array(await editor.save()));
    expect(await cellViews(reopened)).toEqual(await cellViews(editor));
    expect(
      reopened
        .surface!.session.reviewItems()
        .flatMap((item) => (item.kind === 'revision' ? item.addresses : []))
        .map((address) => JSON.stringify(address))
        .sort()
    ).toEqual(addresses);
    expect(reopened.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
    expect(await cellViews(reopened)).toEqual([CELL_ORIGINAL, CELL_ORIGINAL]);
  } finally {
    reopened?.destroy();
    editor.destroy();
  }
});

test('multi-paragraph cell text/marks and manual breaks never join neighboring cells', async () => {
  const editor = mount(cellReviewDocument(cellParagraph('Two') + cellParagraph('More')));
  try {
    const scope = cellScopeFor(editor),
      session = editor.surface!.session;
    const more = session
      .paragraphIds()
      .find((id) => paragraphTextOf(session.part(), id) === 'More')!;
    const original = canonicalOoxmlFingerprint(session.part());
    const result = trackTableCellPlan(
      editor,
      await cellPlan(editor, scope, [
        { paragraphId: scope.paragraphId, text: '😀\tA\nB' },
        {
          paragraphId: more,
          text: 'More updated',
          formats: [{ start: 0, end: 4, target: { properties: [{ localName: 'i' }] } }],
        },
      ]),
      scopeOnly(scope),
      CELL_CONTEXT
    );
    if (!result.success) throw new Error(result.error);
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), scope.paragraphId)).toBe('😀\tA\nB');
    expect(paragraphTextOf(session.part(), more)).toBe('More updated');
    const owner = findNode(session.part(), scope.cellId)!;
    if (owner.kind !== 'tableCell') throw new Error('Missing cell');
    expect(
      owner.children.filter((node) => node.kind === 'paragraph').map((node) => node.id)
    ).toEqual([scope.paragraphId, more]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(original);
  } finally {
    editor.destroy();
  }
});

for (const mode of ['editing', 'suggesting'] as const)
  test(`cell proposals preserve ${mode} mode and AI attribution`, async () => {
    const editor = mount();
    try {
      expect(editor.setEditingMode(mode).ok).toBe(true);
      const scope = cellScopeFor(editor);
      const result = trackTableCellPlan(
        editor,
        await cellPlan(editor, scope, [
          { paragraphId: scope.paragraphId, text: 'Changed', formats: bold },
        ]),
        scopeOnly(scope),
        CELL_CONTEXT
      );
      if (!result.success) throw new Error(result.error);
      expect(
        result.proposals
          .flatMap((item) => item.addresses)
          .every((address) => address.author === 'ScioBot')
      ).toBe(true);
      expect(editor.surface!.editingMode()).toBe(mode === 'editing' ? 'edit' : 'suggest');
    } finally {
      editor.destroy();
    }
  });

test('invalid or stale cell batches refuse before writing any text, property, history, or selection', async () => {
  const editor = mount();
  try {
    const scope = cellScopeFor(editor),
      session = editor.surface!.session;
    const good = await cellPlan(editor, scope, [
      { paragraphId: scope.paragraphId, text: 'Changed', formats: bold },
    ]);
    const neighbor = cellScopeFor(editor, 'One');
    const format = good.operations.find((edit) => edit.kind === 'format')!;
    if (format.kind !== 'format') throw new Error('Missing format');
    const snapshot = () => [
      canonicalOoxmlFingerprint(session.part()),
      session.packageRevision(),
      session.canUndo(),
      session.canRedo(),
      JSON.stringify(editor.surface!.state().selection),
    ];
    const before = snapshot();
    const invalidPlans: CanonicalEditPlan[] = [
      { ...good, revision: { ...good.revision, epoch: 'old-mount' } },
      {
        ...good,
        operations: [
          ...good.operations,
          {
            kind: 'text',
            partName: session.part().name,
            paragraphId: neighbor.paragraphId,
            start: 0,
            end: 3,
            before: 'One',
            after: 'Bad',
          },
        ],
      },
      { ...good, operations: [...good.operations, format] },
      {
        ...good,
        operations: [
          {
            ...format,
            changes: [
              { name: 'b', before: { present: true, value: {} }, after: { present: false } },
            ],
          },
          ...good.operations.filter((edit) => edit.kind === 'text'),
        ],
      },
      {
        ...good,
        operations: [
          {
            ...format,
            changes: [
              {
                name: 'b',
                before: { present: false },
                after: { present: true, value: { val: '\u0000' } },
              },
            ],
          },
        ],
      },
    ];
    for (const plan of invalidPlans) {
      expect(trackTableCellPlan(editor, plan, scopeOnly(scope), CELL_CONTEXT).success).toBe(false);
      expect(snapshot()).toEqual(before);
    }
    for (const bad of [
      { ...scopeOnly(scope), rowId: neighbor.rowId, cellId: neighbor.cellId },
      { ...scopeOnly(scope), tableId: 'missing' },
    ]) {
      expect(trackTableCellPlan(editor, good, bad, CELL_CONTEXT).success).toBe(false);
      expect(snapshot()).toEqual(before);
    }
    editor.setRevisionMarkup({ trackFormatting: false });
    expect(trackTableCellPlan(editor, good, scopeOnly(scope), CELL_CONTEXT)).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
    expect(snapshot()).toEqual(before);
    const empty = createEditPlan(good.revision, []);
    if (!empty.success) throw new Error(empty.error);
    expect(trackTableCellPlan(editor, empty.plan, scopeOnly(scope), CELL_CONTEXT)).toMatchObject({
      success: true,
      proposals: [],
    });
  } finally {
    editor.destroy();
  }
});
