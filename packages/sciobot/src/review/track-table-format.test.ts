import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { CELL_CONTEXT, cellViews } from '../../test/support/table-review';
import { applyDirectPlan } from '../ai/apply-direct';
import type { NativeEditFailure } from '../contracts/document';
import type { CanonicalEditPlan } from '../diff/plan';
import { planTableFormat, type TableFormatTarget } from '../diff/table-format';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackTableFormatPlan } from './track-table-format';
import type { TrackTextContext } from './tracked-inline';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width: number, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${p(text)}</w:tc>`;
/** Fill, borders, a header flag, and a row without properties of its own. */
const document_ = () =>
  docxFixture(
    '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
      `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell('A', 2000, '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>')}${cell('B', 4000, '<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="FF0000"/></w:tcBorders>')}</w:tr>` +
      `<w:tr>${cell('One', 2000)}${cell('Two', 4000)}</w:tr>` +
      '</w:tbl>' +
      p('Tail')
  );
const mount = (bytes = document_(), review = true): DocxEditorInstance =>
  createDocxEditor({
    container: document.createElement('div'),
    document: bytes,
    author: 'User',
    modules: review ? [SCIOBOT_NATIVE_REVIEW_MODULE] : [],
  });
const children = (node: OoxmlNode | null | undefined) =>
  node && node.kind !== 'textValue' ? node.children : [];
const tableOf = (part: OoxmlPart) =>
  children(part.root.children[0]).find((n) => n.kind === 'table')!;
const rowsOf = (part: OoxmlPart) => children(tableOf(part)).filter((n) => n.kind === 'tableRow');
const cellIds = (part: OoxmlPart) =>
  rowsOf(part).map((row) =>
    children(row)
      .filter((n) => n.kind === 'tableCell')
      .map((n) => n.id)
  );
const gridIds = (part: OoxmlPart) =>
  children(
    children(tableOf(part)).find(
      (n) => n.kind !== 'textValue' && 'localName' in n && n.localName === 'tblGrid'
    )
  ).map((n) => n.id);

type Target = (part: OoxmlPart) => TableFormatTarget;
function formatPlan(editor: DocxEditorInstance, target: Target): CanonicalEditPlan {
  const part = editor.surface!.session.part();
  const result = planTableFormat(
    {
      revision: readEditorRevision(editor, CELL_CONTEXT.documentId, CELL_CONTEXT.epoch)!,
      part,
      tableId: tableOf(part).id,
    },
    target(part)
  );
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
function direct(plan: CanonicalEditPlan): OoxmlPart {
  const editor = mount(document_(), false);
  try {
    const result = applyDirectPlan(editor, plan, CELL_CONTEXT);
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    return editor.surface!.session.part();
  } finally {
    editor.destroy();
  }
}
const comparable = (part: OoxmlPart) =>
  serializeOoxmlPart(part).replace(/ w14:(paraId|textId)="[0-9A-F]+"/g, '');
/** Reopened ids carry no history, so a holder the proposal created stays as an empty element. */
const hollowFree = (part: OoxmlPart) => comparable(part).replace(/<w:(tcPr|trPr)\/>/g, '');

export const FORMAT_TARGETS: [string, Target][] = [
  [
    'a cell background',
    (part) => ({
      op: 'setTableCellFill',
      cellIds: [cellIds(part)[1]![1]!],
      color: { kind: 'hex', value: '00FF00' },
    }),
  ],
  [
    'a cleared background',
    (part) => ({ op: 'setTableCellFill', cellIds: [cellIds(part)[0]![0]!], color: null }),
  ],
  [
    'cell borders',
    (part) => ({
      op: 'setTableCellBorders',
      cellIds: [cellIds(part)[0]![1]!, cellIds(part)[1]![1]!],
      scope: 'all',
      spec: { style: 'double', size: 12, color: { kind: 'hex', value: '0000FF' } },
    }),
  ],
  [
    'vertical alignment',
    (part) => ({
      op: 'setTableCellVerticalAlignment',
      cellIds: cellIds(part).flat(),
      alignment: 'center',
    }),
  ],
  ['header rows', () => ({ op: 'setTableProperties', headerRowCount: 2 })],
  ['removed header rows', () => ({ op: 'setTableProperties', headerRowCount: 0 })],
  ['column widths', () => ({ op: 'setTableProperties', columnWidthsTwips: [3500, 2500] })],
  [
    'a divider drag',
    (part) => ({
      op: 'setTableColumnWidths',
      leftGridColumnId: gridIds(part)[0]!,
      rightGridColumnId: gridIds(part)[1]!,
      leftWidthTwips: 2600,
      rightWidthTwips: 3400,
    }),
  ],
  [
    'a row height',
    (part) => ({ op: 'setTableRowHeight', rowId: rowsOf(part)[1]!.id, heightTwips: 720 }),
  ],
];

for (const [name, target] of FORMAT_TARGETS)
  for (const action of ['accept', 'reject'] as const)
    test(`tracked ${name} ${action}s to the direct output or the exact original with one undo entry`, async () => {
      const editor = mount();
      try {
        const session = editor.surface!.session,
          originalPrint = canonicalOoxmlFingerprint(session.part());
        const before = await cellViews(editor);
        const plan = formatPlan(editor, target);
        const expected = direct(plan);
        const result = trackTableFormatPlan(editor, plan, CELL_CONTEXT);
        if (!result.success) throw new Error(`${result.code}: ${result.error}`);
        expect(result.proposals).toHaveLength(1);
        expect(
          result.proposals[0]!.addresses.every((entry) => ['ScioBot', ''].includes(entry.author))
        ).toBe(true);
        // Formatting proposals change no cell text in either view.
        expect(await cellViews(editor)).toEqual(before);
        expect(editor.getReviewItems().length).toBeGreaterThan(0);
        const pending = canonicalOoxmlFingerprint(session.part());
        expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        expect(editor.getReviewItems()).toHaveLength(0);
        if (action === 'accept') expect(comparable(session.part())).toBe(comparable(expected));
        else expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
        expect(await cellViews(editor)).toEqual([before[0], before[0]]);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
        expect(session.canUndo()).toBe(false);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
      } finally {
        editor.destroy();
      }
    }, 30_000);

test('a combined background and width proposal survives save/reopen with stable addresses', async () => {
  for (const action of ['accept', 'reject'] as const) {
    const editor = mount();
    try {
      const part = editor.surface!.session.part();
      const fill = formatPlan(editor, FORMAT_TARGETS[0]![1]);
      const widths = formatPlan(editor, FORMAT_TARGETS[6]![1]);
      const plan = {
        ...fill,
        operations: [...fill.operations, ...widths.operations],
      } as CanonicalEditPlan;
      const expected = direct(plan);
      const result = trackTableFormatPlan(editor, plan, CELL_CONTEXT);
      if (!result.success) throw new Error(result.error);
      expect(result.proposals).toHaveLength(1);
      const saved = new Uint8Array(await editor.save());
      const reopened = mount(saved);
      try {
        const ids = reopened
          .surface!.session.reviewItems()
          .flatMap((item) => (item.kind === 'revision' ? item.addresses : []))
          .map((entry) => entry.id)
          .sort();
        expect(ids).toEqual(result.proposals[0]!.addresses.map((entry) => entry.id).sort());
        expect(reopened.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
        expect(hollowFree(reopened.surface!.session.part())).toBe(
          hollowFree(action === 'accept' ? expected : part)
        );
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  }
}, 60_000);

test('a later proposal keeps the original record and reject still restores the source', () => {
  const editor = mount();
  try {
    const session = editor.surface!.session,
      originalPrint = canonicalOoxmlFingerprint(session.part());
    const first = trackTableFormatPlan(
      editor,
      formatPlan(editor, FORMAT_TARGETS[0]![1]),
      CELL_CONTEXT
    );
    expect(first.success).toBe(true);
    const second = trackTableFormatPlan(
      editor,
      formatPlan(editor, (part) => ({
        op: 'setTableCellFill',
        cellIds: [cellIds(part)[1]![1]!],
        color: { kind: 'hex', value: 'ABCDEF' },
      })),
      CELL_CONTEXT
    );
    if (!second.success) throw new Error(second.error);
    expect(serializeOoxmlPart(session.part())).toContain('w:fill="ABCDEF"');
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(originalPrint);
  } finally {
    editor.destroy();
  }
});

test('invalid, stale, structural, and formatting-disabled plans refuse atomically', () => {
  const editor = mount();
  try {
    const session = editor.surface!.session,
      revision = session.packageRevision(),
      print = canonicalOoxmlFingerprint(session.part());
    const refuse = (
      plan: CanonicalEditPlan,
      code: NativeEditFailure['code'],
      context: TrackTextContext = CELL_CONTEXT
    ) => {
      const result = trackTableFormatPlan(editor, plan, context);
      expect(result.success).toBe(false);
      if (!result.success) expect(result.code).toBe(code);
      expect(session.packageRevision()).toBe(revision);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(print);
    };
    const plan = formatPlan(editor, FORMAT_TARGETS[0]![1]);
    refuse(plan, 'invalid-input', { ...CELL_CONTEXT, author: '' });
    refuse(plan, 'invalid-input', { ...CELL_CONTEXT, date: 'yesterday' });
    refuse({ ...plan, revision: { ...plan.revision, revision: 999 } }, 'stale-revision');
    const table = tableOf(session.part()).id;
    refuse(
      {
        ...plan,
        operations: [
          {
            kind: 'table',
            partName: plan.operations[0]!.partName,
            operation: {
              op: 'deleteTableRow',
              tableId: table,
              rowId: rowsOf(session.part())[1]!.id,
            },
          },
        ],
      } as unknown as CanonicalEditPlan,
      'unsupported-operation'
    );
    editor.setRevisionMarkup({ trackFormatting: false });
    refuse(plan, 'unsupported-operation');
    editor.setRevisionMarkup({ trackFormatting: true });
    expect(trackTableFormatPlan(editor, plan, CELL_CONTEXT).success).toBe(true);
  } finally {
    editor.destroy();
  }
});

for (const mode of ['editing', 'suggesting'] as const)
  test(`table format proposals preserve ${mode} mode`, () => {
    const editor = mount();
    try {
      expect(editor.setEditingMode(mode).ok).toBe(true);
      const result = trackTableFormatPlan(
        editor,
        formatPlan(editor, FORMAT_TARGETS[2]![1]),
        CELL_CONTEXT
      );
      if (!result.success) throw new Error(result.error);
      expect(editor.surface!.editingMode()).toBe(mode === 'editing' ? 'edit' : 'suggest');
    } finally {
      editor.destroy();
    }
  });
