import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  serializeOoxmlPart,
  withPart,
  writeOoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { bodyPart, docxFixture } from '../../test/support/document';
import {
  namedChild,
  nodeChildren,
  rowCells,
  tableRows,
} from '../../test/support/native-table-oracle';
import {
  formatAttr as attr,
  formatBody as body,
  formatCell,
  formatCells as cells,
  formatContent as content,
  formatDescendants as descendants,
  formatNode as node,
  formatParagraph as paragraph,
  formatProperty as property,
  formatRow,
  formatTable,
  formatTables as tables,
} from '../../test/support/table-format-fixture';
import { nativeGridColumns } from './inserted-column-native';
import {
  decodeEditPlan,
  EDIT_PLAN_LIMITS,
  serializeEditPlan,
  type CanonicalEditPlan,
} from './plan';
import { planTableFormat, type TableFormatSource, type TableFormatTarget } from './table-format';

const revision = { documentId: 'table-properties', epoch: 'load-1', revision: 6 };
const source = (part = bodyPart(body()), tableId = tables(part)[0]!.id): TableFormatSource => ({
  revision,
  part,
  tableId,
});
function plan(input: TableFormatSource, target: TableFormatTarget): CanonicalEditPlan {
  const result = planTableFormat(input, target);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
/** Immutable native oracle. A01/R11 own guarded direct/review execution. */
function apply(part: OoxmlPart, planned: CanonicalEditPlan): OoxmlPart {
  for (const edit of planned.operations) {
    if (edit.kind !== 'table') throw new Error('Expected property operation');
    const result = applyTreeOp(part, edit.operation);
    if (!result.ok) throw new Error(result.reason);
    part = result.part;
  }
  return part;
}
const widths = (part: OoxmlPart, tableId: string) =>
  nativeGridColumns(part, tableId).map((entry) => attr(entry, 'w'));
const fill = (
  input: TableFormatSource,
  color: Extract<TableFormatTarget, { op: 'setTableCellFill' }>['color'],
  cellIds = [cells(input.part, input.tableId)[0]![0]!.id]
): TableFormatTarget => ({ op: 'setTableCellFill', cellIds, color });
const spec = { style: 'double', size: 12, color: { kind: 'hex', value: 'ABCDEF' } } as const;

test('one cell background changes without text, marks, neighboring borders, or other tables', () => {
  const input = source(),
    old = cells(input.part, input.tableId),
    selected = old[0]![0]!,
    before = canonicalOoxmlFingerprint(input.part);
  const target = fill(input, { kind: 'hex', value: 'FFEEDD' }),
    planned = plan(input, target);
  expect(planned.operations).toEqual([
    { kind: 'table', partName: input.part.name, operation: { ...target, tableId: input.tableId } },
  ]);
  const result = apply(input.part, planned),
    updated = node(result, selected.id),
    shd = property(updated, 'tcPr', 'shd');
  expect(attr(shd, 'fill')).toBe('FFEEDD');
  expect(attr(shd, 'color')).toBe('112233');
  expect(attr(shd, 'val')).toBe('pct10');
  expect(attr(shd, 'value', 'urn:keep')).toBe('shading');
  expect(property(updated, 'tcPr', 'tcBorders') === property(selected, 'tcPr', 'tcBorders')).toBe(
    true
  );
  expect(content(result, input.tableId)).toEqual(content(input.part, input.tableId));
  for (const p of descendants(selected).filter((entry) => entry.kind === 'paragraph'))
    expect(findNode(result, p.id) === p).toBe(true);
  for (const cell of old.flat().filter((entry) => entry !== selected))
    expect(findNode(result, cell.id) === cell).toBe(true);
  expect(findNode(result, tables(input.part)[1]!.id) === tables(input.part)[1]).toBe(true);
  expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
});

test('clear, theme, and auto backgrounds preserve authored shading patterns and unrelated attributes', () => {
  const input = source(),
    selected = cells(input.part, input.tableId)[0]![0]!;
  for (const color of [
    null,
    { kind: 'theme', slot: 'accent2', resolvedHex: 'AABBCC', tint: 0.4 },
    { kind: 'auto', resolvedHex: '123456' },
  ] as const) {
    const result = apply(input.part, plan(input, fill(input, color))),
      shd = property(node(result, selected.id), 'tcPr', 'shd');
    expect(attr(shd, 'fill')).toBe(
      color === null ? undefined : color.kind === 'theme' ? 'AABBCC' : '123456'
    );
    expect(attr(shd, 'themeFill')).toBe(color?.kind === 'theme' ? 'accent2' : undefined);
    expect(attr(shd, 'themeFillTint')).toBe(color?.kind === 'theme' ? '66' : undefined);
    expect(attr(shd, 'color')).toBe('112233');
    expect(attr(shd, 'val')).toBe('pct10');
    expect(attr(shd, 'value', 'urn:keep')).toBe('shading');
    expect(content(result, input.tableId)).toEqual(content(input.part, input.tableId));
  }
});

test('border changes preserve text, background, unaffected sides, and imported side metadata', () => {
  const input = source(),
    selected = cells(input.part, input.tableId)[0]![0]!,
    before = property(selected, 'tcPr', 'tcBorders');
  const result = apply(
    input.part,
    plan(input, { op: 'setTableCellBorders', cellIds: [selected.id], scope: 'top', spec })
  );
  const after = property(node(result, selected.id), 'tcPr', 'tcBorders'),
    top = namedChild(after, 'top')!;
  expect(attr(top, 'val')).toBe('double');
  expect(attr(top, 'sz')).toBe('12');
  expect(attr(top, 'color')).toBe('ABCDEF');
  expect(top.id).toBe(namedChild(before, 'top')!.id);
  expect(attr(top, 'value', 'urn:keep')).toBe('top');
  for (const side of ['left', 'bottom', 'right'])
    expect(namedChild(after, side) === namedChild(before, side)).toBe(true);
  expect(
    property(node(result, selected.id), 'tcPr', 'shd') === property(selected, 'tcPr', 'shd')
  ).toBe(true);
  expect(content(result, input.tableId)).toEqual(content(input.part, input.tableId));
  const cleared = apply(
    result,
    plan(
      { ...input, part: result },
      { op: 'setTableCellBorders', cellIds: [selected.id], scope: 'none', target: 'top' }
    )
  );
  const clearTop = namedChild(property(node(cleared, selected.id), 'tcPr', 'tcBorders'), 'top')!;
  expect(attr(clearTop, 'val')).toBe('none');
  expect(attr(clearTop, 'value', 'urn:keep')).toBe('top');
});

test('inside border selection writes symmetric facing edges and retains its outer hull', () => {
  const input = source(),
    [left, right] = cells(input.part, input.tableId)[0]!;
  const result = apply(
    input.part,
    plan(input, {
      op: 'setTableCellBorders',
      cellIds: [left!.id, right!.id],
      scope: 'inside',
      spec,
    })
  );
  for (const [cell, side, unchanged] of [
    [left!, 'right', 'left'],
    [right!, 'left', 'right'],
  ] as const) {
    const borders = property(node(result, cell.id), 'tcPr', 'tcBorders');
    expect(attr(namedChild(borders, side)!, 'val')).toBe('double');
    expect(attr(namedChild(borders, side)!, 'sz')).toBe('12');
    expect(
      namedChild(borders, unchanged) === namedChild(property(cell, 'tcPr', 'tcBorders'), unchanged)
    ).toBe(true);
  }
  expect(content(result, input.tableId)).toEqual(content(input.part, input.tableId));
});

test('table style and repeated header flags change independently from row heights and cell content', () => {
  const input = source(),
    rows = tableRows(input.part, input.tableId);
  const result = apply(
    input.part,
    plan(input, { op: 'setTableProperties', styleId: 'NewStyle', headerRowCount: 2 })
  );
  expect(attr(property(node(result, input.tableId), 'tblPr', 'tblStyle'), 'val')).toBe('NewStyle');
  expect(
    tableRows(result, input.tableId).map((row) =>
      Boolean(namedChild(namedChild(row, 'trPr')!, 'tblHeader'))
    )
  ).toEqual([true, true, false]);
  expect(
    property(node(result, input.tableId), 'tblPr', 'tblCaption') ===
      property(node(input.part, input.tableId), 'tblPr', 'tblCaption')
  ).toBe(true);
  for (const row of rows) {
    expect(
      property(node(result, row.id), 'trPr', 'trHeight') === property(row, 'trPr', 'trHeight')
    ).toBe(true);
    for (const cell of rowCells(row)) expect(findNode(result, cell.id) === cell).toBe(true);
  }
  const cleared = apply(
    result,
    plan({ ...input, part: result }, { op: 'setTableProperties', headerRowCount: 0 })
  );
  expect(
    tableRows(cleared, input.tableId).every(
      (row) => !namedChild(namedChild(row, 'trPr')!, 'tblHeader')
    )
  ).toBe(true);
  expect(content(cleared, input.tableId)).toEqual(content(input.part, input.tableId));
});

test.each(['full-grid', 'divider', 'right-edge'] as const)(
  '%s width changes retain unequal grid merges, XML, and content',
  (kind) => {
    const input = source(bodyPart(body(true))),
      old = cells(input.part, input.tableId),
      grid = nativeGridColumns(input.part, input.tableId);
    const target: TableFormatTarget =
      kind === 'full-grid'
        ? { op: 'setTableProperties', columnWidthsTwips: [1500, 2700, 3600] }
        : kind === 'divider'
          ? {
              op: 'setTableColumnWidths',
              leftGridColumnId: grid[0]!.id,
              rightGridColumnId: grid[1]!.id,
              leftWidthTwips: 1800,
              rightWidthTwips: 1800,
            }
          : {
              op: 'setTableRightEdgeWidth',
              gridColumnId: grid[2]!.id,
              columnWidthTwips: 4200,
              tableWidthTwips: 7800,
            };
    const result = apply(input.part, plan(input, target));
    expect(widths(result, input.tableId)).toEqual(
      kind === 'full-grid'
        ? ['1500', '2700', '3600']
        : kind === 'divider'
          ? ['1800', '1800', '3600']
          : ['1200', '2400', '4200']
    );
    for (const cell of old.flat()) {
      const updated = node(result, cell.id);
      for (const key of ['gridSpan', 'vMerge', 'shd', 'tcBorders'])
        expect(
          namedChild(namedChild(updated, 'tcPr')!, key) ===
            namedChild(namedChild(cell, 'tcPr')!, key)
        ).toBe(true);
      const oldWidth = property(cell, 'tcPr', 'tcW'),
        nextWidth = property(updated, 'tcPr', 'tcW');
      expect(nextWidth.id).toBe(oldWidth.id);
      expect(attr(nextWidth, 'value', 'urn:keep')).toBe('cell-width');
      expect(nodeChildren(nextWidth)).toEqual(nodeChildren(oldWidth));
    }
    expect(attr(property(node(result, old[0]![0]!.id), 'tcPr', 'tcW'), 'w')).toBe(
      kind === 'full-grid' ? '4200' : '3600'
    );
    for (const key of ['tblW', 'tblLayout']) {
      const next = property(node(result, input.tableId), 'tblPr', key),
        old = property(node(input.part, input.tableId), 'tblPr', key);
      expect(next.id).toBe(old.id);
      expect(nodeChildren(next)).toEqual(nodeChildren(old));
      expect(attr(next, 'value', 'urn:keep')).toBe(attr(old, 'value', 'urn:keep'));
    }
    expect(content(result, input.tableId)).toEqual(content(input.part, input.tableId));
    expect(findNode(result, tables(input.part)[1]!.id) === tables(input.part)[1]).toBe(true);
  }
);

test('exact row height retains imported metadata, row siblings, cell merges, and content', () => {
  const input = source(),
    row = tableRows(input.part, input.tableId)[0]!,
    height = property(row, 'trPr', 'trHeight');
  const result = apply(
      input.part,
      plan(input, { op: 'setTableRowHeight', rowId: row.id, heightTwips: 600 })
    ),
    next = property(node(result, row.id), 'trPr', 'trHeight');
  expect(next.id).toBe(height.id);
  expect(attr(next, 'val')).toBe('600');
  expect(attr(next, 'hRule')).toBe('exact');
  expect(attr(next, 'value', 'urn:keep')).toBe('height');
  expect(nodeChildren(next)).toEqual(nodeChildren(height));
  for (const cell of rowCells(row)) expect(findNode(result, cell.id) === cell).toBe(true);
  expect(content(result, input.tableId)).toEqual(content(input.part, input.tableId));
  expect(
    plan(input, { op: 'setTableRowHeight', rowId: row.id, heightTwips: 400 }).operations
  ).toHaveLength(1);
  expect(
    plan({ ...input, part: result }, { op: 'setTableRowHeight', rowId: row.id, heightTwips: 600 })
      .operations
  ).toEqual([]);
});

test('vertical alignment and background require the complete physical vertical-merge selection', () => {
  const input = source(bodyPart(body(true))),
    old = cells(input.part, input.tableId),
    ids = [old[0]![0]!.id, old[1]![0]!.id];
  expect(
    planTableFormat(input, fill(input, { kind: 'hex', value: 'FFEEDD' }, [ids[0]!])).success
  ).toBe(false);
  const aligned = apply(
    input.part,
    plan(input, { op: 'setTableCellVerticalAlignment', cellIds: ids, alignment: 'center' })
  );
  const result = apply(
    aligned,
    plan({ ...input, part: aligned }, fill(input, { kind: 'hex', value: 'FFEEDD' }, ids))
  );
  for (const id of ids) {
    expect(attr(property(node(result, id), 'tcPr', 'vAlign'), 'val')).toBe('center');
    expect(attr(property(node(result, id), 'tcPr', 'vAlign'), 'value', 'urn:keep')).toBe(
      'alignment'
    );
    expect(attr(property(node(result, id), 'tcPr', 'shd'), 'fill')).toBe('FFEEDD');
  }
  expect(content(result, input.tableId)).toEqual(content(input.part, input.tableId));
  expect(widths(result, input.tableId)).toEqual(['1200', '2400', '3600']);
});

test('authored no-ops emit no operations while inherited or exactness changes stay explicit', () => {
  const input = source(),
    selected = cells(input.part, input.tableId)[0]![0]!;
  for (const target of [
    fill(input, { kind: 'hex', value: 'EEEEEE' }),
    { op: 'setTableCellVerticalAlignment', cellIds: [selected.id], alignment: 'top' },
    {
      op: 'setTableProperties',
      styleId: 'Original',
      headerRowCount: 1,
      columnWidthsTwips: [1200, 2400, 3600],
    },
  ] as const)
    expect(plan(input, target).operations).toEqual([]);
  const clean = source(
    bodyPart(
      formatTable(
        '<w:tr>' +
          formatCell('A', 1200).replace(/<w:shd[^>]*\/>/, '') +
          formatCell('B') +
          formatCell('C', 3600) +
          '</w:tr>'
      )
    )
  );
  expect(plan(clean, fill(clean, null)).operations).toEqual([]);
  expect(plan(clean, fill(clean, { kind: 'hex', value: 'EEEEEE' })).operations).toHaveLength(1);
});

test('unchanged native divider and right-edge drag proposals retain native refusal semantics', () => {
  const input = source(),
    grid = nativeGridColumns(input.part, input.tableId),
    before = canonicalOoxmlFingerprint(input.part);
  for (const target of [
    {
      op: 'setTableColumnWidths',
      leftGridColumnId: grid[0]!.id,
      rightGridColumnId: grid[1]!.id,
      leftWidthTwips: 1200,
      rightWidthTwips: 2400,
    },
    {
      op: 'setTableRightEdgeWidth',
      gridColumnId: grid[2]!.id,
      columnWidthTwips: 3600,
      tableWidthTwips: 7200,
    },
  ] as const) {
    const result = planTableFormat(input, target);
    expect(result.success).toBe(false);
    if (!result.success) expect(result.code).toBe('unsupported-operation');
  }
  // Full-grid target assignment remains idempotent for AI target composition.
  expect(
    plan(input, { op: 'setTableProperties', columnWidthsTwips: [1200, 2400, 3600] }).operations
  ).toEqual([]);
  expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
});

test('detached frozen plans contain only original selectors and native property values', () => {
  const input = source(),
    target = fill(input, { kind: 'hex', value: 'FFEEDD' }),
    before = JSON.stringify(target),
    planned = plan(input, target);
  expect(JSON.stringify(target)).toBe(before);
  expect(Object.isFrozen(target)).toBe(false);
  expect(Object.isFrozen(planned.operations)).toBe(true);
  const encoded = serializeEditPlan(planned);
  expect(decodeEditPlan(encoded)).toEqual({ success: true, plan: planned });
  expect(encoded).not.toMatch(/"children"|created|textValue|previewPart/);
  if (target.op === 'setTableCellFill') {
    (target.cellIds as string[]).push('later');
    if (target.color?.kind === 'hex') (target.color as { value: string }).value = '000000';
  }
  expect(serializeEditPlan(planned)).toBe(encoded);
});

test('property plans survive DOCX save/reopen and retain all unrelated package parts', () => {
  const loaded = readOoxmlPackage(docxFixture(body()));
  if (!loaded.ok) throw new Error(loaded.reason);
  const pkg = loaded.package,
    input = source(pkg.parts.get(pkg.mainDocumentPart)!),
    selected = cells(input.part, input.tableId)[0]![0]!,
    grid = nativeGridColumns(input.part, input.tableId);
  let part = input.part;
  for (const target of [
    fill(input, { kind: 'hex', value: 'FFEEDD' }),
    {
      op: 'setTableColumnWidths',
      leftGridColumnId: grid[0]!.id,
      rightGridColumnId: grid[1]!.id,
      leftWidthTwips: 1800,
      rightWidthTwips: 1800,
    },
    { op: 'setTableCellBorders', cellIds: [selected.id], scope: 'top', spec },
    { op: 'setTableProperties', headerRowCount: 2 },
  ] as const)
    part = apply(part, plan({ ...input, part }, target));
  const reopened = readOoxmlPackage(writeOoxmlPackage(withPart(pkg, part)));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(canonicalOoxmlFingerprint(reopened.package.parts.get(pkg.mainDocumentPart)!)).toBe(
    canonicalOoxmlFingerprint(part)
  );
  for (const [name, original] of pkg.parts)
    if (name !== pkg.mainDocumentPart)
      expect(canonicalOoxmlFingerprint(reopened.package.parts.get(name)!)).toBe(
        canonicalOoxmlFingerprint(original)
      );
  expect(serializeOoxmlPart(part)).toContain('keep:value="cell-width"');
});

test('invalid selectors, values, units, bounds, and property commands refuse without source writes', () => {
  const input = source(),
    own = cells(input.part, input.tableId),
    ids = own[0]!.map((entry) => entry.id),
    other = tableRows(input.part, tables(input.part)[1]!.id)[0]!,
    grid = nativeGridColumns(input.part, input.tableId),
    before = canonicalOoxmlFingerprint(input.part);
  const invalid: unknown[] = [
    { op: 'insertTableRow' },
    { ...fill(input, null), tableId: input.tableId },
    { op: 'setTableCellFill', cellIds: [], color: null },
    fill(input, null, [ids[0]!, ids[0]!]),
    fill(input, null, ['missing']),
    fill(input, null, [rowCells(other)[0]!.id]),
    fill(input, null, [ids[0]!, own[1]![1]!.id]),
    fill(input, { kind: 'hex', value: 'javascript:alert(1)' }),
    {
      op: 'setTableCellFill',
      cellIds: ids,
      color: { kind: 'theme', slot: 'invalid', resolvedHex: 'FFFFFF' },
    },
    { op: 'setTableCellVerticalAlignment', cellIds: ids, alignment: 'middle' },
    { op: 'setTableCellBorders', cellIds: ids, scope: 'top', spec: { ...spec, size: -1 } },
    { op: 'setTableCellBorders', cellIds: ids, scope: 'none' },
    { op: 'setTableProperties', columnWidthsTwips: [1500, 2700] },
    { op: 'setTableProperties', columnWidthsTwips: [100, 2400, 3600] },
    { op: 'setTableProperties', columnWidthsTwips: [20000, 20000, 20000] },
    { op: 'setTableProperties', headerRowCount: 4 },
    { op: 'setTableProperties', headerRowCount: 0.5 },
    { op: 'setTableRowHeight', rowId: other.id, heightTwips: 600 },
    {
      op: 'setTableRowHeight',
      rowId: tableRows(input.part, input.tableId)[0]!.id,
      heightTwips: 19,
    },
    {
      op: 'setTableColumnWidths',
      leftGridColumnId: grid[0]!.id,
      rightGridColumnId: grid[2]!.id,
      leftWidthTwips: 1800,
      rightWidthTwips: 3000,
    },
    {
      op: 'setTableRightEdgeWidth',
      gridColumnId: grid[2]!.id,
      columnWidthTwips: 4200,
      tableWidthTwips: 9999,
    },
  ];
  for (const target of invalid)
    expect(planTableFormat(input, target as TableFormatTarget).success).toBe(false);
  expect(planTableFormat({ ...input, tableId: 'missing' }, fill(input, null)).success).toBe(false);
  expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
});

test('getters, cycles, sparse lists, hidden fields, and resource excess refuse before evaluation', () => {
  const input = source(),
    target = fill(input, null);
  let calls = 0;
  const getter = Object.defineProperty({ ...target }, 'color', {
    enumerable: true,
    get() {
      calls++;
      return null;
    },
  });
  const hidden = Object.defineProperty({ ...target }, 'hidden', { value: true }),
    cycle: Record<string, unknown> = { ...target };
  cycle.color = cycle;
  const sparse = { ...target, cellIds: new Array(2) },
    deep: Record<string, unknown> = { ...target };
  let current = deep;
  for (let i = 0; i < EDIT_PLAN_LIMITS.maxDepth + 1; i++) {
    const next = {};
    current.next = next;
    current = next;
  }
  for (const value of [
    getter,
    hidden,
    cycle,
    sparse,
    deep,
    { op: 'setTableProperties', styleId: 'x'.repeat(EDIT_PLAN_LIMITS.maxBytes + 1) },
  ])
    expect(planTableFormat(input, value as TableFormatTarget).success).toBe(false);
  for (const key of ['revision', 'part', 'tableId'])
    expect(
      planTableFormat(
        Object.defineProperty({ ...input }, key, {
          enumerable: true,
          get() {
            calls++;
            return input[key as keyof TableFormatSource];
          },
        }),
        target
      ).success
    ).toBe(false);
  expect(calls).toBe(0);
});

test('nested tables isolate property ownership and default native locks refuse before planning', () => {
  const nested = formatTable(formatRow('Nested A', 'Nested B', 'Nested C'));
  const input = source(
      bodyPart(
        formatTable(
          '<w:tr><w:tc>' +
            paragraph('Outer') +
            nested +
            '</w:tc>' +
            formatCell('B') +
            formatCell('C') +
            '</w:tr>'
        )
      )
    ),
    outer = cells(input.part, input.tableId)[0]![0]!,
    inner = nodeChildren(outer).find((entry) => entry.kind === 'table')!,
    inside = source(input.part, inner.id),
    selected = cells(inside.part, inside.tableId)[0]![0]!;
  expect(planTableFormat(input, fill(inside, null, [selected.id])).success).toBe(false);
  const result = apply(input.part, plan(inside, fill(inside, { kind: 'hex', value: 'FFEEDD' })));
  expect(attr(property(node(result, selected.id), 'tcPr', 'shd'), 'fill')).toBe('FFEEDD');
  expect(content(result, input.tableId)).toEqual(content(input.part, input.tableId));
  for (const restriction of [
    '<w:lock w:val="sdtContentLocked"/>',
    '<w:dataBinding w:storeItemID="{00000000-0000-0000-0000-000000000001}" w:xpath="/item"/>',
  ]) {
    const part = bodyPart(
        '<w:sdt><w:sdtPr>' +
          restriction +
          '</w:sdtPr><w:sdtContent>' +
          formatTable(formatRow('A', 'B', 'C')) +
          '</w:sdtContent></w:sdt>'
      ),
      table = descendants(part.root).find((entry) => entry.kind === 'table')!,
      locked = source(part, table.id);
    expect(planTableFormat(locked, fill(locked, null)).success).toBe(false);
  }
});

test('planning leaves live bytes, revision, selection, and existing undo history unchanged', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(body()),
  });
  try {
    const surface = editor.surface!,
      session = surface.session;
    surface.applyAutomationOps(() => [
      { op: 'insertText', paragraphId: session.paragraphIds()[0]!, offset: 0, text: 'User ' },
    ]);
    const current = session.packageRevision(),
      input = { ...source(session.part()), revision: { ...revision, revision: current } },
      bytes = writeOoxmlPackage(session.currentPackage()),
      selection = surface.state().selection,
      history = [session.canUndo(), session.canRedo()];
    expect(plan(input, fill(input, { kind: 'hex', value: 'FFEEDD' })).operations).toHaveLength(1);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(session.packageRevision()).toBe(current);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), session.paragraphIds()[0]!)).toBe('Before');
  } finally {
    editor.destroy();
  }
});
