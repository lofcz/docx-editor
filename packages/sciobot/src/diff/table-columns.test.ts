import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  serializeOoxmlPart,
  withPart,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { bodyPart, docxFixture } from '../../test/support/document';
import { decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';
import { planTableColumn, type TableColumnSource, type TableColumnTarget } from './table-columns';

const revision = { documentId: 'columns', epoch: 'load-1', revision: 6 };
const p = (text: string) =>
  `<w:p><w:pPr><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width = 2400, props = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:shd w:fill="EEEEEE"/>${props}</w:tcPr>${p(text)}</w:tc>`;
const row = (a: string, b: string, c: string) =>
  `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(a, 1200)}${cell(b, 2400, '<keep:extra xmlns:keep="urn:keep" keep:value="retain"/>')}${cell(c, 3600)}</w:tr>`;
const grid = (widths: readonly number[] = [1200, 2400, 3600]) =>
  `<w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>`;
const table = (rows: string, columns = grid()) =>
  `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="7200" w:type="dxa"/><w:tblCaption w:val="Keep table metadata"/></w:tblPr>${columns}${rows}</w:tbl>`;
const body = () =>
  p('Before') +
  table(row('A1', 'B1', 'C1') + row('A2', 'B2', 'C2')) +
  p('Between') +
  table(row('Other A', 'Other B', 'Other C')) +
  p('After');
const children = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;
const child = (node: OoxmlNode, name: string) =>
  children(node).find((entry) => entry.kind !== 'textValue' && entry.localName === name);
const tables = (part: OoxmlPart) =>
  children(bodyStoryRoot(part)!).filter((entry) => entry.kind === 'table');
const rows = (part: OoxmlPart, tableId: string) =>
  children(findNode(part, tableId)!).filter((entry) => entry.kind === 'tableRow');
const cells = (row: OoxmlNode) => children(row).filter((entry) => entry.kind === 'tableCell');
const columns = (part: OoxmlPart, tableId: string) =>
  children(child(findNode(part, tableId)!, 'tblGrid')!).filter(
    (entry) => entry.kind !== 'textValue' && entry.localName === 'gridCol'
  );
const attr = (node: OoxmlNode, name: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find((entry) => entry.localName === name)?.value;
function source(part = bodyPart(body()), tableIndex = 0): TableColumnSource {
  return { revision, part, tableId: tables(part)[tableIndex]!.id };
}
function text(part: OoxmlPart, node: OoxmlNode): string {
  return node.kind === 'paragraph'
    ? paragraphTextOf(part, node.id)!
    : children(node)
        .map((node) => text(part, node))
        .join('');
}
const values = (part: OoxmlPart, id: string) =>
  rows(part, id).map((row) => cells(row).map((cell) => text(part, cell)));
const widths = (part: OoxmlPart, id: string) =>
  columns(part, id).map((node) => Number(attr(node, 'w')));
function plan(input: TableColumnSource, target: TableColumnTarget): CanonicalEditPlan {
  const result = planTableColumn(input, target);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
/** Immutable native oracle. A01/R09 own live execution, review, and collaboration. */
function apply(part: OoxmlPart, planned: CanonicalEditPlan): OoxmlPart {
  let current = part;
  for (const edit of planned.operations) {
    if (edit.kind !== 'table') throw new Error('Expected native table intent');
    const result = applyTreeOp(current, edit.operation);
    if (!result.ok) throw new Error(result.reason);
    current = result.part;
  }
  return current;
}

test.each(['left', 'right'] as const)(
  'insert a middle column on the %s of its anchor with exact inherited widths and stable surviving content',
  (where) => {
    const input = source(),
      originalRows = rows(input.part, input.tableId),
      oldCells = originalRows.flatMap(cells);
    const oldColumns = columns(input.part, input.tableId),
      reference = oldColumns[where === 'left' ? 1 : 0]!;
    const untouched = tables(input.part)[1]!,
      tableBefore = findNode(input.part, input.tableId)!;
    const before = canonicalOoxmlFingerprint(input.part);
    const planned = plan(input, { kind: 'insert', gridColumnId: reference.id, where });
    expect(planned.operations).toEqual([
      {
        kind: 'table',
        partName: input.part.name,
        operation: {
          op: 'insertTableColumn',
          tableId: input.tableId,
          gridColumnId: reference.id,
          where,
        },
      },
    ]);
    const result = apply(input.part, planned),
      expectedWidth = where === 'left' ? 2400 : 1200;
    expect(values(result, input.tableId)).toEqual([
      ['A1', '', 'B1', 'C1'],
      ['A2', '', 'B2', 'C2'],
    ]);
    expect(widths(result, input.tableId)).toEqual([1200, expectedWidth, 2400, 3600]);
    expect(widths(result, input.tableId).reduce((sum, value) => sum + value, 0)).toBe(
      7200 + expectedWidth
    );
    for (const node of [...oldCells, ...oldColumns, untouched])
      expect(findNode(result, node.id) === node).toBe(true);
    const afterTable = findNode(result, input.tableId)!;
    expect(child(afterTable, 'tblPr') === child(tableBefore, 'tblPr')).toBe(true);
    for (let index = 0; index < originalRows.length; index++) {
      const newRow = rows(result, input.tableId)[index]!;
      expect(newRow.id).toBe(originalRows[index]!.id);
      expect(child(newRow, 'trPr') === child(originalRows[index]!, 'trPr')).toBe(true);
      const newCell = cells(newRow)[1]!,
        paragraph = children(newCell).find((node) => node.kind === 'paragraph')!;
      expect(child(child(newCell, 'tcPr')!, 'shd')).toBeDefined();
      expect(child(child(newCell, 'tcPr')!, 'extra')).toBeUndefined();
      expect(child(child(paragraph, 'pPr')!, 'rPr')).toBeDefined();
    }
    expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
  }
);

test('delete the original middle column, preserving surviving cell IDs, properties, and unrelated tables', () => {
  const input = source(),
    oldRows = rows(input.part, input.tableId),
    oldColumns = columns(input.part, input.tableId);
  const untouched = tables(input.part)[1]!,
    planned = plan(input, { kind: 'delete', gridColumnId: oldColumns[1]!.id });
  const result = apply(input.part, planned);
  expect(values(result, input.tableId)).toEqual([
    ['A1', 'C1'],
    ['A2', 'C2'],
  ]);
  expect(widths(result, input.tableId)).toEqual([1200, 3600]);
  expect(widths(result, input.tableId).reduce((sum, value) => sum + value, 0)).toBe(4800);
  expect(findNode(result, oldColumns[1]!.id)).toBeNull();
  for (const row of oldRows) {
    const beforeCells = cells(row),
      survivors = cells(findNode(result, row.id)!);
    expect(survivors[0] === beforeCells[0]).toBe(true);
    expect(survivors[1] === beforeCells[2]).toBe(true);
    expect(findNode(result, beforeCells[1]!.id)).toBeNull();
  }
  expect(findNode(result, untouched.id) === untouched).toBe(true);
  expect(serializeEditPlan(planned)).not.toContain(untouched.id);
});

test('insert then delete the newly allocated middle column restores the original table fingerprint', () => {
  const input = source(),
    old = new Set(columns(input.part, input.tableId).map((node) => node.id));
  const first = apply(
    input.part,
    plan(input, {
      kind: 'insert',
      where: 'left',
      gridColumnId: columns(input.part, input.tableId)[1]!.id,
    })
  );
  const inserted = columns(first, input.tableId).find((node) => !old.has(node.id))!;
  const restored = apply(
    first,
    plan({ ...input, part: first }, { kind: 'delete', gridColumnId: inserted.id })
  );
  expect(canonicalOoxmlFingerprint(findNode(restored, input.tableId)!)).toBe(
    canonicalOoxmlFingerprint(findNode(input.part, input.tableId)!)
  );
});

function mergedSource(): TableColumnSource {
  return source(
    bodyPart(
      table(
        `<w:tr>${cell('AB', 3600, '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>')}${cell('C1', 3600)}</w:tr>` +
          `<w:tr>${cell('', 3600, '<w:gridSpan w:val="2"/><w:vMerge/>')}${cell('C2', 3600)}</w:tr>` +
          row('A3', 'B3', 'C3')
      ) +
        p('Gap') +
        table(row('Other A', 'Other B', 'Other C'))
    )
  );
}

test('the public native dispatcher widens spanning cells while creating cells only in other rows', () => {
  const input = mergedSource(),
    originalRows = rows(input.part, input.tableId),
    originalCells = originalRows.map(cells);
  const result = apply(
    input.part,
    plan(input, {
      kind: 'insert',
      where: 'left',
      gridColumnId: columns(input.part, input.tableId)[1]!.id,
    })
  );
  const afterRows = rows(result, input.tableId),
    afterCells = afterRows.map(cells);
  expect(values(result, input.tableId)).toEqual([
    ['AB', 'C1'],
    ['', 'C2'],
    ['A3', '', 'B3', 'C3'],
  ]);
  expect(widths(result, input.tableId)).toEqual([1200, 2400, 2400, 3600]);
  for (let index = 0; index < 2; index++) {
    expect(afterCells[index]).toHaveLength(2);
    const beforeCell = originalCells[index]![0]!,
      afterCell = afterCells[index]![0]!;
    expect(afterCell.id).toBe(beforeCell.id);
    expect(attr(child(child(afterCell, 'tcPr')!, 'gridSpan')!, 'val')).toBe('3');
    expect(attr(child(child(afterCell, 'tcPr')!, 'tcW')!, 'w')).toBe('6000');
    expect(child(child(afterCell, 'tcPr')!, 'vMerge')!.id).toBe(
      child(child(beforeCell, 'tcPr')!, 'vMerge')!.id
    );
    expect(
      children(afterCell).find((node) => node.kind === 'paragraph') ===
        children(beforeCell).find((node) => node.kind === 'paragraph')
    ).toBe(true);
    expect(afterCells[index]![1] === originalCells[index]![1]).toBe(true);
  }
  for (const cell of originalCells[2]!) expect(findNode(result, cell.id) === cell).toBe(true);
  expect(findNode(result, tables(input.part)[1]!.id) === tables(input.part)[1]).toBe(true);
});

test('deleting inside horizontal and vertical spans narrows them and retains all surviving content', () => {
  const input = mergedSource(),
    originalRows = rows(input.part, input.tableId);
  const result = apply(
    input.part,
    plan(input, { kind: 'delete', gridColumnId: columns(input.part, input.tableId)[1]!.id })
  );
  expect(values(result, input.tableId)).toEqual([
    ['AB', 'C1'],
    ['', 'C2'],
    ['A3', 'C3'],
  ]);
  expect(widths(result, input.tableId)).toEqual([1200, 3600]);
  for (let index = 0; index < 2; index++) {
    const beforeCell = cells(originalRows[index]!)[0]!,
      afterCell = cells(rows(result, input.tableId)[index]!)[0]!;
    expect(afterCell.id).toBe(beforeCell.id);
    expect(child(child(afterCell, 'tcPr')!, 'gridSpan')).toBeUndefined();
    expect(attr(child(child(afterCell, 'tcPr')!, 'tcW')!, 'w')).toBe('1200');
    expect(child(child(afterCell, 'tcPr')!, 'vMerge')).toBeDefined();
    expect(
      children(afterCell).find((node) => node.kind === 'paragraph') ===
        children(beforeCell).find((node) => node.kind === 'paragraph')
    ).toBe(true);
  }
});

test('insertion beside a merged column copies its native vertical continuation to new cells', () => {
  const input = mergedSource();
  const result = apply(
    input.part,
    plan(input, {
      kind: 'insert',
      where: 'right',
      gridColumnId: columns(input.part, input.tableId)[1]!.id,
    })
  );
  const after = rows(result, input.tableId).map(cells);
  expect(values(result, input.tableId)).toEqual([
    ['AB', '', 'C1'],
    ['', '', 'C2'],
    ['A3', 'B3', '', 'C3'],
  ]);
  expect(after[0]).toHaveLength(3);
  expect(after[1]).toHaveLength(3);
  expect(attr(child(child(after[0]![1]!, 'tcPr')!, 'vMerge')!, 'val')).toBe('restart');
  expect(child(child(after[1]![1]!, 'tcPr')!, 'vMerge')).toBeDefined();
  // The native writer uses the OOXML default continuation marker, with no explicit value.
  expect(attr(child(child(after[1]![1]!, 'tcPr')!, 'vMerge')!, 'val')).toBeUndefined();
  expect(widths(result, input.tableId)).toEqual([1200, 2400, 2400, 3600]);
});

test('a native table without a grid uses a first-row reference cell and synthesizes only the required grid', () => {
  const input = source(bodyPart(table(row('A1', 'B1', 'C1') + row('A2', 'B2', 'C2'), '')));
  const oldRows = rows(input.part, input.tableId),
    oldCells = oldRows.flatMap(cells),
    reference = cells(oldRows[0]!)[1]!;
  const planned = plan(input, { kind: 'insert', where: 'left', referenceCellId: reference.id });
  const result = apply(input.part, planned);
  expect(values(result, input.tableId)).toEqual([
    ['A1', '', 'B1', 'C1'],
    ['A2', '', 'B2', 'C2'],
  ]);
  expect(widths(result, input.tableId)).toEqual([1200, 2400, 2400, 3600]);
  for (const cell of oldCells) expect(findNode(result, cell.id) === cell).toBe(true);
  expect(
    planTableColumn(input, {
      kind: 'insert',
      where: 'right',
      referenceCellId: cells(oldRows[1]!)[1]!.id,
    }).success
  ).toBe(false);
  expect(serializeEditPlan(planned)).not.toMatch(/gridColumnId|created|textValue|"children"/);
});

test('nested table column plans keep outer rows, cells, and grids unchanged', () => {
  const nested = table(row('Nested A', 'Nested B', 'Nested C'));
  const part = bodyPart(
    table(`<w:tr><w:tc>${p('Outer A')}${nested}</w:tc>${cell('Outer B')}${cell('Outer C')}</w:tr>`)
  );
  const outer = tables(part)[0]!,
    outerRow = rows(part, outer.id)[0]!,
    outerCells = cells(outerRow);
  const inner = children(outerCells[0]!).find((node) => node.kind === 'table')!;
  const input = { revision, part, tableId: inner.id },
    middle = columns(part, inner.id)[1]!;
  const result = apply(part, plan(input, { kind: 'delete', gridColumnId: middle.id }));
  expect(values(result, inner.id)).toEqual([['Nested A', 'Nested C']]);
  expect(findNode(result, outerCells[1]!.id) === outerCells[1]).toBe(true);
  expect(findNode(result, outerCells[2]!.id) === outerCells[2]).toBe(true);
  for (const node of columns(part, outer.id)) expect(findNode(result, node.id) === node).toBe(true);
  expect(
    planTableColumn({ ...input, tableId: outer.id }, { kind: 'delete', gridColumnId: middle.id })
      .success
  ).toBe(false);
});

test.each(['regular', 'merged'] as const)(
  'column insert/delete preserves native shape and surviving Word properties through save/reopen; %s',
  (shape) => {
    const input = shape === 'merged' ? mergedSource() : source();
    const pkgResult = readOoxmlPackage(
      docxFixture(serializeOoxmlPart(input.part).match(/<w:body>([\s\S]*)<\/w:body>/)![1]!)
    );
    if (!pkgResult.ok) throw new Error(pkgResult.reason);
    const pkg = pkgResult.package,
      original = pkg.parts.get(pkg.mainDocumentPart)!,
      value = source(original);
    for (const kind of ['insert', 'delete'] as const) {
      const middle = columns(original, value.tableId)[1]!;
      const changed = apply(
        original,
        plan(
          value,
          kind === 'insert'
            ? { kind, where: 'left', gridColumnId: middle.id }
            : { kind, gridColumnId: middle.id }
        )
      );
      const saved = readOoxmlPackage(writeOoxmlPackage(withPart(pkg, changed)));
      if (!saved.ok) throw new Error(saved.reason);
      const reopened = saved.package.parts.get(saved.package.mainDocumentPart)!,
        reopenedTable = tables(reopened)[0]!;
      expect(canonicalOoxmlFingerprint(reopened)).toBe(canonicalOoxmlFingerprint(changed));
      expect(widths(reopened, reopenedTable.id)).toEqual(widths(changed, value.tableId));
      expect(values(reopened, reopenedTable.id)).toEqual(values(changed, value.tableId));
      expect(canonicalOoxmlFingerprint(tables(reopened)[1]!)).toBe(
        canonicalOoxmlFingerprint(tables(original)[1]!)
      );
    }
  }
);

test('serialized plans contain only existing anchors and remain detached from caller targets', () => {
  const input = source(),
    target: { kind: 'insert'; where: 'left' | 'right'; gridColumnId: string } = {
      kind: 'insert',
      where: 'left',
      gridColumnId: columns(input.part, input.tableId)[1]!.id,
    };
  const planned = plan(input, target),
    raw = serializeEditPlan(planned);
  expect(decodeEditPlan(raw)).toEqual({ success: true, plan: planned });
  expect(raw).not.toMatch(/created|newCell|paragraphId|textValue|"children"|"author"|resultPart/);
  expect(Object.isFrozen(planned.operations[0])).toBe(true);
  expect(Object.isFrozen(target)).toBe(false);
  target.where = 'right';
  expect(serializeEditPlan(planned)).toBe(raw);
});

test('invalid column inputs, cross-table anchors, getters, and excess data refuse without planning partial operations', () => {
  const input = source(),
    middle = columns(input.part, input.tableId)[1]!,
    other = tables(input.part)[1]!;
  const base = { kind: 'insert', where: 'left', gridColumnId: middle.id },
    before = canonicalOoxmlFingerprint(input.part);
  let calls = 0;
  const getter = Object.defineProperty({ ...base }, 'gridColumnId', {
    enumerable: true,
    get() {
      calls++;
      return middle.id;
    },
  });
  const cyclic: Record<string, unknown> = { ...base };
  cyclic.extra = cyclic;
  const invalid: unknown[] = [
    null,
    [],
    {},
    getter,
    cyclic,
    { ...base, kind: 'replace' },
    { ...base, where: 'above' },
    { ...base, gridColumnId: '' },
    { ...base, gridColumnId: 'missing' },
    { ...base, gridColumnId: columns(input.part, other.id)[1]!.id },
    { ...base, referenceCellId: cells(rows(input.part, input.tableId)[0]!)[1]!.id },
    { ...base, values: ['Cannot hide new cell contents'] },
    { ...base, widthTwips: 1000 },
    { kind: 'delete' },
    { kind: 'delete', gridColumnId: middle.id, where: 'left' },
    {
      kind: 'insert',
      where: 'left',
      referenceCellId: cells(rows(input.part, input.tableId)[0]!)[1]!.id,
    },
    { ...base, gridColumnId: 'x'.repeat(16 * 1024 * 1024) },
  ];
  for (const target of invalid)
    expect(planTableColumn(input, target as TableColumnTarget).success).toBe(false);
  expect(planTableColumn({ ...input, tableId: 'missing' }, base as TableColumnTarget).success).toBe(
    false
  );
  expect(
    planTableColumn(
      Object.defineProperty({ ...input }, 'tableId', {
        get() {
          calls++;
          return input.tableId;
        },
      }),
      base as TableColumnTarget
    ).success
  ).toBe(false);
  expect(calls).toBe(0);
  expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
});

test('native column limits, final-column deletion, malformed topology, and protection refuse atomically', () => {
  const bodies = [
    table(`<w:tr>${cell('Only')}</w:tr>`, grid([2400])),
    table(`<w:tr>${cell('A')}${cell('B')}</w:tr>`),
    table(
      `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>${cell('A')}${cell('B')}${cell('C')}</w:tr>`
    ),
    table(`<w:tr>${cell('A', 2400, '<w:hMerge w:val="restart"/>')}${cell('B')}${cell('C')}</w:tr>`),
  ];
  for (const markup of bodies) {
    const input = source(bodyPart(markup)),
      before = canonicalOoxmlFingerprint(input.part);
    expect(
      planTableColumn(input, {
        kind: 'delete',
        gridColumnId: columns(input.part, input.tableId)[0]!.id,
      }).success
    ).toBe(false);
    expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
  }
  const max = source(
    bodyPart(
      table(
        `<w:tr>${Array.from({ length: 1024 }, () => cell('')).join('')}</w:tr>`,
        grid(Array.from({ length: 1024 }, () => 1200))
      )
    )
  );
  expect(
    planTableColumn(max, {
      kind: 'insert',
      where: 'left',
      gridColumnId: columns(max.part, max.tableId)[0]!.id,
    }).success
  ).toBe(false);
  for (const restriction of [
    '<w:lock w:val="sdtContentLocked"/>',
    '<w:dataBinding w:storeItemID="{00000000-0000-0000-0000-000000000001}" w:xpath="/item"/>',
  ]) {
    const part = bodyPart(
      `<w:sdt><w:sdtPr>${restriction}</w:sdtPr><w:sdtContent>${table(row('A', 'B', 'C'))}</w:sdtContent></w:sdt>`
    );
    const control = children(bodyStoryRoot(part)!)[0]!,
      tableNode = children(child(control, 'sdtContent')!)[0]!;
    const input = { revision, part, tableId: tableNode.id };
    const anchor = columns(part, tableNode.id)[1]!.id;
    expect(
      planTableColumn(input, { kind: 'insert', where: 'left', gridColumnId: anchor }).success
    ).toBe(false);
    expect(planTableColumn(input, { kind: 'delete', gridColumnId: anchor }).success).toBe(false);
  }
});

test('column planning preserves live package bytes, revision, selection, and existing undo/redo history', () => {
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
    const pkg = session.currentPackage(),
      input = source(session.part()),
      current = session.packageRevision();
    const bytes = writeOoxmlPackage(pkg),
      selection = surface.state().selection,
      history = [session.canUndo(), session.canRedo()];
    const middle = columns(input.part, input.tableId)[1]!.id;
    expect(
      planTableColumn(
        { ...input, revision: { ...revision, revision: current } },
        { kind: 'insert', where: 'left', gridColumnId: middle }
      ).success
    ).toBe(true);
    expect(
      planTableColumn(
        { ...input, revision: { ...revision, revision: current } },
        { kind: 'delete', gridColumnId: middle }
      ).success
    ).toBe(true);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(session.packageRevision()).toBe(current);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
  } finally {
    editor.destroy();
  }
});
