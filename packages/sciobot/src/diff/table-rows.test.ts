import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  escapeXml,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  serializeOoxmlPart,
  withPart,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import reference from '../__tests__/fixtures/ai-pairs.json';
import { bodyPart, docxFixture } from '../../test/support/document';
import { decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';
import { planTableRow, type TableRowSource, type TableRowTarget } from './table-rows';

const revision = { documentId: 'rows', epoch: 'load-1', revision: 4 };
const p = (text: string) =>
  `<w:p><w:pPr><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
const cell = (text: string, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:shd w:fill="EEEEEE"/>${extra}</w:tcPr>${p(text)}</w:tc>`;
const row = (first: string, second: string, props = '') =>
  `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="360"/>${props}</w:trPr>${cell(first)}${cell(second)}</w:tr>`;
const table = (rows: string) =>
  `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblCaption w:val="Keep table metadata"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>${rows}</w:tbl>`;
const body = () =>
  p('Before.') +
  table(row('A', 'B', '<w:tblHeader/>') + row('One', 'Two')) +
  p('Between.') +
  table(row('Other1', 'Other2') + row('Other3', 'Other4')) +
  p('After.');
function children(node: OoxmlNode): readonly OoxmlNode[] {
  return node.kind === 'textValue' ? [] : node.children;
}
function child(node: OoxmlNode, name: string) {
  return children(node).find((entry) => entry.kind !== 'textValue' && entry.localName === name);
}
function tables(part: OoxmlPart) {
  return children(bodyStoryRoot(part)!).filter((entry) => entry.kind === 'table');
}
function rows(part: OoxmlPart, id: string) {
  return children(findNode(part, id)!).filter((entry) => entry.kind === 'tableRow');
}
function cells(node: OoxmlNode) {
  return children(node).filter((entry) => entry.kind === 'tableCell');
}
function text(part: OoxmlPart, node: OoxmlNode): string {
  return node.kind === 'paragraph'
    ? paragraphTextOf(part, node.id)!
    : children(node)
        .map((child) => text(part, child))
        .join('');
}
function source(part = bodyPart(body()), tableIndex = 0): TableRowSource {
  return { revision, part, tableId: tables(part)[tableIndex]!.id };
}
function plan(input: TableRowSource, target: TableRowTarget): CanonicalEditPlan {
  const result = planTableRow(input, target);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
/** Immutable native oracle. A01/R08 own guarded transactions, tracking, and collaboration. */
function apply(part: OoxmlPart, planned: CanonicalEditPlan, tracked = false): OoxmlPart {
  let current = part;
  for (const edit of planned.operations) {
    if (edit.kind !== 'table') throw new Error('Unexpected non-table edit');
    const operation = {
      ...edit.operation,
      ...(tracked ? { revision: { author: 'ScioBot', date: '2026-10-19T12:00:00Z' } } : {}),
    } as TreeDocOp;
    const result = applyTreeOp(current, operation);
    if (!result.ok) throw new Error(result.reason);
    current = result.part;
  }
  return current;
}

test('C01 row insertion adds Three/Four while preserving both original cells, rows, grid, properties, and the other table', () => {
  const captured = reference.cases.find((item) => item.name === 'tables and links')!;
  expect(captured.beforeMarkdown).toContain('| One | Two |');
  expect(captured.afterMarkdown).toContain('| Three | Four |');
  const input = source(),
    oldRows = rows(input.part, input.tableId),
    oldCells = cells(oldRows[1]!);
  const fingerprint = canonicalOoxmlFingerprint(input.part),
    other = tables(input.part)[1]!,
    original = findNode(input.part, input.tableId)!;
  const target: TableRowTarget = {
    kind: 'insert',
    rowId: oldRows[1]!.id,
    where: 'below',
    values: ['Three', 'Four'],
  };
  const planned = plan(input, target);
  expect(planned.operations).toEqual([
    {
      kind: 'table',
      partName: input.part.name,
      operation: {
        op: 'authorTable',
        action: {
          kind: 'existing',
          tableId: input.tableId,
          mutation: {
            kind: 'insertRows',
            rowId: oldRows[1]!.id,
            location: 'after',
            count: 1,
            values: [['Three', 'Four']],
          },
        },
      },
    },
  ]);
  const result = apply(input.part, planned),
    after = rows(result, input.tableId);
  expect(after.map((row) => cells(row).map((cell) => text(result, cell)))).toEqual([
    ['A', 'B'],
    ['One', 'Two'],
    ['Three', 'Four'],
  ]);
  expect(after.slice(0, 2).every((row, at) => row === oldRows[at])).toBe(true);
  expect(cells(after[1]!).every((cell, at) => cell === oldCells[at])).toBe(true);
  for (const name of ['tblPr', 'tblGrid'])
    expect(child(findNode(result, input.tableId)!, name) === child(original, name)).toBe(true);
  expect(findNode(result, other.id) === other).toBe(true);
  expect(serializeOoxmlPart(result)).toContain('w:trHeight w:val="360"');
  expect(canonicalOoxmlFingerprint(input.part)).toBe(fingerprint);
  expect(planned.operations).toHaveLength(1);
});

test('inserted row properties and text face follow native inheritance without copying source content or identities', () => {
  const input = source(),
    anchor = rows(input.part, input.tableId)[1]!;
  const result = apply(
    input.part,
    plan(input, { kind: 'insert', rowId: anchor.id, where: 'above', values: ['', '😀\tŽluťoučký'] })
  );
  const inserted = rows(result, input.tableId)[1]!;
  expect(inserted.id).not.toBe(anchor.id);
  expect(cells(inserted).map((cell) => text(result, cell))).toEqual(['', '😀\tŽluťoučký']);
  expect(
    children(child(inserted, 'trPr')!).map((node) => node.kind !== 'textValue' && node.localName)
  ).toEqual(['cantSplit', 'trHeight']);
  expect(serializeOoxmlPart(result)).toContain('w:shd w:fill="EEEEEE"');
  const paragraph = children(cells(inserted)[1]!).find((node) => node.kind === 'paragraph')!;
  const run = children(paragraph).find((node) => node.kind === 'run')!;
  expect(child(child(run, 'rPr')!, 'i')).toBeDefined();
  expect(child(child(run, 'rPr')!, 'b')).toBeUndefined();
  expect(text(result, inserted)).not.toContain('One');
  expect(rows(result, input.tableId)[2] === anchor).toBe(true);
});

test.each(['above', 'below'] as const)(
  'blank row insertion %s uses native row creation without preview IDs',
  (where) => {
    const input = source(),
      anchor = rows(input.part, input.tableId)[0]!;
    const planned = plan(input, { kind: 'insert', rowId: anchor.id, where });
    expect(planned.operations).toEqual([
      {
        kind: 'table',
        partName: input.part.name,
        operation: { op: 'insertTableRow', tableId: input.tableId, rowId: anchor.id, where },
      },
    ]);
    const next = apply(input.part, planned),
      after = rows(next, input.tableId);
    const inserted = after[where === 'above' ? 0 : 1]!;
    expect(cells(inserted).map((cell) => text(next, cell))).toEqual(['', '']);
    expect(child(child(inserted, 'trPr')!, 'tblHeader')).toBeDefined();
    expect(serializeEditPlan(planned)).not.toMatch(
      /createdRowIds|resultPart|textValue|"children"|"author"/
    );
  }
);

test('deleting a row changes only the requested table and keeps surviving row and cell identities', () => {
  const input = source(),
    oldRows = rows(input.part, input.tableId),
    other = tables(input.part)[1]!;
  const planned = plan(input, {
    kind: 'delete',
    rowId: oldRows[1]!.id,
    referenceCellId: cells(oldRows[1]!)[1]!.id,
  });
  expect(planned.operations).toEqual([
    {
      kind: 'table',
      partName: input.part.name,
      operation: {
        op: 'deleteTableRow',
        tableId: input.tableId,
        rowId: oldRows[1]!.id,
        referenceCellId: cells(oldRows[1]!)[1]!.id,
      },
    },
  ]);
  const result = apply(input.part, planned);
  expect(rows(result, input.tableId)).toHaveLength(1);
  expect(rows(result, input.tableId)[0] === oldRows[0]).toBe(true);
  expect(findNode(result, oldRows[1]!.id)).toBeNull();
  expect(findNode(result, other.id) === other).toBe(true);
  expect(canonicalOoxmlFingerprint(findNode(result, other.id)!)).toBe(
    canonicalOoxmlFingerprint(other)
  );
});

test.each(['insert', 'delete'] as const)(
  'the same %s plan supports native direct and tracked semantic oracles',
  (kind) => {
    const input = source(),
      anchor = rows(input.part, input.tableId)[1]!;
    const target: TableRowTarget =
      kind === 'insert'
        ? { kind, rowId: anchor.id, where: 'below', values: ['Three', 'Four'] }
        : { kind, rowId: anchor.id };
    const planned = plan(input, target),
      direct = apply(input.part, planned),
      tracked = apply(input.part, planned, true);
    const accepted = applyTreeOp(tracked, { op: 'acceptAllRevisions' }),
      rejected = applyTreeOp(tracked, { op: 'rejectAllRevisions' });
    if (!accepted.ok || !rejected.ok) throw new Error('Native row decision failed');
    expect(canonicalOoxmlFingerprint(accepted.part)).toBe(canonicalOoxmlFingerprint(direct));
    expect(canonicalOoxmlFingerprint(rejected.part)).toBe(canonicalOoxmlFingerprint(input.part));
    const loaded = readOoxmlPackage(docxFixture(''));
    if (!loaded.ok) throw new Error(loaded.reason);
    const reopened = readOoxmlPackage(writeOoxmlPackage(withPart(loaded.package, direct)));
    if (!reopened.ok) throw new Error(reopened.reason);
    const table = tables(reopened.package.parts.get(reopened.package.mainDocumentPart)!)[0]!;
    expect(children(table).filter((node) => node.kind === 'tableRow')).toHaveLength(
      kind === 'insert' ? 3 : 1
    );
  }
);

test('native blank insertion extends vertical merges and row deletion repairs the surviving merge head', () => {
  const mergeRows = `<w:tr>${cell('Head', '<w:vMerge w:val="restart"/>')}${cell('B')}</w:tr><w:tr>${cell('', '<w:vMerge/>')}${cell('D')}</w:tr>`;
  const input = source(bodyPart(table(mergeRows) + table(row('Other1', 'Other2')))),
    original = rows(input.part, input.tableId),
    other = tables(input.part)[1]!;
  const inserted = apply(
    input.part,
    plan(input, { kind: 'insert', rowId: original[0]!.id, where: 'below' })
  );
  const newRow = rows(inserted, input.tableId)[1]!;
  expect(child(child(cells(newRow)[0]!, 'tcPr')!, 'vMerge')).toBeDefined();
  const removed = apply(input.part, plan(input, { kind: 'delete', rowId: original[0]!.id }));
  const survivor = cells(rows(removed, input.tableId)[0]!)[0]!;
  const marker = child(child(survivor, 'tcPr')!, 'vMerge');
  expect(marker).toBeUndefined();
  const longer = source(
    bodyPart(table(mergeRows + `<w:tr>${cell('', '<w:vMerge/>')}${cell('E')}</w:tr>`))
  );
  const longerRows = rows(longer.part, longer.tableId);
  const repaired = apply(longer.part, plan(longer, { kind: 'delete', rowId: longerRows[0]!.id }));
  const head = child(child(cells(rows(repaired, longer.tableId)[0]!)[0]!, 'tcPr')!, 'vMerge');
  expect(
    head?.kind === 'textValue'
      ? null
      : head?.attributes.find((attr) => attr.localName === 'val')?.value
  ).toBe('restart');
  expect(findNode(removed, other.id) === other).toBe(true);
  expect(
    planTableRow(input, {
      kind: 'insert',
      rowId: original[0]!.id,
      where: 'below',
      values: ['Blocked', 'Value'],
    }).success
  ).toBe(false);
});

test('nested table rows remain scoped to the selected table and do not target outer rows', () => {
  const nested = table(row('Nested1', 'Nested2') + row('Nested3', 'Nested4'));
  const part = bodyPart(table(`<w:tr><w:tc>${p('Outer')}${nested}</w:tc>${cell('Other')}</w:tr>`));
  const outer = tables(part)[0]!,
    outerRow = children(outer).find((node) => node.kind === 'tableRow')!;
  const nestedTable = children(cells(outerRow)[0]!).find((node) => node.kind === 'table')!,
    anchor = rows(part, nestedTable.id)[1]!;
  const input = { revision, part, tableId: nestedTable.id };
  expect(
    planTableRow({ ...input, tableId: outer.id }, { kind: 'delete', rowId: anchor.id })
  ).toMatchObject({ success: false, code: 'invalid-document' });
  const result = apply(part, plan(input, { kind: 'delete', rowId: anchor.id }));
  expect(rows(result, nestedTable.id)).toHaveLength(1);
  expect(cells(rows(result, outer.id)[0]!)[1] === cells(outerRow)[1]).toBe(true);
});

test('plans detach values and roundtrip only canonical anchors and requested content', () => {
  const input = source(),
    values = ['Three', 'Four'],
    anchor = rows(input.part, input.tableId)[1]!;
  const planned = plan(input, { kind: 'insert', rowId: anchor.id, where: 'below', values });
  values[0] = 'Changed caller';
  expect(Object.isFrozen(values)).toBe(false);
  expect(serializeEditPlan(planned)).toContain('Three');
  expect(serializeEditPlan(planned)).not.toContain('Changed caller');
  expect(Object.isFrozen(planned.operations[0])).toBe(true);
  expect(decodeEditPlan(serializeEditPlan(planned))).toEqual({ success: true, plan: planned });
});

test('invalid row, table, reference cell, values, getters, and resource excess refuse atomically', () => {
  const input = source(),
    anchor = rows(input.part, input.tableId)[1]!,
    other = tables(input.part)[1]!,
    before = canonicalOoxmlFingerprint(input.part);
  let calls = 0;
  const getter = Object.defineProperty(
    { kind: 'insert', rowId: anchor.id, where: 'below' },
    'values',
    {
      enumerable: true,
      get() {
        calls++;
        return ['Bad', 'Bad'];
      },
    }
  );
  const cyclic: unknown[] = [];
  cyclic.push(cyclic);
  const base = { kind: 'insert', rowId: anchor.id, where: 'below' };
  const targets: unknown[] = [
    null,
    [],
    {},
    { ...base, where: 'after' },
    { ...base, rowId: 'missing' },
    { ...base, values: ['Too short'] },
    { ...base, values: [1, 2] },
    { ...base, values: 'Bad' },
    { ...base, values: undefined },
    { ...base, values: ['Line\nbreak', 'Bad'] },
    { ...base, values: ['x'.repeat(1_000_001), 'Bad'] },
    { ...base, values: cyclic },
    getter,
    { ...base, revision: { author: 'Injected' } },
    {
      kind: 'delete',
      rowId: anchor.id,
      referenceCellId: cells(rows(input.part, other.id)[0]!)[0]!.id,
    },
  ];
  for (const target of targets)
    expect(planTableRow(input, target as TableRowTarget).success).toBe(false);
  expect(calls).toBe(0);
  expect(
    planTableRow(input, { kind: 'delete', rowId: rows(input.part, other.id)[0]!.id })
  ).toMatchObject({ success: false, code: 'invalid-document' });
  const one = source(bodyPart(table(row('A', 'B'))));
  expect(
    planTableRow(one, { kind: 'delete', rowId: rows(one.part, one.tableId)[0]!.id }).success
  ).toBe(false);
  expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
});

test('locked and bound table content refuses before a plan can be applied', () => {
  for (const restriction of [
    '<w:lock w:val="sdtContentLocked"/>',
    '<w:dataBinding w:storeItemID="{00000000-0000-0000-0000-000000000001}" w:xpath="/item"/>',
  ]) {
    const part = bodyPart(
      `<w:sdt><w:sdtPr>${restriction}</w:sdtPr><w:sdtContent>${table(row('A', 'B') + row('C', 'D'))}</w:sdtContent></w:sdt>`
    );
    const control = children(bodyStoryRoot(part)!)[0]!,
      content = child(control, 'sdtContent')!,
      nativeTable = children(content)[0]!;
    const input = { revision, part, tableId: nativeTable.id },
      anchor = rows(part, nativeTable.id)[1]!;
    expect(
      planTableRow(input, { kind: 'insert', rowId: anchor.id, where: 'below', values: ['X', 'Y'] })
        .success
    ).toBe(false);
    expect(planTableRow(input, { kind: 'delete', rowId: anchor.id }).success).toBe(false);
  }
});

test('planning preserves live bytes, selection, revision, and existing history', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(body()),
  });
  try {
    const surface = editor.surface!,
      session = surface.session,
      paragraphId = session.paragraphIds()[0]!;
    surface.applyAutomationOps(() => [{ op: 'insertText', paragraphId, offset: 0, text: 'User ' }]);
    const pkg = session.currentPackage(),
      part = pkg.parts.get(pkg.mainDocumentPart)!,
      input = source(part),
      anchor = rows(part, input.tableId)[1]!;
    const current = session.packageRevision(),
      bytes = writeOoxmlPackage(pkg),
      selection = surface.state().selection,
      history = [session.canUndo(), session.canRedo()];
    expect(
      planTableRow(
        { ...input, revision: { ...revision, revision: current } },
        { kind: 'insert', rowId: anchor.id, where: 'below', values: ['Three', 'Four'] }
      ).success
    ).toBe(true);
    expect(session.packageRevision()).toBe(current);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
  } finally {
    editor.destroy();
  }
});
