import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { materializeInsertedColumn } from '../../test/support/inserted-column-oracle';
import {
  namedChild,
  nodeChildren,
  rowCells,
  tableRows,
} from '../../test/support/native-table-oracle';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { MDX_BLOCK } from '../projection/mdx';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { nativeGridColumns } from './inserted-column-native';
import {
  planInsertedTableColumn,
  type InsertedTableColumnSource,
  type InsertedTableColumnTarget,
} from './inserted-table-columns';
import { createEditPlan, decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';

const revision = { documentId: 'rich-columns', epoch: 'load-1', revision: 9 };
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width = 2400, props = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:shd w:fill="EEEEEE"/><keep:extra xmlns:keep="urn:keep" keep:value="retain"/>${props}</w:tcPr>${p(text)}</w:tc>`;
const row = (a: string, b: string, c: string) =>
  `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(a, 1200)}${cell(b)}${cell(c, 3600)}</w:tr>`;
const grid =
  '<w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid>';
const table = (rows: string, columns = grid) =>
  `<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblCaption w:val="Keep"/></w:tblPr>${columns}${rows}</w:tbl>`;
const body = () =>
  p('Before') +
  table(row('A1', 'B1', 'C1') + row('A2', 'B2', 'C2')) +
  p('Between') +
  table(row('Other A', 'Other B', 'Other C')) +
  p('After');
const partOf = (pkg: OoxmlPackage) => pkg.parts.get(pkg.mainDocumentPart)!;
const tables = (pkg: OoxmlPackage) =>
  nodeChildren(bodyStoryRoot(partOf(pkg))!).filter((node) => node.kind === 'table');
const attr = (node: OoxmlNode, name: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find((entry) => entry.localName === name)?.value;
function source(xml = body()): InsertedTableColumnSource {
  const loaded = readOoxmlPackage(docxFixture(xml));
  if (!loaded.ok) throw new Error(loaded.reason);
  return {
    revision,
    pkg: loaded.package,
    partName: loaded.package.mainDocumentPart,
    tableId: tables(loaded.package)[0]!.id,
  };
}
function ast(markdown: string): TargetAstNode[] {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  return JSON.parse(JSON.stringify(parsed.target.normalizedAst.children));
}
function target(input: InsertedTableColumnSource): InsertedTableColumnTarget {
  return {
    anchor: { kind: 'grid-column', id: nativeGridColumns(partOf(input.pkg), input.tableId)[1]!.id },
    where: 'left',
    cells: tableRows(partOf(input.pkg), input.tableId).map((row) => ({
      rowId: row.id,
      blocks: [],
    })),
  };
}
function richTarget(input: InsertedTableColumnSource): InsertedTableColumnTarget {
  const base = target(input),
    first = ast(
      '**Bold** [Link](https://example.com/new?x=1&y=2)  \nNext 😀\n\nSecond paragraph.\n\n<activity name="New" duration="5">\n\nApplication child.\n\n</activity>'
    );
  const widget = ast('<Widget state={globalThis.__columnProbe?.()}>\n\n<Child />\n\n</Widget>')[0]!;
  return {
    ...base,
    properties: { widthTwips: 1800 },
    cells: [
      { rowId: base.cells[0]!.rowId, blocks: first },
      { rowId: base.cells[1]!.rowId, blocks: [widget, ...ast('Widget tail.')] },
    ],
    payloads: [
      {
        path: [0, 2],
        label: 'New activity',
        payload: {
          id: 'column-activity',
          type: 'lesson_activity',
          version: 1,
          attributes: { name: 'New', duration: 5, privateState: { keep: true } },
          unknownRoot: ['keep'],
        },
      },
      {
        path: [1, 0],
        label: 'Widget fallback',
        payload: {
          id: 'column-widget',
          type: MDX_BLOCK,
          version: 1,
          attributes: { asset: 'cached-asset' },
          mdast: widget,
          vendor: { keep: true },
        },
      },
    ],
  };
}
function plan(
  input: InsertedTableColumnSource,
  next: InsertedTableColumnTarget
): CanonicalEditPlan {
  const result = planInsertedTableColumn(input, next);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
function descendants(node: OoxmlNode): OoxmlNode[] {
  return [node, ...nodeChildren(node).flatMap(descendants)];
}
const text = (pkg: OoxmlPackage, node: OoxmlNode) =>
  descendants(node)
    .filter((node) => node.kind === 'paragraph')
    .map((node) => paragraphTextOf(partOf(pkg), node.id))
    .filter(Boolean);
const widths = (pkg: OoxmlPackage, tableId = tables(pkg)[0]!.id) =>
  nativeGridColumns(partOf(pkg), tableId).map((node) => Number(attr(node, 'w')));
function reopen(pkg: OoxmlPackage): OoxmlPackage {
  const loaded = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!loaded.ok) throw new Error(loaded.reason);
  return loaded.package;
}

test.each(['left', 'right'] as const)(
  'rich middle column retains bold text, referenced links, breaks, paragraphs, payloads, and explicit width through native insertion and save/reopen; %s',
  (where) => {
    const input = source(),
      next = { ...richTarget(input), where },
      original = partOf(input.pkg);
    const oldRows = tableRows(original, input.tableId),
      oldCells = oldRows.flatMap(rowCells),
      other = tables(input.pkg)[1]!;
    const before = writeOoxmlPackage(input.pkg),
      planned = plan(input, next);
    const decoded = decodeEditPlan(serializeEditPlan(planned));
    if (!decoded.success) throw new Error(decoded.error);
    const result = materializeInsertedColumn(input.pkg, decoded.plan),
      part = partOf(result.pkg);
    const expectedWidths = where === 'left' ? [1200, 1800, 2400, 3600] : [1200, 2400, 1800, 3600];
    expect(widths(result.pkg)).toEqual(expectedWidths);
    expect(widths(result.pkg).reduce((sum, width) => sum + width, 0)).toBe(9000);
    expect(attr(findNode(part, result.gridColumnId)!, 'w')).toBe('1800');
    expect(
      attr(namedChild(namedChild(findNode(part, input.tableId)!, 'tblPr')!, 'tblW')!, 'w')
    ).toBe('9000');
    expect(
      attr(namedChild(namedChild(findNode(part, input.tableId)!, 'tblPr')!, 'tblLayout')!, 'type')
    ).toBe('fixed');
    const newFirst = findNode(part, result.cellIds.get(oldRows[0]!.id)!)!,
      newSecond = findNode(part, result.cellIds.get(oldRows[1]!.id)!)!;
    expect(text(result.pkg, newFirst)).toEqual([
      'Bold Link\nNext 😀',
      'Second paragraph.',
      'Application child.',
    ]);
    expect(text(result.pkg, newSecond)).toEqual(['Widget fallback', 'Widget tail.']);
    for (const node of [...oldRows, ...oldCells]) expect(findNode(part, node.id)?.id).toBe(node.id);
    for (const cell of oldCells)
      expect(canonicalOoxmlFingerprint(findNode(part, cell.id)!)).toBe(
        canonicalOoxmlFingerprint(cell)
      );
    expect(findNode(part, other.id) === other).toBe(true);
    for (const pkg of [result.pkg, reopen(result.pkg)]) {
      const indexed = buildBlockIndex(pkg);
      expect(indexed.diagnostics).toEqual([]);
      expect(indexed.byId.get('column-activity')?.payload).toEqual(next.payloads![0]!.payload);
      expect(indexed.byId.get('column-widget')?.payload).toEqual(next.payloads![1]!.payload);
      expect(widths(pkg)).toEqual(expectedWidths);
      const savedPart = partOf(pkg),
        savedRows = tableRows(savedPart, tables(pkg)[0]!.id),
        saved = rowCells(savedRows[0]!)[where === 'left' ? 1 : 2]!;
      expect(text(pkg, saved)).toEqual(text(result.pkg, newFirst));
      const bold = descendants(saved).find(
        (node) =>
          node.kind === 'run' &&
          descendants(node).some((child) => child.kind === 'textValue' && child.value === 'Bold')
      )!;
      expect(namedChild(namedChild(bold, 'rPr')!, 'b')).toBeDefined();
      const links = descendants(saved).filter(
        (node) => node.kind !== 'textValue' && node.localName === 'hyperlink'
      );
      expect(links).toHaveLength(1);
      expect(
        pkg.externalTargets.find((entry) => entry.id === attr(links[0]!, 'id'))?.rawTarget
      ).toBe('https://example.com/new?x=1&y=2');
      expect(
        descendants(saved).some((node) => node.kind !== 'textValue' && node.localName === 'br')
      ).toBe(true);
      expect(canonicalOoxmlFingerprint(tables(pkg)[1]!)).toBe(canonicalOoxmlFingerprint(other));
      expect(canonicalOoxmlFingerprint(savedPart)).toBe(canonicalOoxmlFingerprint(part));
    }
    expect(writeOoxmlPackage(input.pkg)).toEqual(before);
  }
);

test('row anchors route reordered streams correctly and omission retains native inherited grid width', () => {
  const input = source(),
    base = target(input),
    first = base.cells[0]!.rowId,
    second = base.cells[1]!.rowId;
  const next = {
    ...base,
    cells: [
      { rowId: second, blocks: ast('Second row 😀') },
      { rowId: first, blocks: ast('First row') },
    ],
  };
  const planned = plan(input, next),
    edit = planned.operations[0]!;
  expect(edit.kind === 'insert-column' && edit.intent.properties).toEqual({});
  expect(edit.kind === 'insert-column' && edit.intent.payloads).toEqual([]);
  const result = materializeInsertedColumn(input.pkg, planned),
    part = partOf(result.pkg);
  expect(text(result.pkg, findNode(part, result.cellIds.get(first)!)!)).toEqual(['First row']);
  expect(text(result.pkg, findNode(part, result.cellIds.get(second)!)!)).toEqual(['Second row 😀']);
  expect(widths(result.pkg)).toEqual([1200, 2400, 2400, 3600]);
});

function mergedSource(): InsertedTableColumnSource {
  return source(
    p('Before') +
      table(
        `<w:tr>${cell('AB', 3600, '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>')}${cell('C1', 3600)}</w:tr>` +
          `<w:tr>${cell('', 3600, '<w:gridSpan w:val="2"/><w:vMerge/>')}${cell('C2', 3600)}</w:tr>` +
          row('A3', 'B3', 'C3')
      ) +
      p('After')
  );
}

test('a column that widens every existing row needs no new cell destination and accepts native minimum width', () => {
  const input = source(
    p('Before') +
      table(
        `<w:tr>${cell('Whole A', 7200, '<w:gridSpan w:val="3"/>')}</w:tr>` +
          `<w:tr>${cell('Whole B', 7200, '<w:gridSpan w:val="3"/>')}</w:tr>`
      ) +
      p('After')
  );
  const oldRows = tableRows(partOf(input.pkg), input.tableId),
    next = { ...target(input), properties: { widthTwips: 300 } };
  const result = materializeInsertedColumn(input.pkg, plan(input, next)),
    part = partOf(result.pkg);
  expect(result.cellIds.size).toBe(0);
  expect(widths(result.pkg)).toEqual([1200, 300, 2400, 3600]);
  for (const row of oldRows) {
    const oldCell = rowCells(row)[0]!,
      saved = findNode(part, oldCell.id)!;
    expect(attr(namedChild(namedChild(saved, 'tcPr')!, 'gridSpan')!, 'val')).toBe('4');
    expect(attr(namedChild(namedChild(saved, 'tcPr')!, 'tcW')!, 'w')).toBe('7500');
    expect(text(result.pkg, saved)).toEqual(text(input.pkg, oldCell));
  }
  expect(canonicalOoxmlFingerprint(partOf(reopen(result.pkg)))).toBe(
    canonicalOoxmlFingerprint(part)
  );
});

test('insertion inside a span keeps its native content while rich fragments fill only newly created cells', () => {
  const input = mergedSource(),
    base = target(input),
    part = partOf(input.pkg),
    oldRows = tableRows(part, input.tableId);
  const original = oldRows.slice(0, 2).map((row) => rowCells(row)[0]!);
  const next = {
    ...base,
    properties: { widthTwips: 1800 },
    cells: base.cells.map((entry, index) => ({
      ...entry,
      blocks: index === 2 ? ast('**New cell**\n\nSecond paragraph.') : [],
    })),
  };
  const result = materializeInsertedColumn(input.pkg, plan(input, next)),
    after = partOf(result.pkg);
  expect(result.cellIds.size).toBe(1);
  expect(widths(result.pkg)).toEqual([1200, 1800, 2400, 3600]);
  for (const cell of original) {
    const saved = findNode(after, cell.id)!;
    expect(attr(namedChild(namedChild(saved, 'tcPr')!, 'gridSpan')!, 'val')).toBe('3');
    expect(attr(namedChild(namedChild(saved, 'tcPr')!, 'tcW')!, 'w')).toBe('5400');
    expect(text(result.pkg, saved)).toEqual(text(input.pkg, cell));
    expect(
      nodeChildren(saved).find((node) => node.kind === 'paragraph') ===
        nodeChildren(cell).find((node) => node.kind === 'paragraph')
    ).toBe(true);
  }
  const reopened = reopen(result.pkg);
  expect(canonicalOoxmlFingerprint(partOf(reopened))).toBe(canonicalOoxmlFingerprint(after));
  expect(
    planInsertedTableColumn(input, {
      ...next,
      cells: next.cells.map((entry, index) =>
        index ? entry : { ...entry, blocks: ast('Would overwrite merged content') }
      ),
    }).success
  ).toBe(false);
});

test('new vertical continuation cells require empty streams while visible new cells retain content', () => {
  const input = mergedSource(),
    base = target(input),
    next = {
      ...base,
      where: 'right' as const,
      cells: base.cells.map((entry, index) => ({
        ...entry,
        blocks: index === 1 ? [] : ast(`Visible ${index}`),
      })),
    };
  const result = materializeInsertedColumn(input.pkg, plan(input, next)),
    part = partOf(result.pkg);
  expect(result.cellIds.size).toBe(3);
  const hiddenId = result.cellIds.get(base.cells[1]!.rowId)!;
  expect(text(result.pkg, findNode(part, hiddenId)!)).toEqual([]);
  expect(namedChild(namedChild(findNode(part, hiddenId)!, 'tcPr')!, 'vMerge')).toBeDefined();
  expect(
    planInsertedTableColumn(input, {
      ...next,
      cells: next.cells.map((entry, index) =>
        index === 1 ? { ...entry, blocks: ast('Hidden loss') } : entry
      ),
    })
  ).toMatchObject({ success: false, code: 'unsupported-operation' });
});

test('implicit regular grids resolve the new column from a canonical first-row cell', () => {
  const input = source(
    p('Before') + table(row('A1', 'B1', 'C1') + row('A2', 'B2', 'C2'), '') + p('After')
  );
  const rows = tableRows(partOf(input.pkg), input.tableId),
    reference = rowCells(rows[0]!)[1]!;
  const next: InsertedTableColumnTarget = {
    anchor: { kind: 'reference-cell', id: reference.id },
    where: 'left',
    properties: { widthTwips: 1800 },
    cells: rows.map((row, index) => ({ rowId: row.id, blocks: ast(`Implicit ${index}`) })),
  };
  const result = materializeInsertedColumn(input.pkg, plan(input, next));
  expect(widths(result.pkg)).toEqual([1200, 1800, 2400, 3600]);
  for (const [index, row] of rows.entries())
    expect(text(result.pkg, findNode(partOf(result.pkg), result.cellIds.get(row.id)!)!)).toEqual([
      `Implicit ${index}`,
    ]);
  expect(canonicalOoxmlFingerprint(partOf(reopen(result.pkg)))).toBe(
    canonicalOoxmlFingerprint(partOf(result.pkg))
  );
});

test('native allocation changes after planning cannot redirect fragments or the requested new width', () => {
  const input = source(),
    next = richTarget(input),
    planned = plan(input, next),
    original = partOf(input.pkg);
  const allocated = applyTreeOp(original, {
    op: 'insertTableColumn',
    tableId: input.tableId,
    gridColumnId: nativeGridColumns(original, input.tableId)[0]!.id,
    where: 'left',
  });
  if (!allocated.ok) throw new Error(allocated.reason);
  // Immutable allocation oracle only. A01/R09 must reject stale revisions before live mutation.
  const result = materializeInsertedColumn(withPart(input.pkg, allocated.part), planned),
    part = partOf(result.pkg);
  expect(widths(result.pkg)).toEqual([1200, 1200, 1800, 2400, 3600]);
  const firstRow = findNode(part, next.cells[0]!.rowId)!;
  expect(text(result.pkg, rowCells(firstRow)[0]!)).toEqual([]);
  expect(text(result.pkg, rowCells(firstRow)[2]!)).toEqual([
    'Bold Link\nNext 😀',
    'Second paragraph.',
    'Application child.',
  ]);
  const ids = descendants(part.root).map((node) => node.id);
  expect(new Set(ids).size).toBe(ids.length);
});

test('opaque expressions remain inert, while serialized plans retain only detached new content and original anchors', () => {
  const probe = globalThis as unknown as Record<string, unknown>;
  let calls = 0;
  probe.__columnProbe = () => {
    calls++;
    throw new Error('Opaque expression executed');
  };
  try {
    const input = source(),
      next = richTarget(input),
      planned = plan(input, next),
      raw = serializeEditPlan(planned);
    expect(decodeEditPlan(raw)).toEqual({ success: true, plan: planned });
    expect(raw).not.toMatch(
      /preview|createdRow|newCellId|paragraphId|controlId|textValue|resultPart/
    );
    expect(raw).not.toContain('Other A');
    expect(raw).not.toContain('B1');
    expect(Object.isFrozen(planned.operations[0])).toBe(true);
    expect(Object.isFrozen(next.cells)).toBe(false);
    materializeInsertedColumn(input.pkg, planned);
    expect(calls).toBe(0);
    (next.payloads![0]!.payload.attributes as Record<string, unknown>).name = 'Caller mutation';
    expect(serializeEditPlan(planned)).toBe(raw);
  } finally {
    delete probe.__columnProbe;
  }
});

test('row ownership, binding consistency, native width limits, and unknown sizing refuse without writes', () => {
  const input = source(),
    next = richTarget(input),
    base = target(input),
    before = writeOoxmlPackage(input.pkg);
  const invalid: unknown[] = [
    null,
    [],
    {},
    { ...next, cells: next.cells.slice(1) },
    { ...next, cells: [next.cells[0], next.cells[0]] },
    {
      ...next,
      cells: [
        { ...next.cells[0], rowId: tableRows(partOf(input.pkg), tables(input.pkg)[1]!.id)[0]!.id },
        next.cells[1],
      ],
    },
    { ...next, payloads: [] },
    { ...next, payloads: [next.payloads![0], next.payloads![0]] },
    { ...next, payloads: [{ ...next.payloads![0], path: [0, 0] }, next.payloads![1]] },
    { ...next, properties: null },
    { ...next, payloads: null },
    { ...base, properties: { widthTwips: 299 } },
    { ...base, properties: { widthTwips: 31681 } },
    { ...base, properties: { widthTwips: 31680 } },
    { ...base, properties: { widthTwips: 1800.5 } },
    { ...base, properties: { heightTwips: 1800 } },
    { ...base, nativeCellId: 'preview-id' },
    { ...base, anchor: { kind: 'grid-column', id: 'missing' } },
    { ...base, anchor: { kind: 'cell', id: base.anchor.id } },
    { ...base, anchor: { ...base.anchor, other: 'forged' } },
    {
      ...base,
      cells: [
        { ...base.cells[0], blocks: [{ type: 'text', value: 'Needs paragraph' }] },
        base.cells[1],
      ],
    },
    {
      ...base,
      cells: [
        { ...base.cells[0], blocks: [{ type: 'paragraph', kind: 'paragraph', children: [] }] },
        base.cells[1],
      ],
    },
  ];
  for (const value of invalid)
    expect(planInsertedTableColumn(input, value as InsertedTableColumnTarget).success).toBe(false);
  const unknown = source(body().replace('w:w="1200"', 'w:w="unknown"'));
  expect(
    planInsertedTableColumn(unknown, { ...target(unknown), properties: { widthTwips: 1800 } })
      .success
  ).toBe(false);
  expect(writeOoxmlPackage(input.pkg)).toEqual(before);
});

test('new payload IDs cannot collide with orphan records or native references outside diagnostic bounds', () => {
  const input = source(),
    next = richTarget(input),
    stored = withPayloadRecord(input.pkg, next.payloads![0]!.payload, 'Orphan');
  if (!stored.ok) throw new Error(stored.reason);
  expect(planInsertedTableColumn({ ...input, pkg: stored.pkg }, next).success).toBe(false);
  const missing = source(
    body() +
      Array.from(
        { length: 70 },
        (_, index) =>
          `<w:sdt><w:sdtPr><w:tag w:val="scio:1:missing-${index}"/></w:sdtPr><w:sdtContent>${p('Missing')}</w:sdtContent></w:sdt>`
      ).join('') +
      `<w:sdt><w:sdtPr><w:tag w:val="scio:1:column-activity"/></w:sdtPr><w:sdtContent>${p('Collision')}</w:sdtContent></w:sdt>`
  );
  expect(buildBlockIndex(missing.pkg).omittedDiagnostics).toBeGreaterThan(0);
  expect(planInsertedTableColumn(missing, richTarget(missing)).success).toBe(false);
});

test('getters, cycles, depth, bytes, payload bounds, and forged source anchors fail before side effects', () => {
  const input = source(),
    base = target(input),
    before = writeOoxmlPackage(input.pkg);
  let calls = 0;
  const getter = Object.defineProperty({ ...base }, 'cells', {
    enumerable: true,
    get() {
      calls++;
      return [];
    },
  });
  const cyclic: unknown[] = [];
  cyclic.push(cyclic);
  let deep: unknown = { type: 'text', value: 'end' };
  for (let index = 0; index < 70; index++) deep = { type: 'blockquote', children: [deep] };
  const rich = richTarget(input);
  const invalid: unknown[] = [
    getter,
    { ...base, cells: cyclic },
    { ...base, cells: [{ ...base.cells[0], blocks: [deep] }, base.cells[1]] },
    {
      ...base,
      cells: [
        {
          ...base.cells[0],
          blocks: [
            {
              type: 'paragraph',
              children: [{ type: 'text', value: 'x'.repeat(16 * 1024 * 1024) }],
            },
          ],
        },
        base.cells[1],
      ],
    },
    {
      ...rich,
      payloads: [
        {
          ...rich.payloads![0],
          payload: { ...rich.payloads![0]!.payload, extra: 'x'.repeat(2 * 1024 * 1024) },
        },
        rich.payloads![1],
      ],
    },
  ];
  for (const value of invalid)
    expect(planInsertedTableColumn(input, value as InsertedTableColumnTarget).success).toBe(false);
  expect(planInsertedTableColumn({ ...input, partName: '/missing.xml' }, base).success).toBe(false);
  expect(planInsertedTableColumn({ ...input, tableId: 'missing' }, base).success).toBe(false);
  expect(
    planInsertedTableColumn(
      Object.defineProperty({ ...input }, 'tableId', {
        get() {
          calls++;
          return input.tableId;
        },
      }),
      base
    ).success
  ).toBe(false);
  expect(calls).toBe(0);
  expect(writeOoxmlPackage(input.pkg)).toEqual(before);
});

test('plan decoder validates row anchors, exact rich payload bindings, and explicit width shape', () => {
  const input = source(),
    planned = plan(input, richTarget(input)),
    edit = planned.operations[0]!;
  if (edit.kind !== 'insert-column') throw new Error('Missing rich column');
  for (const intent of [
    { ...edit.intent, previewCellId: 'forged' },
    { ...edit.intent, payloads: [] },
    { ...edit.intent, properties: { widthTwips: 0 } },
    { ...edit.intent, cells: [edit.intent.cells[0], edit.intent.cells[0]] },
  ])
    expect(createEditPlan(revision, [{ ...edit, intent } as never]).success).toBe(false);
});

test('rich column planning preserves live package bytes, revision, selection, and undo/redo history', () => {
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
      part = partOf(pkg),
      current = session.packageRevision();
    const input: InsertedTableColumnSource = {
      revision: { ...revision, revision: current },
      pkg,
      partName: part.name,
      tableId: tables(pkg)[0]!.id,
    };
    const bytes = writeOoxmlPackage(pkg),
      selection = surface.state().selection,
      history = [session.canUndo(), session.canRedo()];
    expect(planInsertedTableColumn(input, richTarget(input)).success).toBe(true);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(session.packageRevision()).toBe(current);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
  } finally {
    editor.destroy();
  }
});
