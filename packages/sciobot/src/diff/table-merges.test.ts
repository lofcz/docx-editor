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
import {
  namedChild,
  nodeChildren,
  rowCells,
  tableRows,
} from '../../test/support/native-table-oracle';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { nativeGridColumns } from './inserted-column-native';
import { decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';
import { planTableMerge, type TableMergeSource, type TableMergeTarget } from './table-merges';

const revision = { documentId: 'merges', epoch: 'load-1', revision: 6 };
const p = (text: string) =>
  `<w:p><w:pPr><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width = 2400, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:shd w:fill="EEEEEE"/><keep:extra xmlns:keep="urn:keep" keep:value="retain"/>${extra}</w:tcPr>${p(text)}</w:tc>`;
const row = (a: string, b: string, c: string) =>
  `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(a, 1200)}${cell(b)}${cell(c, 3600)}</w:tr>`;
const grid =
  '<w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid>';
const table = (rows: string, columns = grid) =>
  `<w:tbl><w:tblPr><w:tblCaption w:val="Keep table metadata"/></w:tblPr>${columns}${rows}</w:tbl>`;
const body = () =>
  p('Before') +
  table(row('A1', 'B1', 'C1') + row('A2', 'B2', 'C2') + row('A3', 'B3', 'C3')) +
  p('Between') +
  table(row('Other A', 'Other B', 'Other C')) +
  p('After');
const tables = (part: OoxmlPart) =>
  nodeChildren(bodyStoryRoot(part)!).filter((node) => node.kind === 'table');
const attr = (node: OoxmlNode, name: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find((entry) => entry.localName === name)?.value;
function source(part = bodyPart(body()), tableIndex = 0): TableMergeSource {
  return { revision, part, tableId: tables(part)[tableIndex]!.id };
}
const cells = (input: TableMergeSource) => tableRows(input.part, input.tableId).map(rowCells);
function descendants(node: OoxmlNode): OoxmlNode[] {
  return [node, ...nodeChildren(node).flatMap(descendants)];
}
const text = (part: OoxmlPart, node: OoxmlNode) =>
  descendants(node)
    .filter((node) => node.kind === 'paragraph')
    .map((node) => paragraphTextOf(part, node.id))
    .filter(Boolean);
const values = (part: OoxmlPart, tableId: string) =>
  tableRows(part, tableId).map((row) => rowCells(row).map((cell) => text(part, cell)));
function plan(input: TableMergeSource, target: TableMergeTarget): CanonicalEditPlan {
  const result = planTableMerge(input, target);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
/** Pure native oracle; A01/R10 own guarded mutations, native review, and collaboration. */
function apply(part: OoxmlPart, planned: CanonicalEditPlan): OoxmlPart {
  let current = part;
  for (const edit of planned.operations) {
    if (edit.kind !== 'table') throw new Error('Expected native table edit');
    const result = applyTreeOp(current, edit.operation);
    if (!result.ok) throw new Error(result.reason);
    current = result.part;
  }
  return current;
}

test('horizontal merge preserves reading-order text, paragraph identities, marks, imported XML, and every unselected cell', () => {
  const input = source(),
    before = canonicalOoxmlFingerprint(input.part),
    oldCells = cells(input),
    owner = oldCells[0]![0]!,
    removed = oldCells[0]![1]!;
  const paragraphs = [owner, removed].map(
    (cell) => nodeChildren(cell).find((node) => node.kind === 'paragraph')!
  );
  const planned = plan(input, { kind: 'merge', cellIds: [removed.id, owner.id] });
  expect(planned.operations).toEqual([
    {
      kind: 'table',
      partName: input.part.name,
      operation: { op: 'mergeTableCells', tableId: input.tableId, cellIds: [removed.id, owner.id] },
    },
  ]);
  const result = apply(input.part, planned),
    merged = findNode(result, owner.id)!;
  expect(values(result, input.tableId)).toEqual([
    [['A1', 'B1'], ['C1']],
    [['A2'], ['B2'], ['C2']],
    [['A3'], ['B3'], ['C3']],
  ]);
  expect(attr(namedChild(namedChild(merged, 'tcPr')!, 'gridSpan')!, 'val')).toBe('2');
  expect(attr(namedChild(namedChild(merged, 'tcPr')!, 'tcW')!, 'w')).toBe('3600');
  for (const paragraph of paragraphs)
    expect(findNode(result, paragraph.id) === paragraph).toBe(true);
  for (const paragraph of paragraphs) {
    const run = nodeChildren(paragraph).find((node) => node.kind === 'run')!;
    expect(namedChild(namedChild(run, 'rPr')!, 'b')).toBeDefined();
  }
  for (const cell of oldCells
    .flat()
    .filter((cell) => cell.id !== owner.id && cell.id !== removed.id))
    expect(findNode(result, cell.id) === cell).toBe(true);
  expect(findNode(result, removed.id)).toBeNull();
  expect(serializeOoxmlPart(result)).toContain('keep:value="retain"');
  for (const node of nativeGridColumns(input.part, input.tableId))
    expect(findNode(result, node.id) === node).toBe(true);
  expect(findNode(result, tables(input.part)[1]!.id) === tables(input.part)[1]).toBe(true);
  expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
});

test('vertical merge moves all nonempty content into its owner and retains native continuation geometry', () => {
  const input = source(),
    original = cells(input),
    owner = original[0]![0]!,
    lower = original[1]![0]!;
  const moved = nodeChildren(lower).find((node) => node.kind === 'paragraph')!;
  const result = apply(input.part, plan(input, { kind: 'merge', cellIds: [lower.id, owner.id] }));
  expect(values(result, input.tableId)).toEqual([
    [['A1', 'A2'], ['B1'], ['C1']],
    [[], ['B2'], ['C2']],
    [['A3'], ['B3'], ['C3']],
  ]);
  const after = tableRows(result, input.tableId).map(rowCells);
  expect(attr(namedChild(namedChild(after[0]![0]!, 'tcPr')!, 'vMerge')!, 'val')).toBe('restart');
  expect(namedChild(namedChild(after[1]![0]!, 'tcPr')!, 'vMerge')).toBeDefined();
  expect(after[0]![0]!.id).toBe(owner.id);
  expect(findNode(result, moved.id) === moved).toBe(true);
  expect(nodeChildren(after[0]![0]!).some((node) => node.id === moved.id)).toBe(true);
  for (const cell of original.flat().filter((cell) => cell.id !== owner.id && cell.id !== lower.id))
    expect(findNode(result, cell.id) === cell).toBe(true);
});

test('rectangle endpoints merge every intended cell, then existing-grid split restores original unequal slots without redistributing content', () => {
  const input = source(),
    original = cells(input),
    ownerId = original[0]![0]!.id;
  const merged = apply(
    input.part,
    plan(input, { kind: 'merge', cellIds: [ownerId, original[1]![1]!.id] })
  );
  expect(text(merged, findNode(merged, ownerId)!)).toEqual(['A1', 'B1', 'A2', 'B2']);
  const continuationId = rowCells(tableRows(merged, input.tableId)[1]!)[0]!.id;
  const splitPlan = plan(
    { ...input, part: merged },
    {
      kind: 'split',
      cellId: continuationId,
      rows: 2,
      columns: 2,
      columnDistribution: 'existing-grid',
    }
  );
  const restored = apply(merged, splitPlan);
  expect(values(restored, input.tableId)).toEqual([
    [['A1', 'B1', 'A2', 'B2'], [], ['C1']],
    [[], [], ['C2']],
    [['A3'], ['B3'], ['C3']],
  ]);
  expect(nativeGridColumns(restored, input.tableId).map((node) => attr(node, 'w'))).toEqual([
    '1200',
    '2400',
    '3600',
  ]);
  expect(rowCells(tableRows(restored, input.tableId)[0]!)[0]!.id).toBe(ownerId);
  for (const cell of [original[0]![2]!, original[1]![2]!, ...original[2]!])
    expect(findNode(restored, cell.id) === cell).toBe(true);
  for (const node of nativeGridColumns(input.part, input.tableId))
    expect(findNode(restored, node.id) === node).toBe(true);
  expect(serializeEditPlan(splitPlan)).not.toMatch(
    /created|newCell|resultPart|"children"|textValue/
  );
});

test.each(['omitted', 'equal-width'] as const)(
  'explicit Word split retains native equal-width behavior; distribution=%s',
  (distribution) => {
    const input = source(),
      owner = cells(input)[0]![0]!,
      second = cells(input)[0]![1]!;
    const merged = apply(
      input.part,
      plan(input, { kind: 'merge', cellIds: [owner.id, second.id] })
    );
    const restored = apply(
      merged,
      plan(
        { ...input, part: merged },
        {
          kind: 'split',
          cellId: owner.id,
          rows: 1,
          columns: 2,
          ...(distribution === 'omitted' ? {} : { columnDistribution: distribution }),
        }
      )
    );
    expect(nativeGridColumns(restored, input.tableId).map((node) => attr(node, 'w'))).toEqual([
      '1200',
      '600',
      '1800',
      '3600',
    ]);
    expect(text(restored, findNode(restored, owner.id)!)).toEqual(['A1', 'B1']);
    expect(text(restored, findNode(restored, cells(input)[1]![1]!.id)!)).toEqual(['B2']);
  }
);

test('one-row split inserts native rows and preserves neighboring vertical spans and all original paragraph IDs', () => {
  const input = source(),
    originalRows = tableRows(input.part, input.tableId),
    originalCells = cells(input),
    selected = originalCells[0]![0]!;
  const originalParagraphs = originalCells
    .flat()
    .flatMap((cell) => nodeChildren(cell).filter((node) => node.kind === 'paragraph'));
  const planned = plan(input, {
    kind: 'split',
    cellId: selected.id,
    rows: 3,
    columns: 1,
    columnDistribution: 'existing-grid',
  });
  const result = apply(input.part, planned),
    after = tableRows(result, input.tableId);
  expect(after).toHaveLength(5);
  expect(values(result, input.tableId)).toEqual([
    [['A1'], ['B1'], ['C1']],
    [[], [], []],
    [[], [], []],
    [['A2'], ['B2'], ['C2']],
    [['A3'], ['B3'], ['C3']],
  ]);
  expect(after[3]!.id).toBe(originalRows[1]!.id);
  for (const paragraph of originalParagraphs)
    expect(findNode(result, paragraph.id) === paragraph).toBe(true);
  expect(namedChild(namedChild(rowCells(after[0]!)[1]!, 'tcPr')!, 'vMerge')).toBeDefined();
  expect(namedChild(namedChild(rowCells(after[1]!)[1]!, 'tcPr')!, 'vMerge')).toBeDefined();
  expect(serializeEditPlan(planned)).not.toMatch(/created|newRow|paragraphId|"children"/);
});

test.each(['blank head', 'all blank'])(
  'merging %s retains only intended native content and a valid paragraph',
  (name) => {
    const input = source(bodyPart(table(row('', name === 'all blank' ? '' : 'Keep', 'C'))));
    const before = cells(input)[0]!,
      result = apply(
        input.part,
        plan(input, { kind: 'merge', cellIds: before.slice(0, 2).map((cell) => cell.id) })
      );
    const owner = findNode(result, before[0]!.id)!;
    expect(text(result, owner)).toEqual(name === 'all blank' ? [] : ['Keep']);
    expect(nodeChildren(owner).some((node) => node.kind === 'paragraph')).toBe(true);
  }
);

test('overlapping visual spans refuse rather than silently growing the requested rectangle', () => {
  const input = source(),
    beforeCells = cells(input);
  const merged = apply(
    input.part,
    plan(input, { kind: 'merge', cellIds: [beforeCells[0]![1]!.id, beforeCells[1]![1]!.id] })
  );
  const next = { ...input, part: merged },
    native = cells(next),
    before = canonicalOoxmlFingerprint(merged);
  const result = planTableMerge(next, {
    kind: 'merge',
    cellIds: [native[1]![0]!.id, native[2]![1]!.id],
  });
  expect(result).toMatchObject({ success: false, code: 'unsupported-operation' });
  expect(result.success ? result.plan : null).toBeNull();
  expect(canonicalOoxmlFingerprint(merged)).toBe(before);
});

test('merge and split retain owned application controls and payload parts through native DOCX save/reopen', () => {
  const loaded = readOoxmlPackage(
    docxFixture(
      body().replace(
        p('B1'),
        '<w:sdt><w:sdtPr><w:tag w:val="scio:1:owned"/></w:sdtPr><w:sdtContent>' +
          p('Owned activity') +
          '</w:sdtContent></w:sdt>'
      )
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const stored = withPayloadRecord(
    loaded.package,
    {
      id: 'owned',
      type: 'lesson_activity',
      version: 1,
      attributes: { name: 'Owned', duration: 5, asset: { id: 'keep' } },
      unknownRoot: ['retain'],
    },
    'Owned'
  );
  if (!stored.ok) throw new Error(stored.reason);
  const pkg = stored.pkg,
    original = pkg.parts.get(pkg.mainDocumentPart)!,
    input = source(original);
  const originalIndex = buildBlockIndex(pkg).byId.get('owned')!,
    originalControl = findNode(original, originalIndex.controlId)!;
  const selected = cells(input)[0]!,
    merged = apply(
      original,
      plan(input, { kind: 'merge', cellIds: selected.slice(0, 2).map((cell) => cell.id) })
    );
  const split = apply(
    merged,
    plan(
      { ...input, part: merged },
      {
        kind: 'split',
        cellId: selected[0]!.id,
        rows: 1,
        columns: 2,
        columnDistribution: 'existing-grid',
      }
    )
  );
  expect(findNode(split, originalControl.id) === originalControl).toBe(true);
  const output = withPart(pkg, split),
    reopened = readOoxmlPackage(writeOoxmlPackage(output));
  if (!reopened.ok) throw new Error(reopened.reason);
  for (const candidate of [output, reopened.package]) {
    expect(buildBlockIndex(candidate).diagnostics).toEqual([]);
    expect(buildBlockIndex(candidate).byId.get('owned')?.payload).toEqual(originalIndex.payload);
    expect(canonicalOoxmlFingerprint(candidate.parts.get(candidate.mainDocumentPart)!)).toBe(
      canonicalOoxmlFingerprint(split)
    );
    for (const [name, part] of pkg.parts)
      if (name !== pkg.mainDocumentPart)
        expect(canonicalOoxmlFingerprint(candidate.parts.get(name)!)).toBe(
          canonicalOoxmlFingerprint(part)
        );
  }
});

test('nested cell ownership cannot escape into another canonical table', () => {
  const nested = table(row('Nested A', 'Nested B', 'Nested C'));
  const input = source(
    bodyPart(
      table(`<w:tr><w:tc>${p('Outer')}${nested}</w:tc>${cell('Other')}${cell('Tail')}</w:tr>`)
    )
  );
  const outerCells = cells(input)[0]!,
    inner = nodeChildren(outerCells[0]!).find((node) => node.kind === 'table')!;
  const innerInput = { ...input, tableId: inner.id },
    innerCells = cells(innerInput)[0]!;
  expect(
    planTableMerge(input, { kind: 'merge', cellIds: innerCells.slice(0, 2).map((cell) => cell.id) })
      .success
  ).toBe(false);
  const result = apply(
    input.part,
    plan(innerInput, { kind: 'merge', cellIds: innerCells.slice(0, 2).map((cell) => cell.id) })
  );
  expect(values(result, inner.id)).toEqual([[['Nested A', 'Nested B'], ['Nested C']]]);
  expect(findNode(result, outerCells[1]!.id) === outerCells[1]).toBe(true);
});

test('plans detach caller arrays and contain only canonical anchors and requested parameters', () => {
  const input = source(),
    ids = cells(input)[0]!
      .slice(0, 2)
      .map((cell) => cell.id),
    planned = plan(input, { kind: 'merge', cellIds: ids }),
    raw = serializeEditPlan(planned);
  expect(decodeEditPlan(raw)).toEqual({ success: true, plan: planned });
  expect(Object.isFrozen(planned.operations[0])).toBe(true);
  expect(Object.isFrozen(ids)).toBe(false);
  ids.reverse();
  expect(serializeEditPlan(planned)).toBe(raw);
  expect(raw).not.toMatch(/textValue|"children"|resultPart|created|"author"/);
});

test('invalid requests, counts, accessors, cycles, and resource excess refuse before producing a partial plan', () => {
  const input = source(),
    native = cells(input),
    ids = native[0]!.slice(0, 2).map((cell) => cell.id),
    before = canonicalOoxmlFingerprint(input.part);
  let calls = 0;
  const getter = Object.defineProperty({ kind: 'merge' }, 'cellIds', {
    enumerable: true,
    get() {
      calls++;
      return ids;
    },
  });
  const cyclic: unknown[] = [];
  cyclic.push(cyclic);
  const invalid: unknown[] = [
    null,
    [],
    {},
    getter,
    { kind: 'merge', cellIds: cyclic },
    { kind: 'merge', cellIds: [] },
    { kind: 'merge', cellIds: [ids[0], ids[0]] },
    { kind: 'merge', cellIds: [ids[0]] },
    { kind: 'merge', cellIds: [ids[0], 'missing'] },
    { kind: 'merge', cellIds: [ids[0], null] },
    { kind: 'merge', cellIds: ids, row: 0 },
    { kind: 'merge', cellIds: ['x'.repeat(16 * 1024 * 1024)] },
    { kind: 'split', cellId: ids[0], rows: 0, columns: 2 },
    { kind: 'split', cellId: ids[0], rows: 1, columns: 1 },
    { kind: 'split', cellId: ids[0], rows: 64, columns: 2 },
    { kind: 'split', cellId: ids[0], rows: 1, columns: 2.5 },
    { kind: 'split', cellId: ids[0], rows: 1, columns: 2, columnDistribution: 'guess' },
    { kind: 'split', cellId: ids[0], rows: 1, columns: 2, columnDistribution: null },
    { kind: 'split', cellId: 'missing', rows: 1, columns: 2 },
    { kind: 'split', cellId: ids[0], rows: 1, columns: 2, previewIds: [] },
  ];
  for (const target of invalid)
    expect(planTableMerge(input, target as TableMergeTarget).success).toBe(false);
  expect(
    planTableMerge({ ...input, tableId: 'missing' }, { kind: 'merge', cellIds: ids }).success
  ).toBe(false);
  expect(
    planTableMerge(
      Object.defineProperty({ ...input }, 'tableId', {
        get() {
          calls++;
          return input.tableId;
        },
      }),
      { kind: 'merge', cellIds: ids }
    ).success
  ).toBe(false);
  expect(calls).toBe(0);
  expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
});

test('locked, bound, and pending native cell revisions refuse before any merge/split plan returns', () => {
  for (const restriction of [
    '<w:lock w:val="sdtContentLocked"/>',
    '<w:dataBinding w:storeItemID="{00000000-0000-0000-0000-000000000001}" w:xpath="/item"/>',
  ]) {
    const part = bodyPart(
      `<w:sdt><w:sdtPr>${restriction}</w:sdtPr><w:sdtContent>${table(row('A', 'B', 'C'))}</w:sdtContent></w:sdt>`
    );
    const tableId = descendants(part.root).find((node) => node.kind === 'table')!.id;
    const input = { revision, part, tableId },
      selected = cells(input)[0]!,
      before = canonicalOoxmlFingerprint(part);
    expect(
      planTableMerge(input, { kind: 'merge', cellIds: selected.slice(0, 2).map((cell) => cell.id) })
        .success
    ).toBe(false);
    expect(
      planTableMerge(input, { kind: 'split', cellId: selected[0]!.id, rows: 1, columns: 2 }).success
    ).toBe(false);
    expect(canonicalOoxmlFingerprint(part)).toBe(before);
  }
  const input = source(
    bodyPart(
      table(
        `<w:tr>${cell('Tracked', 1200, '<w:cellIns w:id="7" w:author="Alice"/>')}${cell('B')}${cell('C', 3600)}</w:tr>`
      )
    )
  );
  const selected = cells(input)[0]!;
  expect(
    planTableMerge(input, { kind: 'merge', cellIds: selected.slice(0, 2).map((cell) => cell.id) })
      .success
  ).toBe(false);
  expect(
    planTableMerge(input, { kind: 'split', cellId: selected[0]!.id, rows: 1, columns: 2 }).success
  ).toBe(false);
});

test('merge/split planning preserves live bytes, revision, selection, and existing history', () => {
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
      selected = cells(input)[0]!;
    const bytes = writeOoxmlPackage(session.currentPackage()),
      selection = surface.state().selection,
      history = [session.canUndo(), session.canRedo()];
    expect(
      planTableMerge(input, { kind: 'merge', cellIds: selected.slice(0, 2).map((cell) => cell.id) })
        .success
    ).toBe(true);
    expect(
      planTableMerge(input, { kind: 'split', cellId: selected[1]!.id, rows: 1, columns: 2 }).success
    ).toBe(true);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(session.packageRevision()).toBe(current);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
  } finally {
    editor.destroy();
  }
});
