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
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import {
  materializeInsertedRow,
  namedChild,
  nodeChildren,
  rowCells,
  tableRows,
} from '../../test/support/inserted-row-oracle';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import { MDX_BLOCK } from '../projection/mdx';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import {
  planInsertedTableRow,
  type InsertedTableRowSource,
  type InsertedTableRowTarget,
} from './inserted-table-rows';
import { createEditPlan, decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';

const revision = { documentId: 'rich-rows', epoch: 'load-1', revision: 8 };
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, properties = '') =>
  `<w:tc><w:tcPr><w:shd w:fill="EEEEEE"/>${properties}</w:tcPr>${p(text)}</w:tc>`;
const row = (left: string, right: string) =>
  `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="360"/></w:trPr>${cell(left)}${cell(right)}</w:tr>`;
const table = (rows: string) =>
  `<w:tbl><w:tblPr><w:tblCaption w:val="Keep"/></w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>${rows}</w:tbl>`;
const body = () =>
  p('Before') +
  table(row('One', 'Two') + row('Keep left', 'Keep right')) +
  p('Between') +
  table(row('Unrelated', 'Table')) +
  p('After');
function source(bodyXml = body()): InsertedTableRowSource {
  const loaded = readOoxmlPackage(docxFixture(bodyXml));
  if (!loaded.ok) throw new Error(loaded.reason);
  const pkg = loaded.package,
    part = pkg.parts.get(pkg.mainDocumentPart)!;
  const tableId = nodeChildren(bodyStoryRoot(part)!).find((node) => node.kind === 'table')!.id;
  return { revision, pkg, partName: part.name, tableId };
}
const partOf = (pkg: OoxmlPackage) => pkg.parts.get(pkg.mainDocumentPart)!;
const ast = (markdown: string): TargetAstNode[] => {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  return JSON.parse(JSON.stringify(parsed.target.normalizedAst.children));
};
function target(input: InsertedTableRowSource): InsertedTableRowTarget {
  return {
    rowId: tableRows(partOf(input.pkg), input.tableId)[0]!.id,
    where: 'below',
    cells: [[], []],
  };
}
function richTarget(input: InsertedTableRowSource): InsertedTableRowTarget {
  const first = ast(
    '**Bold** [Link](https://example.com/a?b=1&c=2)  \nNext 😀\n\nSecond paragraph.\n\n<activity name="New" duration="5">\n\nApplication child.\n\n</activity>'
  );
  const widget = ast('<Widget state={{mode: "safe"}}>\n\n<Child />\n\n</Widget>')[0]!;
  return {
    ...target(input),
    properties: { heightTwips: 500 },
    cells: [
      first,
      [
        {
          type: 'table',
          align: [null],
          children: [
            {
              type: 'tableRow',
              children: [
                {
                  type: 'tableCell',
                  children: [
                    { type: 'strong', children: [{ type: 'text', value: 'Nested table' }] },
                  ],
                },
              ],
            },
          ],
        },
        ...ast('Another paragraph.'),
        widget,
      ],
    ],
    payloads: [
      {
        path: [0, 2],
        label: 'New activity',
        payload: {
          id: 'new-activity',
          type: 'lesson_activity',
          version: 1,
          attributes: { name: 'New', duration: 5, privateState: { keep: true } },
          unknownRoot: ['retain'],
        },
      },
      {
        path: [1, 2],
        label: 'Widget fallback',
        payload: {
          id: 'new-widget',
          type: MDX_BLOCK,
          version: 1,
          attributes: { host: 'react' },
          mdast: widget,
          vendor: { keep: true },
        },
      },
    ],
  };
}
function plan(input: InsertedTableRowSource, next: InsertedTableRowTarget): CanonicalEditPlan {
  const result = planInsertedTableRow(input, next);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
function descendants(node: OoxmlNode): OoxmlNode[] {
  return [node, ...nodeChildren(node).flatMap(descendants)];
}
function content(part: OoxmlPart, node: OoxmlNode) {
  return descendants(node)
    .filter((entry) => entry.kind === 'paragraph')
    .map((entry) => paragraphTextOf(part, entry.id));
}
function reopen(pkg: OoxmlPackage): OoxmlPackage {
  const result = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!result.ok) throw new Error(result.reason);
  return result.package;
}

test('rich row preserves marks, links, breaks, paragraphs, nested tables, opaque React data, and explicit height through native insertion and save/reopen', () => {
  const input = source(),
    original = partOf(input.pkg),
    next = richTarget(input);
  const oldRows = tableRows(original, input.tableId),
    oldCells = oldRows.flatMap(rowCells);
  const unrelated = nodeChildren(bodyStoryRoot(original)!).filter(
    (node) => node.kind === 'table'
  )[1]!;
  const before = writeOoxmlPackage(input.pkg),
    planned = plan(input, next);
  expect(writeOoxmlPackage(input.pkg)).toEqual(before);
  expect(planned.operations).toEqual([
    {
      kind: 'insert-row',
      partName: original.name,
      intent: {
        tableId: input.tableId,
        ...next,
        properties: next.properties!,
        payloads: next.payloads!,
      },
    },
  ]);
  const decoded = decodeEditPlan(serializeEditPlan(planned));
  if (!decoded.success) throw new Error(decoded.error);
  const result = materializeInsertedRow(input.pkg, decoded.plan);
  const part = partOf(result.pkg),
    inserted = findNode(part, result.rowId)!;
  expect(tableRows(part, input.tableId)).toHaveLength(3);
  expect(content(part, inserted).filter(Boolean)).toEqual([
    'Bold Link\nNext 😀',
    'Second paragraph.',
    'Application child.',
    'Nested table',
    'Another paragraph.',
    'Widget fallback',
  ]);
  for (const node of [...oldRows, ...oldCells, unrelated])
    expect(findNode(part, node.id) === node).toBe(true);
  const xml = serializeOoxmlPart(part);
  expect(xml).toContain('<w:b');
  expect(xml).toContain('<w:br');
  expect(xml).toContain('w:hRule="exact"');
  const height = namedChild(namedChild(inserted, 'trPr')!, 'trHeight')!;
  expect(
    height.kind !== 'textValue' && height.attributes.find((attr) => attr.localName === 'val')?.value
  ).toBe('500');
  expect(namedChild(namedChild(inserted, 'trPr')!, 'cantSplit')).toBeDefined();
  expect(descendants(inserted).filter((node) => node.kind === 'table')).toHaveLength(1);
  expect(
    result.pkg.externalTargets.filter((link) => link.rawTarget === 'https://example.com/a?b=1&c=2')
  ).toHaveLength(1);
  for (const pkg of [result.pkg, reopen(result.pkg)]) {
    const indexed = buildBlockIndex(pkg);
    expect(indexed.diagnostics).toEqual([]);
    expect(indexed.byId.get('new-activity')?.payload).toEqual(next.payloads![0]!.payload);
    expect(indexed.byId.get('new-widget')?.payload).toEqual(next.payloads![1]!.payload);
    const savedPart = partOf(pkg),
      savedTables = nodeChildren(bodyStoryRoot(savedPart)!).filter((node) => node.kind === 'table');
    const saved = tableRows(savedPart, savedTables[0]!.id)[1]!;
    expect(content(savedPart, saved).filter(Boolean)).toEqual(
      content(part, inserted).filter(Boolean)
    );
    expect(serializeOoxmlPart(savedPart)).toContain('w:hRule="exact"');
    const savedHeight = namedChild(namedChild(saved, 'trPr')!, 'trHeight')!;
    expect(
      savedHeight.kind !== 'textValue' &&
        savedHeight.attributes.find((attr) => attr.localName === 'val')?.value
    ).toBe('500');
    expect(descendants(saved).filter((node) => node.kind === 'table')).toHaveLength(1);
    expect(
      pkg.externalTargets.some((link) => link.rawTarget === 'https://example.com/a?b=1&c=2')
    ).toBe(true);
    const links = descendants(saved).filter(
      (node) => node.kind !== 'textValue' && node.localName === 'hyperlink'
    );
    expect(links).toHaveLength(1);
    const link = links[0]!;
    const relId =
      link.kind !== 'textValue' && link.attributes.find((attr) => attr.localName === 'id')?.value;
    expect(pkg.externalTargets.find((entry) => entry.id === relId)?.rawTarget).toBe(
      'https://example.com/a?b=1&c=2'
    );
    const boldRun = descendants(saved).find(
      (node) =>
        node.kind === 'run' &&
        descendants(node).some((entry) => entry.kind === 'textValue' && entry.value === 'Bold')
    )!;
    expect(namedChild(namedChild(boldRun, 'rPr')!, 'b')).toBeDefined();
    const other = savedTables[1]!;
    expect(content(savedPart, other)).toEqual(['Unrelated', 'Table']);
    expect(canonicalOoxmlFingerprint(other)).toBe(canonicalOoxmlFingerprint(unrelated));
    const savedRows = tableRows(savedPart, savedTables[0]!.id);
    expect(canonicalOoxmlFingerprint(savedRows[0]!)).toBe(canonicalOoxmlFingerprint(oldRows[0]!));
    expect(canonicalOoxmlFingerprint(savedRows[2]!)).toBe(canonicalOoxmlFingerprint(oldRows[1]!));
    for (const [name, beforePart] of input.pkg.parts)
      if (name !== input.partName && !name.endsWith('.rels') && name !== '/[Content_Types].xml')
        expect(canonicalOoxmlFingerprint(pkg.parts.get(name)!)).toEqual(
          canonicalOoxmlFingerprint(beforePart)
        );
  }
  expect(writeOoxmlPackage(input.pkg)).toEqual(before);
});

test.each(['above', 'below'] as const)(
  'empty and Unicode rich cells insert %s with native row inheritance',
  (where) => {
    const input = source(),
      next = { ...target(input), where, cells: [[], ast('Žluťoučký 😀')] };
    const planned = plan(input, next),
      edit = planned.operations[0]!;
    expect(edit.kind === 'insert-row' && edit.intent.properties).toEqual({});
    expect(edit.kind === 'insert-row' && edit.intent.payloads).toEqual([]);
    const result = materializeInsertedRow(input.pkg, planned),
      part = partOf(result.pkg);
    const inserted = findNode(part, result.rowId)!;
    expect(content(part, inserted).filter(Boolean)).toEqual(['Žluťoučký 😀']);
    expect(serializeOoxmlPart(part)).toContain('w:shd w:fill="EEEEEE"');
    expect(namedChild(namedChild(inserted, 'trPr')!, 'cantSplit')).toBeDefined();
  }
);

test('plan serialization contains only canonical anchors, new fragments, properties, and detached owned payloads', () => {
  const input = source(),
    next = richTarget(input),
    planned = plan(input, next);
  const serialized = serializeEditPlan(planned),
    edit = planned.operations[0]!;
  expect(decodeEditPlan(serialized)).toEqual({ success: true, plan: planned });
  expect(serialized).not.toMatch(
    /createdRowIds|newRowId|cellId|paragraphId|textValue|resultPart|controlId/
  );
  expect(serialized).not.toContain('Keep left');
  expect(serialized).not.toContain('Unrelated');
  expect(edit.kind === 'insert-row' && Object.isFrozen(edit.intent.cells[0]![0])).toBe(true);
  expect(Object.isFrozen(next.cells)).toBe(false);
  (next.payloads![0]!.payload.attributes as Record<string, unknown>).name = 'Caller mutation';
  expect(serializeEditPlan(planned)).toBe(serialized);
});

test('runtime resolution uses native row/cell IDs after unrelated allocation, without preview IDs', () => {
  const input = source(),
    next = { ...target(input), cells: [ast('Runtime content'), []] },
    planned = plan(input, next);
  const part = partOf(input.pkg),
    before = tableRows(part, input.tableId).map((node) => node.id);
  const allocated = applyTreeOp(part, {
    op: 'insertTableRow',
    tableId: input.tableId,
    rowId: next.rowId,
    where: 'above',
  });
  if (!allocated.ok) throw new Error(allocated.reason);
  // This immutable oracle bypasses revision checks deliberately to test allocation only.
  // A01 must refuse stale revisions before a live transaction.
  const result = materializeInsertedRow(withPart(input.pkg, allocated.part), planned),
    actual = partOf(result.pkg);
  expect(tableRows(actual, input.tableId)).toHaveLength(4);
  expect(content(actual, findNode(actual, result.rowId)!).filter(Boolean)).toEqual([
    'Runtime content',
  ]);
  expect(before.every((id) => !!findNode(actual, id))).toBe(true);
  expect(new Set(descendants(actual.root).map((node) => node.id)).size).toBe(
    descendants(actual.root).length
  );
});

test('native grid spans determine physical cell count; hidden vertical continuations accept only omitted content', () => {
  const merged = source(
    table(
      `<w:tr>${cell('A', '<w:vMerge w:val="restart"/>')}${cell('B')}</w:tr>` +
        `<w:tr>${cell('', '<w:vMerge/>')}${cell('C')}</w:tr>`
    )
  );
  const base = target(merged);
  expect(planInsertedTableRow(merged, { ...base, cells: [[], ast('Visible')] }).success).toBe(true);
  expect(planInsertedTableRow(merged, { ...base, cells: [ast('Hidden loss'), []] })).toMatchObject({
    success: false,
    code: 'unsupported-operation',
  });
  const spanned = source(table(`<w:tr>${cell('Spanned', '<w:gridSpan w:val="2"/>')}</w:tr>`));
  expect(
    planInsertedTableRow(spanned, { ...target(spanned), cells: [ast('One physical cell')] }).success
  ).toBe(true);
  expect(planInsertedTableRow(spanned, target(spanned)).success).toBe(false);
});

test('payload bindings reject missing, ambiguous, mismatched, and unreachable owned AST data', () => {
  const input = source(),
    next = richTarget(input),
    activity = next.payloads![0]!,
    widget = next.payloads![1]!;
  const invalid = [
    { ...next, payloads: [] },
    { ...next, payloads: [activity] },
    { ...next, payloads: [activity, activity, widget] },
    { ...next, payloads: [{ ...activity, path: [0, 0] }, widget] },
    { ...next, payloads: [{ ...activity, path: [0, 2, 5] }, widget] },
    { ...next, payloads: [{ ...activity, path: [0, -1] }, widget] },
    {
      ...next,
      payloads: [
        {
          ...activity,
          payload: { ...activity.payload, attributes: { name: 'Wrong', duration: 5 } },
        },
        widget,
      ],
    },
    {
      ...next,
      payloads: [
        activity,
        { ...widget, payload: { ...widget.payload, mdast: ast('<Different />')[0]! } },
      ],
    },
    {
      ...next,
      payloads: [activity, { ...widget, payload: { ...widget.payload, id: activity.payload.id } }],
    },
    { ...next, payloads: [{ ...activity, label: '\u0000' }, widget] },
    {
      ...next,
      payloads: [
        activity,
        widget,
        {
          ...activity,
          path: [1, 2, 0],
          payload: { ...activity.payload, id: 'nested-under-opaque' },
        },
      ],
    },
  ];
  for (const value of invalid) expect(planInsertedTableRow(input, value).success).toBe(false);
  const before = writeOoxmlPackage(input.pkg);
  expect(plan(input, next).operations).toHaveLength(1);
  expect(writeOoxmlPackage(input.pkg)).toEqual(before);
});

test('new payload IDs cannot collide with orphan records or unresolved native control tags', () => {
  const input = source(),
    next = richTarget(input),
    existing = next.payloads![0]!.payload;
  const stored = withPayloadRecord(input.pkg, existing, 'Orphan');
  if (!stored.ok) throw new Error(stored.reason);
  expect(readPayloadStore(stored.pkg).status).toBe('ready');
  expect(planInsertedTableRow({ ...input, pkg: stored.pkg }, next).success).toBe(false);
  const missing = source(
    body() +
      '<w:sdt><w:sdtPr><w:tag w:val="scio:1:new-activity"/></w:sdtPr><w:sdtContent>' +
      p('Missing payload') +
      '</w:sdtContent></w:sdt>'
  );
  expect(buildBlockIndex(missing.pkg).diagnostics.some((item) => item.id === 'new-activity')).toBe(
    true
  );
  expect(planInsertedTableRow(missing, richTarget(missing)).success).toBe(false);
  const manyMissing = source(
    body() +
      Array.from(
        { length: 70 },
        (_, index) =>
          `<w:sdt><w:sdtPr><w:tag w:val="scio:1:missing-${index}"/></w:sdtPr><w:sdtContent>${p('Missing')}</w:sdtContent></w:sdt>`
      ).join('') +
      '<w:sdt><w:sdtPr><w:tag w:val="scio:1:new-activity"/></w:sdtPr><w:sdtContent>' +
      p('Collision') +
      '</w:sdtContent></w:sdt>'
  );
  expect(buildBlockIndex(manyMissing.pkg).omittedDiagnostics).toBeGreaterThan(0);
  expect(planInsertedTableRow(manyMissing, richTarget(manyMissing)).success).toBe(false);
});

test('invalid source anchors, explicit properties, AST shapes, getters, cycles, and bounds refuse without writes', () => {
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
  for (let depth = 0; depth < 70; depth++) deep = { type: 'blockquote', children: [deep] };
  const invalid: unknown[] = [
    null,
    [],
    {},
    getter,
    { ...base, cells: cyclic },
    { ...base, cells: [[deep], []] },
    { ...base, where: 'after' },
    { ...base, where: { toString: 'inert' } },
    { ...base, rowId: 'missing' },
    { ...base, cells: [[]] },
    { ...base, cells: [[{ type: 'text', value: 'Needs paragraph' }], []] },
    { ...base, cells: [[{ type: 'paragraph', kind: 'paragraph', children: [] }], []] },
    { ...base, properties: null },
    { ...base, payloads: null },
    { ...base, properties: { heightTwips: 19 } },
    { ...base, properties: { heightTwips: 31681 } },
    { ...base, properties: { heightTwips: 500.5 } },
    { ...base, properties: { size: 20 } },
    { ...base, previewRowId: 'forged' },
    {
      ...base,
      cells: [
        [{ type: 'paragraph', children: [{ type: 'text', value: 'x'.repeat(16 * 1024 * 1024) }] }],
        [],
      ],
    },
  ];
  for (const value of invalid)
    expect(planInsertedTableRow(input, value as InsertedTableRowTarget).success).toBe(false);
  expect(planInsertedTableRow({ ...input, tableId: 'missing' }, base).success).toBe(false);
  expect(planInsertedTableRow({ ...input, partName: '/missing.xml' }, base).success).toBe(false);
  expect(
    planInsertedTableRow(
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

test('plan decoder enforces rich row schema and exact binding consistency', () => {
  const input = source(),
    planned = plan(input, richTarget(input)),
    edit = planned.operations[0]!;
  if (edit.kind !== 'insert-row') throw new Error('Missing rich row');
  for (const intent of [
    { ...edit.intent, newRowId: 'native-preview' },
    { ...edit.intent, payloads: [] },
    { ...edit.intent, properties: { heightTwips: 0 } },
    { ...edit.intent, payloads: [{ ...edit.intent.payloads[0], path: [0, 0] }] },
  ])
    expect(createEditPlan(revision, [{ ...edit, intent } as never]).success).toBe(false);
});

test('planning keeps live revision, canonical bytes, selection, and undo/redo history unchanged', () => {
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
    const input: InsertedTableRowSource = {
      revision: { ...revision, revision: current },
      pkg,
      partName: part.name,
      tableId: nodeChildren(bodyStoryRoot(part)!).find((node) => node.kind === 'table')!.id,
    };
    const before = writeOoxmlPackage(pkg),
      selection = surface.state().selection;
    const history = [session.canUndo(), session.canRedo()];
    expect(planInsertedTableRow(input, richTarget(input)).success).toBe(true);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
    expect(surface.state().selection).toEqual(selection);
    expect(session.packageRevision()).toBe(current);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
  } finally {
    editor.destroy();
  }
});
