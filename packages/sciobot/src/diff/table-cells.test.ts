import { expect, test } from 'bun:test';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  AUTHORABLE_RUN_PROPERTIES,
  applyTreeOp,
  authoredProperties,
  canonicalOoxmlFingerprint,
  findNode,
  openHeadlessDocument,
  paragraphTextOf,
  propertyContainer,
  runsCovering,
  type OoxmlNode,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import reference from '../__tests__/fixtures/ai-pairs.json';
import { docxFixture } from '../../test/support/document';
import { projectTables } from '../projection/tables';
import { directParagraphTextOps } from './paragraph-text';
import { decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';
import { planTableCellEdits, type CellParagraphPatch, type TableCellSource } from './table-cells';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string, properties = '') =>
  `<w:tc><w:tcPr>${properties}<w:shd w:fill="EEEEEE"/></w:tcPr>${content}</w:tc>`;
const row = (...cells: string[]) => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cells.join('')}</w:tr>`;
const table = (rows: string) =>
  `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders><w:bottom w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${rows}</w:tbl>`;
const standard = () =>
  table(row(cell(p('A')), cell(p('B'))) + row(cell(p('One')), cell(p('Two')))) + p('Tail');
async function source(body = standard(), wanted = 'Two') {
  const opened = openHeadlessDocument(docxFixture(body));
  if (!opened.ok) throw new Error(opened.reason);
  const view = opened.view,
    part = view.part();
  const revision = { documentId: 'cell-test', epoch: 'mount-1', revision: view.packageRevision() };
  const exported = openDocumentForExport(view);
  if (!exported.ok) throw new Error(exported.reason);
  try {
    const result = projectTables(await exported.session.layout(), revision);
    if (!result.success) throw new Error(result.error);
    const selected = result.projection.cells.find((record) =>
      record.paragraphIds.some((id) => paragraphTextOf(part, id) === wanted)
    );
    if (!selected) throw new Error('Missing test cell');
    const input: TableCellSource = {
      revision,
      part,
      projection: result.projection,
      tableId: selected.tableId,
      rowId: selected.rowId,
      cellId: selected.id,
      displayMode: 'proposed',
    };
    return { input, selected };
  } finally {
    exported.session.dispose();
  }
}
function plan(input: TableCellSource, patches: readonly CellParagraphPatch[]) {
  const result = planTableCellEdits(input, patches);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
/** Pure native oracle. A01/R07 own guarded production execution and native review attribution. */
function apply(part: OoxmlPart, planned: CanonicalEditPlan): OoxmlPart {
  let current = part;
  for (const edit of planned.operations) {
    let ops: readonly TreeDocOp[];
    if (edit.kind === 'text') ops = directParagraphTextOps(edit);
    else if (edit.kind === 'format' && edit.target.kind === 'run') {
      const target = edit.target;
      const run = runsCovering(current, target.paragraphId, target.start, target.end)[0]!;
      const bag = new Map(
        authoredProperties(
          propertyContainer(run, 'runProperties', 'rPr'),
          AUTHORABLE_RUN_PROPERTIES
        ).map((property) => [property.localName, property])
      );
      for (const change of edit.changes) {
        if (!change.after.present) bag.delete(change.name);
        else
          bag.set(change.name, {
            localName: change.name,
            attributes: change.after.value as Record<string, string>,
          });
      }
      ops = [
        {
          op: 'setRunProperties',
          paragraphId: target.paragraphId,
          start: target.start,
          end: target.end,
          properties: [...bag.values()],
        },
      ];
    } else throw new Error('A cell content plan emitted a structural operation');
    for (const op of ops) {
      const result = applyTreeOp(current, op);
      if (!result.ok) throw new Error(result.reason);
      current = result.part;
    }
  }
  return current;
}
function child(node: OoxmlNode | null | undefined, name: string): OoxmlNode | undefined {
  return node?.kind === 'textValue'
    ? undefined
    : node?.children.find((entry) => entry.kind !== 'textValue' && entry.localName === name);
}

test('C01 table cell replacement retains all other cell IDs/properties and native table geometry', async () => {
  const fixture = reference.cases.find((item) => item.name === 'tables and links')!;
  expect(fixture.beforeMarkdown).toContain('| One | Two |');
  expect(fixture.afterMarkdown).toContain('| One | Changed |');
  const { input, selected } = await source();
  const before = canonicalOoxmlFingerprint(input.part);
  const planned = plan(input, [{ paragraphId: selected.paragraphIds[0]!, text: 'Changed' }]);
  expect(
    planned.operations.every(
      (edit) => edit.kind === 'text' && edit.paragraphId === selected.paragraphIds[0]
    )
  ).toBe(true);
  const applied = apply(input.part, planned);
  expect(paragraphTextOf(applied, selected.paragraphIds[0]!)).toBe('Changed');
  for (const keep of input.projection.cells.filter((record) => record.id !== selected.id)) {
    expect(findNode(applied, keep.id) === findNode(input.part, keep.id)).toBe(true);
    expect(canonicalOoxmlFingerprint(findNode(applied, keep.id)!)).toBe(
      canonicalOoxmlFingerprint(findNode(input.part, keep.id)!)
    );
  }
  for (const [id, names] of [
    [input.tableId, ['tblPr', 'tblGrid']],
    [input.rowId, ['trPr']],
    [input.cellId, ['tcPr']],
  ] as const)
    for (const name of names) {
      const original = child(findNode(input.part, id), name)!;
      const after = child(findNode(applied, id), name)!;
      expect([after.id, canonicalOoxmlFingerprint(after)]).toEqual([
        original.id,
        canonicalOoxmlFingerprint(original),
      ]);
    }
  expect(findNode(applied, input.tableId)?.id).toBe(input.tableId);
  expect(findNode(applied, input.rowId)?.id).toBe(input.rowId);
  expect(findNode(applied, input.cellId)?.id).toBe(input.cellId);
  expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
  expect(decodeEditPlan(serializeEditPlan(planned))).toEqual({ success: true, plan: planned });
});

test('cell mark changes precede text changes and replacement text inherits the requested native marks', async () => {
  const { input, selected } = await source();
  const id = selected.paragraphIds[0]!;
  const planned = plan(input, [
    {
      paragraphId: id,
      text: 'Changed',
      formats: [
        { start: 0, end: 3, target: { properties: [{ localName: 'b' }, { localName: 'i' }] } },
      ],
    },
  ]);
  expect(planned.operations.map((edit) => edit.kind)).toEqual(['format', 'text']);
  const applied = apply(input.part, planned);
  expect(paragraphTextOf(applied, id)).toBe('Changed');
  const run = runsCovering(applied, id, 0, 7)[0]!;
  expect(
    authoredProperties(propertyContainer(run, 'runProperties', 'rPr'), AUTHORABLE_RUN_PROPERTIES)
      .map((property) => property.localName)
      .sort()
  ).toEqual(['b', 'i']);
  expect(canonicalOoxmlFingerprint(child(findNode(applied, input.cellId), 'tcPr')!)).toBe(
    canonicalOoxmlFingerprint(child(findNode(input.part, input.cellId), 'tcPr')!)
  );
});

test('multi-paragraph cell edits preserve paragraph and cell boundaries', async () => {
  const { input, selected } = await source(
    table(row(cell(p('One')), cell(p('First') + p('Second')))) + p('Tail'),
    'First'
  );
  const planned = plan(input, [
    { paragraphId: selected.paragraphIds[0]!, text: 'First updated' },
    { paragraphId: selected.paragraphIds[1]!, text: 'Second updated' },
  ]);
  const applied = apply(input.part, planned);
  expect(selected.paragraphIds.map((id) => paragraphTextOf(applied, id))).toEqual([
    'First updated',
    'Second updated',
  ]);
  const owner = findNode(applied, input.cellId)!;
  if (owner.kind !== 'tableCell') throw new Error('Missing native cell');
  expect(owner.children.filter((node) => node.kind === 'paragraph').map((node) => node.id)).toEqual(
    [...selected.paragraphIds]
  );
});

test('cell paragraphs inside native content controls retain their wrapper identity', async () => {
  const wrapped =
    '<w:sdt><w:sdtPr><w:tag w:val="question"/></w:sdtPr><w:sdtContent>' +
    p('Question') +
    '</w:sdtContent></w:sdt>';
  const { input, selected } = await source(
    table(row(cell(p('One')), cell(wrapped))) + p('Tail'),
    'Question'
  );
  const owner = findNode(input.part, input.cellId)!;
  const wrapper = child(owner, 'sdt')!;
  const planned = plan(input, [
    { paragraphId: selected.paragraphIds[0]!, text: 'Question updated' },
  ]);
  const applied = apply(input.part, planned);
  expect(paragraphTextOf(applied, selected.paragraphIds[0]!)).toBe('Question updated');
  expect(findNode(applied, wrapper.id)?.id).toBe(wrapper.id);
  expect(canonicalOoxmlFingerprint(child(findNode(applied, wrapper.id), 'sdtPr')!)).toBe(
    canonicalOoxmlFingerprint(child(wrapper, 'sdtPr')!)
  );
});

test('nested table paragraphs only belong to the closest native cell', async () => {
  const { input, selected } = await source(
    table(
      row(
        cell(
          p('Outer before') +
            table(row(cell(p('Inner')), cell(p('Inner neighbor')))) +
            p('Outer after')
        ),
        cell(p('Outside'))
      )
    ) + p('Tail'),
    'Outer before'
  );
  const nested = input.projection.cells.find(
    (record) =>
      record.id !== selected.id &&
      record.paragraphIds.some((id) => paragraphTextOf(input.part, id) === 'Inner')
  )!;
  expect(
    planTableCellEdits(input, [{ paragraphId: nested.paragraphIds[0]!, text: 'Wrong owner' }])
  ).toMatchObject({ success: false, code: 'ambiguous-target' });
  const forged = {
    ...input,
    projection: {
      ...input.projection,
      cells: input.projection.cells.map((record) =>
        record.id === selected.id
          ? { ...record, paragraphIds: [...record.paragraphIds, ...nested.paragraphIds] }
          : record
      ),
    },
  };
  expect(
    planTableCellEdits(forged, [
      { paragraphId: nested.paragraphIds[0]!, text: 'Still wrong owner' },
    ])
  ).toMatchObject({ success: false, code: 'ambiguous-target' });
  const nestedSource = {
    ...input,
    tableId: nested.tableId,
    rowId: nested.rowId,
    cellId: nested.id,
  };
  const applied = apply(
    input.part,
    plan(nestedSource, [{ paragraphId: nested.paragraphIds[0]!, text: 'Inner updated' }])
  );
  expect(paragraphTextOf(applied, nested.paragraphIds[0]!)).toBe('Inner updated');
  for (const id of selected.paragraphIds)
    expect(paragraphTextOf(applied, id)).toBe(paragraphTextOf(input.part, id));
});

test('native merged-cell metadata keeps grid spans and visible anchors intact', async () => {
  const body =
    table(
      row(cell(p('Merged'), ' <w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>')) +
        row(cell('<w:p/>', '<w:gridSpan w:val="2"/><w:vMerge/>'))
    ) + p('Tail');
  const { input, selected } = await source(body, 'Merged');
  expect(selected.gridSpan).toBe(2);
  const planned = plan(input, [{ paragraphId: selected.paragraphIds[0]!, text: 'Merged updated' }]);
  const applied = apply(input.part, planned);
  expect(paragraphTextOf(applied, selected.paragraphIds[0]!)).toBe('Merged updated');
  expect(canonicalOoxmlFingerprint(child(findNode(applied, input.cellId), 'tcPr')!)).toBe(
    canonicalOoxmlFingerprint(child(findNode(input.part, input.cellId), 'tcPr')!)
  );
  const continuation = input.projection.cells.find((record) => record.verticalContinuation)!;
  expect(Boolean(continuation)).toBe(true);
  expect(
    planTableCellEdits({ ...input, cellId: continuation.id, rowId: continuation.rowId }, [])
  ).toMatchObject({ success: false, code: 'unsupported-operation' });
});

test('empty cells accept native Unicode text and manual breaks without changing table shape', async () => {
  const { input, selected } = await source(
    table(row(cell(p('One')), cell('<w:p/>'))) + p('Tail'),
    ''
  );
  const id = selected.paragraphIds[0]!;
  const applied = apply(input.part, plan(input, [{ paragraphId: id, text: '😀\tA\nB' }]));
  expect(paragraphTextOf(applied, id)).toBe('😀\tA\nB');
  expect(input.projection.cells.map((record) => findNode(applied, record.id)?.kind)).toEqual([
    'tableCell',
    'tableCell',
  ]);
});

test('no-op cell requests emit no operations or revisions', async () => {
  const { input, selected } = await source();
  expect(plan(input, [])).toEqual({ schemaVersion: 1, revision: input.revision, operations: [] });
  expect(plan(input, [{ paragraphId: selected.paragraphIds[0]!, text: 'Two' }]).operations).toEqual(
    []
  );
});

test('stale projection, cross-cell ownership, duplicate patches, invalid formats, and ambiguous metadata refuse atomically', async () => {
  const { input, selected } = await source();
  const before = canonicalOoxmlFingerprint(input.part),
    id = selected.paragraphIds[0]!;
  const neighbor = input.projection.cells.find((record) => record.id !== selected.id)!;
  for (const altered of [
    { ...input, revision: { ...input.revision, revision: input.revision.revision + 1 } },
    { ...input, rowId: neighbor.rowId, cellId: neighbor.id },
    { ...input, tableId: 'missing' },
    { ...input, projection: { ...input.projection, cells: [...input.projection.cells, selected] } },
  ])
    expect(planTableCellEdits(altered, [{ paragraphId: id, text: 'Changed' }]).success).toBe(false);
  const patches: CellParagraphPatch[][] = [
    [
      { paragraphId: id, text: 'Changed' },
      { paragraphId: id, text: 'Again' },
    ],
    [{ paragraphId: neighbor.paragraphIds[0]!, text: 'Wrong cell' }],
    [{ paragraphId: 'removed', text: 'Missing' }],
    [
      {
        paragraphId: id,
        formats: [
          { start: 0, end: 3, target: { properties: [{ localName: 'b' }] } },
          { start: 1, end: 2, target: { properties: [{ localName: 'i' }] } },
        ],
      },
    ],
    [
      { paragraphId: id, formats: [null, { start: 0, end: 3, target: { properties: [] } }] },
    ] as unknown as CellParagraphPatch[],
  ];
  for (const patch of patches) expect(planTableCellEdits(input, patch).success).toBe(false);
  let read = false;
  const patch = Object.defineProperty({ paragraphId: id }, 'text', {
    enumerable: true,
    get() {
      read = true;
      return 'Bad';
    },
  }) as CellParagraphPatch;
  expect(planTableCellEdits(input, [patch]).success).toBe(false);
  expect(read).toBe(false);
  expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
});
