import { findNode } from '../package/ooxml-edit.ts';
// Tracked cell merges and splits through Word cell revisions.

import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  serializeOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import {
  TreeDocumentStore,
  collectRevisionSites,
  planRevisionBatch,
  revisionItemsOf,
} from '../index.ts';
import { applyTreeOp, validateTreeOp, type TreeDocOp } from '../store/tree-ops.ts';

const W = WML_NAMESPACE_URI;
const revision = { author: 'Reviewer', date: '2026-10-20T08:00:00Z' };
const p = (text: string) => (text === '' ? '<w:p/>' : `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`);
const cell = (text: string, width: number, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${text.split('|').map(p).join('')}</w:tc>`;
const body = (first = 'a', second = '') =>
  `<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid>` +
  `<w:tr>${cell(first, 1200)}${cell(second, 2400)}${cell('c', 3600)}</w:tr>` +
  `<w:tr>${cell('d|d2', 1200)}${cell('e', 2400)}${cell('f', 3600)}</w:tr>` +
  `<w:tr>${cell('g', 1200)}${cell('h', 2400)}${cell('i', 3600)}</w:tr></w:tbl><w:p><w:r><w:t>Tail</w:t></w:r></w:p>`;
function load(xml: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${xml}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
const kids = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);
const table = (part: OoxmlPart) => kids(kids(part.root)[0]!).find((node) => node.kind === 'table')!;
const cells = (part: OoxmlPart) =>
  kids(table(part))
    .filter((row) => row.kind === 'tableRow')
    .map((row) =>
      kids(row)
        .filter((node) => node.kind === 'tableCell')
        .map((node) => node.id)
    );
function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
function decide(part: OoxmlPart, action: 'accept' | 'reject'): OoxmlPart {
  const batch = planRevisionBatch(part, action);
  expect(batch.result.skipped).toEqual([]);
  let current = part;
  for (const op of batch.ops) current = apply(current, op);
  expect(revisionItemsOf(current)).toHaveLength(0);
  return current;
}
/** Fresh Word paragraph IDs differ. Review decisions spell the default `continue` explicitly. */
const comparable = (part: OoxmlPart) =>
  serializeOoxmlPart(part)
    .replace(/ w14:(paraId|textId)="[0-9A-F]+"/g, '')
    .replace(/<w:vMerge w:val="continue"\/>/g, '<w:vMerge/>');
function exact(part: OoxmlPart, op: TreeDocOp) {
  expect(validateTreeOp(part, { ...op, revision } as TreeDocOp)).toBeNull();
  const pending = apply(part, { ...op, revision } as TreeDocOp);
  expect(revisionItemsOf(pending).length).toBeGreaterThan(0);
  expect(canonicalOoxmlFingerprint(decide(pending, 'reject'))).toBe(
    canonicalOoxmlFingerprint(part)
  );
  expect(comparable(decide(pending, 'accept'))).toBe(comparable(apply(part, op)));
  return pending;
}

const merges: [string, string, string, (ids: string[][]) => string[]][] = [
  ['horizontal', 'a', '', (ids) => [ids[0]![0]!, ids[0]![1]!]],
  ['horizontal into an empty owner', '', 'b', (ids) => [ids[0]![0]!, ids[0]![1]!]],
  ['vertical', 'a', 'b', (ids) => [ids[0]![0]!, ids[1]![0]!]],
  ['rectangle', 'a', 'b', (ids) => [ids[0]![0]!, ids[0]![1]!, ids[1]![0]!, ids[1]![1]!]],
  ['three-row column', 'a', 'b', (ids) => [ids[0]![1]!, ids[1]![1]!, ids[2]![1]!]],
];
for (const [name, first, second, pick] of merges)
  test(`tracked ${name} merge accepts to the direct merge and rejects exactly`, () => {
    const part = load(body(first, second));
    const pending = exact(part, {
      op: 'mergeTableCells',
      tableId: table(part).id,
      cellIds: pick(cells(part)),
    });
    expect(serializeOoxmlPart(pending)).toMatch(/<w:cell(Del|Merge) /);
  });

for (const [name, rows, columns, distribution] of [
  ['two columns', 1, 2, undefined],
  ['three rows', 3, 1, undefined],
  ['a grid of two by two', 2, 2, 'equal-width'],
] as const)
  test(`tracked split into ${name} accepts to the direct split and rejects exactly`, () => {
    const part = load(body());
    const pending = exact(part, {
      op: 'splitTableCell',
      tableId: table(part).id,
      cellId: cells(part)[1]![1]!,
      rows,
      columns,
      ...(distribution ? { columnDistribution: distribution } : {}),
    });
    expect(serializeOoxmlPart(pending)).toContain('<w:cellIns ');
  });

test('a tracked split of a merged cell along existing grid columns restores the merge on reject', () => {
  const merged = load(body());
  const id = table(merged).id;
  const part = apply(merged, {
    op: 'mergeTableCells',
    tableId: id,
    cellIds: [cells(merged)[0]![0]!, cells(merged)[0]![1]!],
  });
  exact(part, {
    op: 'splitTableCell',
    tableId: id,
    cellId: cells(part)[0]![0]!,
    rows: 1,
    columns: 2,
    columnDistribution: 'existing-grid',
  });
});

test('tracked merges refuse stacked histories, moved tables, and invalid attribution', () => {
  const part = load(body());
  const id = table(part).id,
    ids = cells(part);
  const pending = apply(part, {
    op: 'mergeTableCells',
    tableId: id,
    cellIds: [ids[0]![0]!, ids[0]![1]!],
    revision,
  });
  expect(
    applyTreeOp(pending, {
      op: 'splitTableCell',
      tableId: id,
      cellId: ids[2]![2]!,
      rows: 1,
      columns: 2,
      revision,
    })
  ).toMatchObject({ ok: false, reason: 'unsupported' });
  expect(
    validateTreeOp(part, {
      op: 'mergeTableCells',
      tableId: id,
      cellIds: [ids[0]![0]!, ids[0]![1]!],
      revision: { author: '' },
    })
  ).toBe('invalidArgs');
  const nested = load(
    body().replace(
      `${p('e')}</w:tc>`,
      `${p('e')}<w:tbl><w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc>${p('n')}</w:tc></w:tr></w:tbl>${p('')}</w:tc>`
    )
  );
  const nids = cells(nested);
  expect(
    applyTreeOp(nested, {
      op: 'mergeTableCells',
      tableId: table(nested).id,
      cellIds: [nids[1]![0]!, nids[1]![1]!],
      revision,
    })
  ).toMatchObject({ ok: false, reason: 'unsupported' });
});

for (const covered of ['cover', 'cover|']) {
  test(`a tracked merge resolves this transaction's dependent text and retains blank paragraphs: ${covered}`, () => {
    const part = load(body('head', covered)),
      ids = cells(part),
      targetCell = ids[0]![1]!;
    const paragraph = kids(findNode(part, targetCell)!).find((node) => node.kind === 'paragraph')!;
    const op = {
      op: 'mergeTableCells' as const,
      tableId: table(part).id,
      cellIds: ids[0]!.slice(0, 2),
    };
    const direct = apply(
      apply(part, { op: 'deleteText', paragraphId: paragraph.id, start: 0, end: 5 }),
      op
    );
    const store = new TreeDocumentStore(part);
    let publications = 0;
    store.subscribe(() => publications++);
    expect(
      store.transact((ctx) => {
        ctx.apply({ op: 'deleteText', paragraphId: paragraph.id, start: 0, end: 5, revision });
        ctx.apply({ ...op, revision });
      }).ok
    ).toBe(true);
    expect(publications).toBe(1);
    const pending = store.part;
    expect(comparable(decide(pending, 'accept'))).toBe(comparable(direct));
    expect(canonicalOoxmlFingerprint(decide(pending, 'reject'))).toBe(
      canonicalOoxmlFingerprint(part)
    );
    expect(store.undo()).not.toBeNull();
    expect(store.part).toBe(part);
    expect(store.redo()).not.toBeNull();
    expect(store.part).toBe(pending);
  });
}

test('a tracked merge refuses earlier same-author text proposals instead of adopting them', () => {
  const part = load(body('head', 'cover')),
    ids = cells(part),
    paragraph = kids(findNode(part, ids[0]![1]!)!).find((node) => node.kind === 'paragraph')!;
  const pending = apply(part, {
      op: 'deleteText',
      paragraphId: paragraph.id,
      start: 0,
      end: 5,
      revision,
    }),
    store = new TreeDocumentStore(pending);
  let publications = 0;
  store.subscribe(() => publications++);
  expect(
    store.transact((ctx) => {
      ctx.apply({
        op: 'mergeTableCells',
        tableId: table(pending).id,
        cellIds: ids[0]!.slice(0, 2),
        revision,
      });
    }).ok
  ).toBe(false);
  expect(store.part).toBe(pending);
  expect(store.revision).toBe(0);
  expect(publications).toBe(0);
  expect(store.canUndo).toBe(false);
});

test('a store merge and later tracked typing share one collision-free revision counter', () => {
  const part = load(body('head', 'cover')),
    ids = cells(part),
    store = new TreeDocumentStore(part);
  const tail = kids(kids(part.root)[0]!).find((node) => node.kind === 'paragraph')!;
  expect(
    store.transact((ctx) => {
      ctx.apply({
        op: 'mergeTableCells',
        tableId: table(part).id,
        cellIds: ids[0]!.slice(0, 2),
        revision,
      });
      ctx.apply({ op: 'insertText', paragraphId: tail.id, offset: 4, text: ' more', revision });
    }).ok
  ).toBe(true);
  const families = new Map<string, Set<string>>();
  for (const site of collectRevisionSites(store.part)) {
    const id = site.node.attributes.find(
      (attribute) => attribute.namespaceUri === W && attribute.localName === 'id'
    )!.value;
    const names = families.get(id) ?? new Set<string>();
    names.add(site.node.localName);
    families.set(id, names);
  }
  expect([...families.values()].map((names) => names.size)).toEqual(
    [...families.values()].map(() => 1)
  );
  expect(comparable(decide(store.part, 'reject'))).toBe(comparable(part));
});
