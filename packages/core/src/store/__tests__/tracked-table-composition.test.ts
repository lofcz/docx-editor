import { projectRevisionFormatting } from '../../layout/revision-formatting-projection.ts';
import { readTableStructure } from '../../layout/semantic-table.ts';
import { findNode, insertChildren, createNodeIdAllocator } from '../package/ooxml-edit.ts';
import { cloneWithNewIds } from '../store/tree-op-nodes.ts';
import { withPart } from '../index.ts';
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
import { applyTreeOp, type TreeDocOp } from '../store/tree-ops.ts';

const W = WML_NAMESPACE_URI;
const revision = { author: 'Reviewer', date: '2026-10-20T08:00:00Z' };
const p = (text: string) =>
  `<w:p><w:pPr><w:rPr/></w:pPr>${text === '' ? '' : `<w:r><w:t>${text}</w:t></w:r>`}</w:p>`;
const cell = (text: string, width: number, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${text.split('|').map(p).join('')}</w:tc>`;
const body = (first = 'a', second = '') =>
  `<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid>` +
  `<w:tr><w:trPr/>${cell(first, 1200)}${cell(second, 2400)}${cell('c', 3600)}</w:tr>` +
  `<w:tr><w:trPr/>${cell('d|d2', 1200)}${cell('e', 2400)}${cell('f', 3600)}</w:tr>` +
  `<w:tr><w:trPr/>${cell('g', 1200)}${cell('h', 2400)}${cell('i', 3600)}</w:tr></w:tbl><w:p><w:r><w:t>Tail</w:t></w:r></w:p>`;
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

const current = (ctx: import('../store/tree-transaction-context.ts').TransactionContext) =>
  ctx.currentPackage().parts.get('/word/document.xml')!;
const grid = (part: OoxmlPart) =>
  kids(kids(table(part)).find((node) => node.kind === 'tableGrid')!);
const rowIds = (part: OoxmlPart) =>
  kids(table(part))
    .filter((node) => node.kind === 'tableRow')
    .map((node) => node.id);
const scenarios = [
  'row-merge',
  'column-merge',
  'merge-column',
  'two-columns',
  'column-row-delete',
  'deleted-column-merge',
  'two-merges',
  'born-column-merge',
  'born-row-merge',
  'column-split',
  'row-split',
  'inside-span-split',
  'row-column-born-merge',
  'author-rows-merge',
  'column-author-rows',
  'opaque-column-merge',
] as const;
for (const scenario of scenarios)
  test(`native compound ${scenario} matches direct decisions and exact native history`, () => {
    const base = body('head', 'cover');
    const xml =
      scenario === 'opaque-column-merge'
        ? base
            .replace(
              '<w:tcPr>',
              '<w:tcPr><x:bag xmlns:x="urn:test:opaque"><x:value x:key="keep"/></x:bag>'
            )
            .replace(
              '<w:t>head</w:t>',
              '<w:t>head</w:t><x:note xmlns:x="urn:test:opaque" x:flag="keep"/>'
            )
        : base;
    const merged = xml.replace(
      cell('head', 1200) + cell('cover', 2400),
      cell('head', 3600, '<w:gridSpan w:val="2"/>')
    );
    const part = load(
        ['column-split', 'row-split', 'inside-span-split'].includes(scenario) ? merged : xml
      ),
      store = new TreeDocumentStore(part),
      tableId = table(part).id,
      initial = cells(part);
    const run = (tracked: boolean) => {
      const apply = (get: () => OoxmlPart, put: (op: TreeDocOp) => void) => {
        const merge = (ids: readonly string[]) =>
          put({ op: 'mergeTableCells', tableId, cellIds: ids, ...(tracked ? { revision } : {}) });
        const column = (where: 'left' | 'right' = 'right', index = 2) =>
          put({
            op: 'insertTableColumn',
            tableId,
            gridColumnId: grid(get())[index]!.id,
            where,
            ...(tracked ? { revision } : {}),
          });
        const row = () =>
          put({
            op: 'insertTableRow',
            tableId,
            rowId: rowIds(get()).at(-1)!,
            where: 'below',
            ...(tracked ? { revision } : {}),
          });
        if (scenario === 'author-rows-merge' || scenario === 'column-author-rows') {
          if (scenario === 'column-author-rows') column();
          put({
            op: 'authorTable',
            action: {
              kind: 'existing',
              tableId,
              mutation: {
                kind: 'insertRows',
                rowId: rowIds(get()).at(-1)!,
                location: 'after',
                count: 2,
              },
            },
            ...(tracked ? { revision } : {}),
          });
          if (scenario === 'author-rows-merge') merge(initial[0]!.slice(0, 2));
        } else if (
          scenario === 'column-split' ||
          scenario === 'row-split' ||
          scenario === 'inside-span-split'
        ) {
          if (scenario === 'row-split') row();
          else column('right', scenario === 'inside-span-split' ? 0 : 2);
          put({
            op: 'splitTableCell',
            tableId,
            cellId: initial[0]![0]!,
            rows: 1,
            columns: scenario === 'inside-span-split' ? 3 : 2,
            columnDistribution: 'existing-grid',
            ...(tracked ? { revision } : {}),
          });
        } else if (scenario === 'row-column-born-merge') {
          row();
          column();
          merge(cells(get()).at(-1)!.slice(0, 2));
        } else if (scenario === 'row-merge') {
          row();
          merge(initial[0]!.slice(0, 2));
        } else if (scenario === 'column-merge' || scenario === 'opaque-column-merge') {
          column();
          merge(initial[0]!.slice(0, 2));
        } else if (scenario === 'merge-column') {
          merge(initial[0]!.slice(0, 2));
          column();
        } else if (scenario === 'two-columns') {
          column();
          column();
        } else if (scenario === 'column-row-delete') {
          column();
          put({
            op: 'deleteTableRow',
            tableId,
            rowId: rowIds(get())[1]!,
            ...(tracked ? { revision } : {}),
          });
        } else if (scenario === 'deleted-column-merge') {
          put({
            op: 'deleteTableColumn',
            tableId,
            gridColumnId: grid(get())[1]!.id,
            ...(tracked ? { revision } : {}),
          });
          merge([initial[0]![0]!, initial[0]![2]!]);
        } else if (scenario === 'two-merges') {
          merge(initial[0]!.slice(0, 2));
          merge(initial[2]!.slice(0, 2));
        } else if (scenario === 'born-column-merge') {
          column('left', 1);
          merge(cells(get())[0]!.slice(0, 2));
        } else {
          row();
          merge(cells(get()).at(-1)!.slice(0, 2));
        }
      };
      if (tracked) {
        let count = 0;
        store.subscribe(() => count++);
        const result = store.transact((ctx) =>
          apply(
            () => current(ctx),
            (op) => {
              ctx.apply(op);
            }
          )
        );
        expect(result).toMatchObject({ ok: true });
        expect(count).toBe(1);
        return store.part;
      }
      let direct = part;
      apply(
        () => direct,
        (op) => {
          const result = applyTreeOp(
            direct,
            scenario === 'deleted-column-merge' && op.op === 'deleteTableColumn'
              ? { ...op, revision }
              : op
          );
          if (!result.ok) throw new Error(result.reason);
          direct =
            scenario === 'deleted-column-merge' && op.op === 'deleteTableColumn'
              ? decide(result.part, 'accept')
              : result.part;
        }
      );
      return direct;
    };
    const golden = run(false),
      pending = run(true);
    const shape = (source: OoxmlPart, mode: 'original' | 'proposed') => {
      const structure = readTableStructure(
        projectRevisionFormatting(
          table(source) as import('../package/ooxml-tree.ts').OoxmlElement,
          mode
        ),
        500,
        0,
        undefined,
        mode
      )!;
      return structure.rows.map((row) =>
        row.cells.map((cell) => ({
          column: cell.resolvedGridColumn ?? cell.gridColumn,
          span: cell.resolvedGridSpan ?? cell.gridSpan,
          continued: cell.vMergeContinue,
          width: structure.columnWidthsPt
            .slice(cell.gridColumn, cell.gridColumn + cell.gridSpan)
            .reduce((sum, value) => sum + value, 0),
        }))
      );
    };
    expect(shape(pending, 'original')).toEqual(shape(part, 'original'));
    expect(shape(pending, 'proposed')).toEqual(shape(golden, 'proposed'));
    expect(collectRevisionSites(pending).every((site) => !site.refused)).toBe(true);
    expect(comparable(decide(pending, 'reject'))).toBe(comparable(part));
    expect(comparable(decide(pending, 'accept'))).toBe(comparable(golden));
    const fingerprint = canonicalOoxmlFingerprint(pending);
    expect(store.undo()).not.toBeNull();
    expect(store.part).toBe(part);
    expect(store.redo()).not.toBeNull();
    expect(canonicalOoxmlFingerprint(store.part)).toBe(fingerprint);
    const read = readOoxmlPart(serializeOoxmlPart(pending), {
      name: pending.name,
      contentType: pending.contentType,
    });
    if (!read.ok) throw new Error(read.reason);
    const reopened = read.part;
    expect(comparable(decide(reopened, 'accept'))).toBe(comparable(golden));
    expect(comparable(decide(reopened, 'reject'))).toBe(comparable(part));
  });

for (const earlier of ['row', 'column', 'merge'] as const)
  test(`same-author earlier ${earlier} structure never gains current transaction ownership`, () => {
    const initial = load(body('head', 'cover')),
      tableId = table(initial).id,
      originalCells = cells(initial);
    const pending = apply(
      initial,
      earlier === 'row'
        ? {
            op: 'insertTableRow',
            tableId,
            rowId: rowIds(initial).at(-1)!,
            where: 'below',
            revision,
          }
        : earlier === 'column'
          ? {
              op: 'insertTableColumn',
              tableId,
              gridColumnId: grid(initial).at(-1)!.id,
              where: 'right',
              revision,
            }
          : { op: 'mergeTableCells', tableId, cellIds: originalCells[0]!.slice(0, 2), revision }
    );
    const store = new TreeDocumentStore(pending),
      tail = kids(kids(pending.root)[0]!).find((node) => node.kind === 'paragraph')!;
    expect(
      store.transact((ctx) => {
        ctx.apply({ op: 'insertText', paragraphId: tail.id, offset: 4, text: ' redo' });
      }).ok
    ).toBe(true);
    const redo = canonicalOoxmlFingerprint(store.part);
    expect(store.undo()).not.toBeNull();
    const before = store.part;
    let notifications = 0;
    store.subscribe(() => notifications++);
    const result = store.transact((ctx) => {
      ctx.apply({ op: 'insertText', paragraphId: tail.id, offset: 4, text: ' staged', revision });
      ctx.apply({
        op: 'insertTableColumn',
        tableId,
        gridColumnId: grid(current(ctx)).at(-1)!.id,
        where: 'right',
        revision,
      });
    });
    expect(result.ok).toBe(false);
    expect(store.part).toBe(before);
    expect(notifications).toBe(0);
    expect(store.canRedo).toBe(true);
    expect(store.redo()).not.toBeNull();
    expect(canonicalOoxmlFingerprint(store.part)).toBe(redo);
  });

for (const mutation of [
  'insert-column',
  'delete-column',
  'split',
  'delete-row',
  'insert-row',
] as const)
  test(`later ${mutation} inside a virtual merge refuses without reinterpreting physical slots`, () => {
    const part = load(body('head', 'cover')),
      store = new TreeDocumentStore(part),
      tableId = table(part).id,
      originalCells = cells(part);
    let publications = 0;
    store.subscribe(() => publications++);
    const result = store.transact((ctx) => {
      ctx.apply({
        op: 'mergeTableCells',
        tableId,
        cellIds:
          mutation === 'insert-row'
            ? [originalCells[0]![0]!, originalCells[1]![0]!]
            : originalCells[0]!.slice(0, 2),
        revision,
      });
      if (mutation === 'insert-column')
        ctx.apply({
          op: 'insertTableColumn',
          tableId,
          gridColumnId: grid(current(ctx))[0]!.id,
          where: 'right',
          revision,
        });
      else if (mutation === 'delete-column')
        ctx.apply({
          op: 'deleteTableColumn',
          tableId,
          gridColumnId: grid(current(ctx))[0]!.id,
          revision,
        });
      else if (mutation === 'split')
        ctx.apply({
          op: 'splitTableCell',
          tableId,
          cellId: originalCells[0]![0]!,
          rows: 1,
          columns: 2,
          revision,
        });
      else if (mutation === 'delete-row')
        ctx.apply({ op: 'deleteTableRow', tableId, rowId: rowIds(current(ctx))[0]!, revision });
      else
        ctx.apply({
          op: 'insertTableRow',
          tableId,
          rowId: rowIds(current(ctx))[0]!,
          where: 'below',
          revision,
        });
    });
    expect(result).toMatchObject({ ok: false, reason: 'unsupported' });
    expect(store.part).toBe(part);
    expect(store.revision).toBe(0);
    expect(publications).toBe(0);
    expect(store.canUndo).toBe(false);
  });

test('a later native failure discards compound geometry and preserves prior redo', () => {
  const part = load(body('head', 'cover')),
    store = new TreeDocumentStore(part),
    tableId = table(part).id,
    tail = kids(kids(part.root)[0]!).find((node) => node.kind === 'paragraph')!;
  expect(
    store.transact((ctx) => {
      ctx.apply({ op: 'insertText', paragraphId: tail.id, offset: 4, text: ' redo' });
    }).ok
  ).toBe(true);
  const redo = canonicalOoxmlFingerprint(store.part);
  expect(store.undo()).not.toBeNull();
  let publications = 0;
  store.subscribe(() => publications++);
  expect(
    store.transact((ctx) => {
      ctx.apply({
        op: 'insertTableColumn',
        tableId,
        gridColumnId: grid(current(ctx)).at(-1)!.id,
        where: 'right',
        revision,
      });
      ctx.apply({ op: 'mergeTableCells', tableId, cellIds: cells(part)[0]!.slice(0, 2), revision });
      ctx.apply({
        op: 'insertText',
        paragraphId: 'missing-native-owner',
        offset: 0,
        text: 'bad',
        revision,
      });
    }).ok
  ).toBe(false);
  expect(store.part).toBe(part);
  expect(publications).toBe(0);
  expect(store.canRedo).toBe(true);
  expect(store.redo()).not.toBeNull();
  expect(canonicalOoxmlFingerprint(store.part)).toBe(redo);
});

test('compound table stamps and later text share collision-free native revision allocation across package callbacks', () => {
  const part = load(body('head', 'cover')),
    store = new TreeDocumentStore(part),
    tableId = table(part).id,
    tail = kids(kids(part.root)[0]!).find((node) => node.kind === 'paragraph')!;
  expect(
    store.transact((ctx) => {
      const sites = new Set<string>(),
        addresses = new Set<string>();
      const write = (op: TreeDocOp) => {
        expect(ctx.apply(op)).toBe(true);
        const currentSites = collectRevisionSites(current(ctx));
        const fresh = currentSites.filter((site) => !sites.has(site.node.id));
        const freshIds = new Set(
          fresh.map(
            (site) =>
              site.node.attributes.find(
                (attribute) => attribute.namespaceUri === W && attribute.localName === 'id'
              )!.value
          )
        );
        // One row stamp deliberately shares its address with its own cellIns markers.
        // A later operation must never join that earlier logical native revision.
        expect([...freshIds].filter((id) => addresses.has(id))).toEqual([]);
        for (const site of currentSites) sites.add(site.node.id);
        for (const id of freshIds) addresses.add(id);
      };
      write({
        op: 'insertTableRow',
        tableId,
        rowId: rowIds(part).at(-1)!,
        where: 'below',
        revision,
      });
      ctx.applyPackage((pkg) => pkg);
      write({
        op: 'authorTable',
        action: {
          kind: 'existing',
          tableId,
          mutation: {
            kind: 'insertRows',
            rowId: rowIds(current(ctx)).at(-1)!,
            location: 'after',
            count: 2,
          },
        },
        revision,
      });
      write({
        op: 'insertTableColumn',
        tableId,
        gridColumnId: grid(current(ctx)).at(-1)!.id,
        where: 'right',
        revision,
      });
      write({ op: 'mergeTableCells', tableId, cellIds: cells(part)[0]!.slice(0, 2), revision });
      write({ op: 'insertText', paragraphId: tail.id, offset: 4, text: ' added', revision });
    }).ok
  ).toBe(true);
  expect(comparable(decide(store.part, 'reject'))).toBe(comparable(part));
});

test('a package callback cannot grant a foreign table history current-frame ownership by matching its complete revision address', () => {
  const part = load(body('head', 'cover')),
    store = new TreeDocumentStore(part),
    tableId = table(part).id;
  let count = 0;
  store.subscribe(() => count++);
  const result = store.transact((ctx) => {
    expect(
      ctx.apply({
        op: 'insertTableRow',
        tableId,
        rowId: rowIds(part).at(-1)!,
        where: 'below',
        revision,
      })
    ).toBe(true);
    const foreign = apply(part, {
      op: 'setTableRowHeight',
      tableId,
      rowId: rowIds(part)[1]!,
      heightTwips: 900,
      revision,
    });
    const marker = collectRevisionSites(foreign).find(
      (site) => site.node.localName === 'trPrChange'
    )!.node;
    const prior = collectRevisionSites(current(ctx)).find(
      (site) => site.node.localName === 'ins'
    )!.node;
    for (const name of ['id', 'author', 'date'])
      expect(marker.attributes.find((attr) => attr.localName === name)?.value).toBe(
        prior.attributes.find((attr) => attr.localName === name)?.value
      );
    ctx.applyPackage((pkg) => {
      const before = pkg.parts.get(part.name)!,
        row = findNode(before, rowIds(part)[1]!)!,
        properties = kids(row).find(
          (node) => node.kind !== 'textValue' && node.localName === 'trPr'
        )!;
      const copy = cloneWithNewIds(marker, createNodeIdAllocator(before)),
        inserted = insertChildren(before, properties.id, kids(properties).length, [copy]);
      if (!inserted.ok) throw new Error('Invalid foreign fixture');
      return withPart(pkg, inserted.part);
    });
    ctx.apply({ op: 'mergeTableCells', tableId, cellIds: cells(part)[0]!.slice(0, 2), revision });
  });
  expect(result.ok).toBe(false);
  expect(store.part).toBe(part);
  expect(store.revision).toBe(0);
  expect(count).toBe(0);
  expect(store.canUndo).toBe(false);
});
