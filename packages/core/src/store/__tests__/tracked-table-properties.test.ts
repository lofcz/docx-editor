// Tracked table property commands through Word property-change markup.

import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  serializeOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { planRevisionBatch, revisionItemsOf } from '../index.ts';
import { applyTreeOp, validateTreeOp, type TreeDocOp } from '../store/tree-ops.ts';

const W = WML_NAMESPACE_URI;
const revision = { author: 'Reviewer', date: '2026-10-20T08:00:00Z' };
const other = { author: 'Other', date: '2026-10-19T08:00:00Z' };
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width: number, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${p(text)}</w:tc>`;
/** The last cell has no `w:tcPr` and the last row no `w:trPr`, as Word allows. */
const BODY =
  `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
  `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell('a', 2000, '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>')}${cell('b', 4000, '<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="FF0000"/></w:tcBorders>')}</w:tr>` +
  `<w:tr>${cell('c', 2000)}<w:tc>${p('d')}</w:tc></w:tr></w:tbl>${p('Tail')}`;
function load(xml: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${xml}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
const reopen = (part: OoxmlPart) => {
  const result = readOoxmlPart(serializeOoxmlPart(part), {
    name: part.name,
    contentType: part.contentType,
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
};
const kids = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);
const table = (part: OoxmlPart) => kids(kids(part.root)[0]!).find((node) => node.kind === 'table')!;
const rows = (part: OoxmlPart) => kids(table(part)).filter((row) => row.kind === 'tableRow');
const cells = (part: OoxmlPart) =>
  rows(part).map((row) =>
    kids(row)
      .filter((node) => node.kind === 'tableCell')
      .map((node) => node.id)
  );
const gridIds = (part: OoxmlPart) =>
  kids(
    kids(table(part)).find(
      (node) => node.kind !== 'textValue' && 'localName' in node && node.localName === 'tblGrid'
    )!
  ).map((node) => node.id);
function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`${result.reason} ${'detail' in result ? result.detail : ''}`);
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
const comparable = (part: OoxmlPart) =>
  serializeOoxmlPart(part).replace(/ w14:(paraId|textId)="[0-9A-F]+"/g, '');
const xmlOf = (part: OoxmlPart) => serializeOoxmlPart(part);
const hollowFree = (part: OoxmlPart) => comparable(part).replace(/<w:(tcPr|trPr)\/>/g, '');
const textOf = (part: OoxmlPart) => xmlOf(part).replace(/<[^>]+>/g, '');

function exact(part: OoxmlPart, op: TreeDocOp, by = revision): OoxmlPart {
  const tracked = { ...op, revision: by } as TreeDocOp;
  expect(validateTreeOp(part, tracked)).toBeNull();
  const pending = apply(part, tracked);
  expect(revisionItemsOf(pending).length).toBeGreaterThan(0);
  expect(textOf(pending)).toBe(textOf(part));
  expect(canonicalOoxmlFingerprint(decide(pending, 'reject'))).toBe(
    canonicalOoxmlFingerprint(part)
  );
  expect(comparable(decide(pending, 'accept'))).toBe(comparable(apply(part, op)));
  // Saved bytes keep the proposal and both decisions. Ids carry no history in a reopened file,
  // so a holder the proposal created stays as an empty element there (as for tracked rows).
  const reopened = reopen(pending);
  expect(hollowFree(decide(reopened, 'reject'))).toBe(hollowFree(reopen(part)));
  expect(hollowFree(decide(reopened, 'accept'))).toBe(hollowFree(reopen(apply(part, op))));
  return pending;
}

const commands: [string, (part: OoxmlPart) => TreeDocOp][] = [
  [
    'cell fill',
    (part) =>
      ({
        op: 'setTableCellFill',
        tableId: table(part).id,
        cellIds: [cells(part)[0]![0]!],
        color: { kind: 'hex', value: '00FF00' },
      }) as TreeDocOp,
  ],
  [
    'cleared fill',
    (part) =>
      ({
        op: 'setTableCellFill',
        tableId: table(part).id,
        cellIds: [cells(part)[0]![0]!],
        color: null,
      }) as TreeDocOp,
  ],
  [
    'fill on a cell without properties',
    (part) =>
      ({
        op: 'setTableCellFill',
        tableId: table(part).id,
        cellIds: [cells(part)[1]![1]!],
        color: { kind: 'hex', value: '0000FF' },
      }) as TreeDocOp,
  ],
  [
    'borders',
    (part) =>
      ({
        op: 'setTableCellBorders',
        tableId: table(part).id,
        cellIds: [cells(part)[0]![1]!, cells(part)[1]![1]!],
        scope: 'all',
        spec: { style: 'double', size: 12, color: { kind: 'hex', value: '0000FF' } },
      }) as TreeDocOp,
  ],
  [
    'vertical alignment',
    (part) =>
      ({
        op: 'setTableCellVerticalAlignment',
        tableId: table(part).id,
        cellIds: cells(part).flat(),
        alignment: 'center',
      }) as TreeDocOp,
  ],
  [
    'header rows',
    (part) =>
      ({ op: 'setTableProperties', tableId: table(part).id, headerRowCount: 2 }) as TreeDocOp,
  ],
  [
    'column widths',
    (part) =>
      ({
        op: 'setTableProperties',
        tableId: table(part).id,
        columnWidthsTwips: [3000, 3000],
      }) as TreeDocOp,
  ],
  [
    'divider drag',
    (part) =>
      ({
        op: 'setTableColumnWidths',
        tableId: table(part).id,
        leftGridColumnId: gridIds(part)[0]!,
        rightGridColumnId: gridIds(part)[1]!,
        leftWidthTwips: 2500,
        rightWidthTwips: 3500,
      }) as TreeDocOp,
  ],
  [
    'right edge',
    (part) =>
      ({
        op: 'setTableRightEdgeWidth',
        tableId: table(part).id,
        gridColumnId: gridIds(part)[1]!,
        columnWidthTwips: 5000,
        tableWidthTwips: 7000,
      }) as TreeDocOp,
  ],
  [
    'row height on a row without properties',
    (part) =>
      ({
        op: 'setTableRowHeight',
        tableId: table(part).id,
        rowId: rows(part)[1]!.id,
        heightTwips: 720,
      }) as TreeDocOp,
  ],
];
for (const [name, build] of commands)
  test(`tracked ${name} rejects exactly and accepts to the direct command`, () => {
    const part = load(BODY);
    exact(part, build(part));
  });

test('records name the old values and the author', () => {
  const part = load(BODY);
  const pending = exact(part, commands[0]![1](part));
  const xml = xmlOf(pending);
  expect(xml).toMatch(
    /<w:tcPrChange w:author="Reviewer" w:date="2026-10-20T08:00:00Z" w:id="\d+"><w:tcPr><w:tcW w:type="dxa" w:w="2000"\/><w:shd w:color="auto" w:fill="FFFF00" w:val="clear"\/><\/w:tcPr><\/w:tcPrChange>/
  );
  // Rejecting a row record resets unrecorded cells, so the changed row records both cells.
  expect(xml.match(/<w:tcPrChange /g)).toHaveLength(2);
  const header = exact(part, commands[5]![1](part));
  expect(xmlOf(header).match(/<w:trPrChange /g)).toHaveLength(2);
  const widths = exact(part, commands[6]![1](part));
  expect(xmlOf(widths)).toContain('<w:tblGridChange');
  // A grid change records every cell, as a tracked column does.
  expect(xmlOf(widths).match(/<w:tcPrChange /g)).toHaveLength(4);
});

test('a second change keeps the original record; returning to it removes the record', () => {
  const part = load(BODY);
  const [, build] = commands[0]!;
  const cellId = cells(part)[0]![0]!;
  const first = apply(part, { ...build(part), revision } as TreeDocOp);
  const second = apply(first, {
    op: 'setTableCellFill',
    tableId: table(part).id,
    cellIds: [cellId],
    color: { kind: 'hex', value: 'ABCDEF' },
    revision,
  } as TreeDocOp);
  expect(xmlOf(second).match(/<w:tcPrChange /g)).toHaveLength(2);
  expect(xmlOf(second).match(/w:fill="FFFF00"/g)).toHaveLength(1);
  expect(xmlOf(second)).toContain('w:fill="FFFF00"');
  expect(canonicalOoxmlFingerprint(decide(second, 'reject'))).toBe(canonicalOoxmlFingerprint(part));
  const back = apply(second, {
    op: 'setTableCellFill',
    tableId: table(part).id,
    cellIds: [cellId],
    color: { kind: 'hex', value: 'FFFF00' },
    revision,
  } as TreeDocOp);
  expect(revisionItemsOf(back)).toHaveLength(0);
});

test("returning to the original keeps another author's record", () => {
  const part = load(BODY);
  const cellId = cells(part)[0]![0]!;
  const theirs = apply(part, {
    op: 'setTableCellFill',
    tableId: table(part).id,
    cellIds: [cellId],
    color: { kind: 'hex', value: '00FF00' },
    revision: other,
  } as TreeDocOp);
  const mine = apply(theirs, {
    op: 'setTableCellFill',
    tableId: table(part).id,
    cellIds: [cellId],
    color: { kind: 'hex', value: 'FFFF00' },
    revision,
  } as TreeDocOp);
  expect(xmlOf(mine)).toContain('w:author="Other"');
  expect(revisionItemsOf(mine).length).toBeGreaterThan(0);
});

test('property proposals refuse over pending structure and invalid attribution, atomically', () => {
  const part = load(BODY);
  const withRow = apply(part, {
    op: 'insertTableRow',
    tableId: table(part).id,
    rowId: rows(part)[0]!.id,
    where: 'below',
    revision,
  } as TreeDocOp);
  const fill = {
    op: 'setTableCellFill',
    tableId: table(part).id,
    cellIds: [cells(part)[0]![0]!],
    color: { kind: 'hex', value: '00FF00' },
  };
  expect(validateTreeOp(withRow, { ...fill, revision } as TreeDocOp)).toBe('unsupported');
  expect(applyTreeOp(withRow, { ...fill, revision } as TreeDocOp).ok).toBe(false);
  expect(validateTreeOp(part, { ...fill, revision: { author: '' } } as TreeDocOp)).toBe(
    'invalidArgs'
  );
  expect(validateTreeOp(part, { ...fill, revision, extra: 1 } as unknown as TreeDocOp)).toBe(
    'invalidArgs'
  );
  // An untracked command over the same table is unchanged by this lane.
  expect(validateTreeOp(withRow, fill as TreeDocOp)).toBeNull();
});

test('a no-op tracked command records nothing', () => {
  const part = load(BODY);
  const same = apply(part, {
    op: 'setTableCellFill',
    tableId: table(part).id,
    cellIds: [cells(part)[0]![0]!],
    color: { kind: 'hex', value: 'FFFF00' },
    revision,
  } as TreeDocOp);
  expect(revisionItemsOf(same)).toHaveLength(0);
});

test('reverting a proposal drops the record holders it created', () => {
  // The second cell of the first row has no properties of its own.
  const part = load(BODY.replace(/(<w:tc><w:tcPr><w:tcW w:w="4000"[^]*?<\/w:tcPr>)/, '<w:tc>'));
  expect(xmlOf(part)).toContain('<w:tc><w:p><w:r><w:t>b</w:t>');
  const a = cells(part)[0]![0]!;
  const fill = (value: string) =>
    ({
      op: 'setTableCellFill',
      tableId: table(part).id,
      cellIds: [a],
      color: { kind: 'hex', value },
      revision,
    }) as TreeDocOp;
  const proposed = apply(part, fill('0000FF'));
  expect(xmlOf(proposed).match(/<w:tcPrChange /g)).toHaveLength(2);
  const reverted = apply(proposed, fill('FFFF00'));
  expect(revisionItemsOf(reverted)).toHaveLength(0);
  expect(canonicalOoxmlFingerprint(reverted)).toBe(canonicalOoxmlFingerprint(part));
});
