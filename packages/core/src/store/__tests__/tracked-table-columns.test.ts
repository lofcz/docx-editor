// Tracked column insertion and deletion through Word cell and table-property revisions.

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
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width: number, extra = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${p(text)}</w:tc>`;
const grid = (widths: readonly number[]) =>
  `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
const tblPr = (width: number) =>
  `<w:tblPr><w:tblW w:w="${width}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>`;
/** Three columns, a horizontal span in the first row, and a vertical merge in the last column. */
const merged = () =>
  `<w:tbl>${tblPr(7200)}${grid([1200, 2400, 3600])}` +
  `<w:tr>${cell('ab', 3600, '<w:gridSpan w:val="2"/>')}${cell('c', 3600, '<w:vMerge w:val="restart"/>')}</w:tr>` +
  `<w:tr>${cell('d', 1200)}${cell('e', 2400)}${cell('', 3600, '<w:vMerge/>')}</w:tr>` +
  `</w:tbl>${p('Tail')}`;
const regular = () =>
  `<w:tbl>${tblPr(7200)}${grid([1200, 2400, 3600])}` +
  `<w:tr>${cell('a', 1200)}${cell('b', 2400)}${cell('c', 3600)}</w:tr>` +
  `<w:tr><w:tc>${p('d')}</w:tc><w:tc>${p('e')}</w:tc><w:tc>${p('f')}</w:tc></w:tr>` +
  `</w:tbl>${p('Tail')}`;

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
const kids = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);
const table = (part: OoxmlPart) => kids(kids(part.root)[0]!).find((node) => node.kind === 'table')!;
const gridColumns = (part: OoxmlPart) =>
  kids(
    kids(table(part)).find((node) => node.kind !== 'textValue' && node.localName === 'tblGrid')!
  ).filter((node) => node.kind !== 'textValue' && node.localName === 'gridCol');
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
const comparable = (part: OoxmlPart) =>
  serializeOoxmlPart(part).replace(/ w14:(paraId|textId)="[0-9A-F]+"/g, '');
const texts = (part: OoxmlPart) =>
  [...serializeOoxmlPart(part).matchAll(/<w:tr\b.*?<\/w:tr>/g)].map((row) =>
    [...row[0].matchAll(/<w:tc>.*?<\/w:tc>/g)].map((c) =>
      [...c[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('')
    )
  );

for (const [name, body] of [
  ['regular', regular],
  ['merged', merged],
] as const)
  for (const where of ['left', 'right'] as const)
    test(`${name}: tracked insertion ${where} of the middle column rejects exactly and accepts to the direct result`, () => {
      const part = load(body());
      const column = gridColumns(part)[1]!;
      const op = {
        op: 'insertTableColumn',
        tableId: table(part).id,
        where,
        gridColumnId: column.id,
      } as const;
      expect(validateTreeOp(part, { ...op, revision })).toBeNull();
      const pending = apply(part, { ...op, revision });
      const xml = serializeOoxmlPart(pending);
      expect(xml).toContain('<w:tblGridChange');
      expect(xml).toContain('<w:cellIns ');
      expect(xml.match(/<w:tcPrChange /g)?.length).toBeGreaterThan(0);
      const direct = apply(part, op);
      expect(gridColumns(pending)).toHaveLength(4);
      expect(canonicalOoxmlFingerprint(decide(pending, 'reject'))).toBe(
        canonicalOoxmlFingerprint(part)
      );
      expect(comparable(decide(pending, 'accept'))).toBe(comparable(direct));
    });

test('cells without property holders regain exact absence on both decisions', () => {
  const part = load(regular());
  const op = {
    op: 'insertTableColumn',
    tableId: table(part).id,
    where: 'right',
    gridColumnId: gridColumns(part)[2]!.id,
  } as const;
  const pending = apply(part, { ...op, revision });
  expect(serializeOoxmlPart(pending)).toContain('<w:tc><w:tcPr><w:tcPrChange ');
  const rejected = decide(pending, 'reject');
  expect(canonicalOoxmlFingerprint(rejected)).toBe(canonicalOoxmlFingerprint(part));
  expect(serializeOoxmlPart(rejected)).not.toContain('<w:tcPr/>');
  expect(comparable(decide(pending, 'accept'))).toBe(comparable(apply(part, op)));
});

for (const [name, body] of [
  ['regular', regular],
  ['merged', merged],
] as const)
  test(`${name}: tracked deletion of the last column keeps cells pending, rejects exactly, and accepts to a valid grid`, () => {
    const part = load(body());
    const columns = gridColumns(part);
    const op = {
      op: 'deleteTableColumn',
      tableId: table(part).id,
      gridColumnId: columns[2]!.id,
    } as const;
    const pending = apply(part, { ...op, revision });
    expect(serializeOoxmlPart(pending).match(/<w:cellDel /g)).toHaveLength(2);
    expect(gridColumns(pending)).toHaveLength(3);
    expect(canonicalOoxmlFingerprint(decide(pending, 'reject'))).toBe(
      canonicalOoxmlFingerprint(part)
    );
    const accepted = decide(pending, 'accept');
    expect(gridColumns(accepted)).toHaveLength(2);
    expect(texts(accepted)).toEqual(
      name === 'regular'
        ? [
            ['a', 'b'],
            ['d', 'e'],
          ]
        : [['ab'], ['d', 'e']]
    );
  });

test('regular middle-column deletion accepts to Word width transfer with expected text', () => {
  const part = load(regular());
  const pending = apply(part, {
    op: 'deleteTableColumn',
    tableId: table(part).id,
    gridColumnId: gridColumns(part)[1]!.id,
    revision,
  });
  const accepted = decide(pending, 'accept');
  expect(texts(accepted)).toEqual([
    ['a', 'c'],
    ['d', 'f'],
  ]);
  expect(gridColumns(accepted)).toHaveLength(2);
  expect(canonicalOoxmlFingerprint(decide(pending, 'reject'))).toBe(
    canonicalOoxmlFingerprint(part)
  );
});

test('tracked column changes refuse spans, stacked histories, and invalid attribution', () => {
  const part = load(merged());
  const id = table(part).id,
    columns = gridColumns(part);
  const through = {
    op: 'deleteTableColumn',
    tableId: id,
    gridColumnId: columns[0]!.id,
    revision,
  } as const;
  expect(validateTreeOp(part, through)).toBe('unsupported');
  expect(applyTreeOp(part, through)).toMatchObject({ ok: false, reason: 'unsupported' });
  const pending = apply(part, {
    op: 'insertTableColumn',
    tableId: id,
    where: 'right',
    gridColumnId: columns[2]!.id,
    revision,
  });
  expect(
    applyTreeOp(pending, {
      op: 'deleteTableColumn',
      tableId: id,
      gridColumnId: gridColumns(pending)[1]!.id,
      revision,
    })
  ).toMatchObject({ ok: false, reason: 'unsupported' });
  expect(
    validateTreeOp(part, {
      op: 'insertTableColumn',
      tableId: id,
      where: 'right',
      gridColumnId: columns[2]!.id,
      revision: { author: ' ' },
    })
  ).toBe('invalidArgs');
});

test('an explicitly sized tracked insertion rejects exactly and accepts to insertion plus native sizing', () => {
  const part = load(merged());
  const id = table(part).id;
  const op = {
    op: 'insertTableColumn',
    tableId: id,
    where: 'right',
    gridColumnId: gridColumns(part)[1]!.id,
  } as const;
  const pending = apply(part, { ...op, revision, widthTwips: 1800 });
  expect(serializeOoxmlPart(pending)).toContain('<w:tblPrChange ');
  const inserted = apply(part, op);
  const widths = gridColumns(inserted).map((column, index) =>
    index === 2
      ? 1800
      : Number(
          (column.kind !== 'textValue' &&
            column.attributes.find((a) => a.localName === 'w')?.value) ||
            0
        )
  );
  const direct = apply(inserted, {
    op: 'setTableProperties',
    tableId: id,
    columnWidthsTwips: widths,
  });
  expect(canonicalOoxmlFingerprint(decide(pending, 'reject'))).toBe(
    canonicalOoxmlFingerprint(part)
  );
  expect(comparable(decide(pending, 'accept'))).toBe(comparable(direct));
  expect(validateTreeOp(part, { ...op, widthTwips: 1800 })).toBe('invalidArgs');
  expect(
    validateTreeOp(part, {
      op: 'deleteTableColumn',
      tableId: id,
      gridColumnId: gridColumns(part)[2]!.id,
      revision,
      widthTwips: 1,
    } as unknown as TreeDocOp)
  ).toBe('invalidArgs');
});
