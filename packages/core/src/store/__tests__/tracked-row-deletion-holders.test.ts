// A tracked row deletion must not leave property holders behind when it is rejected.

import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  serializeOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { applyTreeOp } from '../store/tree-ops.ts';

const W = WML_NAMESPACE_URI;
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const revision = { author: 'Reviewer', date: '2026-10-20T08:00:00Z' };

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
const kids = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);
function rows(part: OoxmlPart) {
  const table = kids(kids(part.root)[0]!).find((node) => node.kind === 'table')!;
  return { table, rows: kids(table).filter((node) => node.kind === 'tableRow') };
}
function apply(part: OoxmlPart, op: Parameters<typeof applyTreeOp>[1]): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}
function trackedDelete(part: OoxmlPart, index = 1): OoxmlPart {
  const { table, rows: list } = rows(part);
  return apply(part, { op: 'deleteTableRow', tableId: table.id, rowId: list[index]!.id, revision });
}
const bare = (a: string, b: string) => `<w:tr><w:tc>${p(a)}</w:tc><w:tc>${p(b)}</w:tc></w:tr>`;
const grid = '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>';

test('rejecting a tracked deletion of a row without trPr/tcPr restores the exact original row', () => {
  const part = load(`<w:tbl>${grid}${bare('A', 'B')}${bare('One', 'Two')}</w:tbl>${p('Tail')}`);
  const pending = trackedDelete(part);
  const xml = serializeOoxmlPart(pending);
  expect(xml).toContain('<w:trPr><w:del ');
  expect(xml.match(/<w:tcPr><w:cellDel /g)).toHaveLength(2);
  const rejected = apply(pending, { op: 'rejectAllRevisions' });
  expect(canonicalOoxmlFingerprint(rejected)).toBe(canonicalOoxmlFingerprint(part));
  expect(serializeOoxmlPart(rejected)).not.toMatch(/<w:trPr\/>|<w:tcPr\/>/);
  // Accepting still removes the row.
  const accepted = apply(pending, { op: 'acceptAllRevisions' });
  expect(rows(accepted).rows).toHaveLength(1);
});

test('rejection keeps imported holders, including an explicitly empty one', () => {
  const part = load(
    `<w:tbl>${grid}${bare('A', 'B')}<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr/>${p('One')}</w:tc><w:tc><w:tcPr><w:shd w:fill="EEEEEE"/></w:tcPr>${p('Two')}</w:tc></w:tr></w:tbl>${p('Tail')}`
  );
  const rejected = apply(trackedDelete(part), { op: 'rejectAllRevisions' });
  expect(canonicalOoxmlFingerprint(rejected)).toBe(canonicalOoxmlFingerprint(part));
  expect(serializeOoxmlPart(rejected)).toContain('<w:tcPr/>');
});

test('holders created by a deletion carry mark ids; a reopened file keeps emptied holders', () => {
  const part = load(`<w:tbl>${grid}${bare('A', 'B')}${bare('One', 'Two')}</w:tbl>${p('Tail')}`);
  const pending = trackedDelete(part);
  const row = rows(pending).rows[1]!;
  const holder = kids(row).find((node) => node.kind !== 'textValue' && node.localName === 'trPr')!;
  expect(holder.id).toContain('#mark:');
  // Saved bytes carry no allocation history, so reopened holders are source holders.
  const reopened = load(
    serializeOoxmlPart(pending)
      .replace(/^.*?<w:body>/s, '')
      .replace(/<\/w:body>.*$/s, '')
  );
  const rejected = apply(reopened, { op: 'rejectAllRevisions' });
  expect(serializeOoxmlPart(rejected)).toContain('<w:trPr/>');
});

test('a tracked row insertion keeps its fresh holders when accepted', () => {
  const part = load(`<w:tbl>${grid}${bare('A', 'B')}</w:tbl>${p('Tail')}`);
  const { table, rows: list } = rows(part);
  const pending = apply(part, {
    op: 'insertTableRow',
    tableId: table.id,
    rowId: list[0]!.id,
    where: 'below',
    revision,
  });
  const direct = apply(part, {
    op: 'insertTableRow',
    tableId: table.id,
    rowId: list[0]!.id,
    where: 'below',
  });
  const accepted = apply(pending, { op: 'acceptAllRevisions' });
  const strip = (value: string) => value.replace(/ w14:(paraId|textId)="[0-9A-F]+"/g, '');
  expect(strip(serializeOoxmlPart(accepted))).toBe(strip(serializeOoxmlPart(direct)));
});
