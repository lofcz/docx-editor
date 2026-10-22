import { expect, test } from 'bun:test';
import {
  readOoxmlPart,
  canonicalOoxmlFingerprint,
  serializeOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
} from '../package/ooxml-tree';
import { findNode } from '../package/ooxml-edit';
import { applyTreeOp } from '../store/tree-ops';
import type { TreeDocOp } from '../store/tree-op-types';

const child = (node: OoxmlNode, name: string): OoxmlNode | undefined =>
  node.kind === 'textValue'
    ? undefined
    : node.children.find(
        (entry) =>
          entry.kind !== 'textValue' &&
          entry.namespaceUri === WML_NAMESPACE_URI &&
          entry.localName === name
      );
const attr = (node: OoxmlNode, name: string, ns = WML_NAMESPACE_URI) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find((entry) => entry.namespaceUri === ns && entry.localName === name)?.value;
function fixture() {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}" xmlns:k="urn:keep"><w:body><w:tbl><w:tblPr><w:tblW w:w="3600" w:type="dxa" k:tag="table-width"><k:inside/></w:tblW><w:tblLayout w:type="fixed" k:tag="layout"><k:inside/></w:tblLayout><k:tcW k:tag="foreign-name"/></w:tblPr><w:tblGrid><w:gridCol w:w="1200" k:tag="column"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="400" w:hRule="atLeast" k:tag="row-height"><k:inside/></w:trHeight><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa" k:tag="cell-width"><k:inside/></w:tcW><w:shd w:fill="EEEEEE"/><w:vAlign w:val="top" k:tag="alignment"><k:inside/></w:vAlign></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Keep A</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa" k:tag="other-width"/></w:tcPr><w:p><w:r><w:t>Keep B</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!result.ok) throw new Error(result.reason);
  const part = result.part,
    body = child(part.root, 'body')!,
    table = child(body, 'tbl')!,
    row = child(table, 'tr')!,
    cell = child(row, 'tc')!,
    grid = child(table, 'tblGrid')!;
  const columns = grid.kind === 'textValue' ? [] : grid.children;
  return { part, table, row, cell, columns };
}

for (const kind of ['divider', 'right-edge', 'row-height', 'full-grid', 'alignment'] as const)
  test(`${kind} preserves imported attributes, children, and identities on modified table properties`, () => {
    const { part, table, row, cell, columns } = fixture(),
      before = canonicalOoxmlFingerprint(part);
    const op: TreeDocOp =
      kind === 'divider'
        ? {
            op: 'setTableColumnWidths',
            tableId: table.id,
            leftGridColumnId: columns[0]!.id,
            rightGridColumnId: columns[1]!.id,
            leftWidthTwips: 1800,
            rightWidthTwips: 1800,
          }
        : kind === 'right-edge'
          ? {
              op: 'setTableRightEdgeWidth',
              tableId: table.id,
              gridColumnId: columns[1]!.id,
              columnWidthTwips: 3000,
              tableWidthTwips: 4200,
            }
          : kind === 'row-height'
            ? { op: 'setTableRowHeight', tableId: table.id, rowId: row.id, heightTwips: 600 }
            : kind === 'alignment'
              ? {
                  op: 'setTableCellVerticalAlignment',
                  tableId: table.id,
                  cellIds: [cell.id],
                  alignment: 'center',
                }
              : { op: 'setTableProperties', tableId: table.id, columnWidthsTwips: [1500, 2700] };
    const result = applyTreeOp(part, op);
    if (!result.ok) throw new Error(result.reason);
    const properties = [
      child(child(table, 'tblPr')!, 'tblW')!,
      child(child(table, 'tblPr')!, 'tblLayout')!,
      child(child(row, 'trPr')!, 'trHeight')!,
      child(child(cell, 'tcPr')!, 'tcW')!,
      child(child(cell, 'tcPr')!, 'vAlign')!,
    ];
    for (const old of properties) {
      const next = findNode(result.part, old.id);
      expect(next).not.toBeNull();
      expect(attr(next!, 'tag', 'urn:keep')).toBe(attr(old, 'tag', 'urn:keep'));
      if (old.kind !== 'textValue' && next?.kind !== 'textValue')
        expect(next!.children).toEqual(old.children);
    }
    const changedRow = child(child(findNode(result.part, row.id)!, 'trPr')!, 'trHeight')!;
    expect(attr(changedRow, 'val')).toBe(kind === 'row-height' ? '600' : '400');
    expect(attr(changedRow, 'hRule')).toBe(kind === 'row-height' ? 'exact' : 'atLeast');
    expect(attr(child(child(findNode(result.part, cell.id)!, 'tcPr')!, 'vAlign')!, 'val')).toBe(
      kind === 'alignment' ? 'center' : 'top'
    );
    const reopened = readOoxmlPart(serializeOoxmlPart(result.part), {
      name: part.name,
      contentType: part.contentType,
    });
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(canonicalOoxmlFingerprint(reopened.part)).toBe(canonicalOoxmlFingerprint(result.part));
    expect(serializeOoxmlPart(result.part)).toContain('Keep A');
    expect(serializeOoxmlPart(result.part)).toContain('Keep B');
    expect(canonicalOoxmlFingerprint(part)).toBe(before);
  });
