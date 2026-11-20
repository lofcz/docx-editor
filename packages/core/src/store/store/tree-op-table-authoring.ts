// Lossless table style, repeated headers, and grid widths; merged cells take their spans' sum.
import { createNodeIdAllocator, replaceNode, type EditOptions } from '../package/ooxml-edit.ts';
import { wmlFreshNamespaceContextAt } from '../package/wml-namespace.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { fromEdit, TEXT_DEPS } from './tree-op-nodes.ts';
import {
  patchTblPrChild,
  patchTrPrChild,
  patchTcPrChild,
  removeTrPrChild,
} from './tree-op-table-properties.ts';
import { readTableLook, tableLookAttributes } from '../table-look.ts';
import { removeTblPrChild } from './tree-op-table-properties.ts';
import { isWmlElement, wmlChildNamed } from './tree-op-table-shared.ts';
import { readEditableTableTopology } from './tree-op-table-topology.ts';
import { MIN_TABLE_COLUMN_WIDTH_TWIPS } from './table-constraints.ts';
import { resizedGridCells, validateTableResizeGrid } from './table-resize-grid.ts';
import type { TreeDocOp, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

type Op = Extract<TreeDocOp, { op: 'setTableProperties' }>;

export function validateTableProperties(part: OoxmlPart, op: Op): TreeOpRejection | null {
  if (
    !Object.keys(op).every((key) =>
      [
        'op',
        'tableId',
        'styleId',
        'headerRowCount',
        'columnWidthsTwips',
        'indentTwips',
        'autoFit',
        'styleOptions',
        'preferredWidth',
        'justification',
      ].includes(key)
    )
  )
    return 'invalidArgs';
  const preferred = op.preferredWidth;
  if (
    preferred !== undefined &&
    preferred !== null &&
    (!['pct', 'dxa', 'auto'].includes(preferred.type) ||
      !Number.isInteger(preferred.value) ||
      preferred.value < 0 ||
      preferred.value > (preferred.type === 'auto' ? 0 : preferred.type === 'pct' ? 5000 : 31680))
  )
    return 'invalidArgs';
  if (
    op.autoFit !== undefined &&
    (!['contents', 'window', 'fixed'].includes(op.autoFit) ||
      preferred !== undefined ||
      (op.autoFit === 'fixed') !== (op.columnWidthsTwips !== undefined))
  )
    return 'invalidArgs';
  if (op.justification != null && !['left', 'center', 'right'].includes(op.justification))
    return 'invalidArgs';
  if (
    op.styleOptions !== undefined &&
    (!op.styleOptions ||
      typeof op.styleOptions !== 'object' ||
      !Object.keys(op.styleOptions).length ||
      Object.entries(op.styleOptions).some(
        ([key, value]) =>
          !Object.hasOwn(readTableLook(undefined), key) || typeof value !== 'boolean'
      ))
  )
    return 'invalidArgs';
  if (typeof op.tableId !== 'string') return 'invalidArgs';
  const result = readEditableTableTopology(part.root, op.tableId);
  if (!result.ok) return result.reason === 'duplicate-node-id' ? 'unknown-table' : result.reason;
  const { rows, gridColumns, hasMerge } = result.topology;
  if (op.styleId !== undefined && (typeof op.styleId !== 'string' || op.styleId.length > 253))
    return 'invalidArgs';
  if (
    op.headerRowCount !== undefined &&
    (!Number.isInteger(op.headerRowCount) ||
      op.headerRowCount < 0 ||
      op.headerRowCount > rows.length)
  )
    return 'invalidArgs';
  if (
    op.indentTwips !== undefined &&
    (!Number.isInteger(op.indentTwips) || Math.abs(op.indentTwips) > 31680)
  )
    return 'invalidArgs';
  if (op.columnWidthsTwips !== undefined) {
    // Merged cells take the sum of the columns they span, so the grid must place them.
    const merged = hasMerge ? validateTableResizeGrid(result.topology) : null;
    if (merged) return merged;
    if (
      !Array.isArray(op.columnWidthsTwips) ||
      op.columnWidthsTwips.length !== gridColumns.length ||
      gridColumns.length === 0 ||
      (!hasMerge && rows.some(({ cells }) => cells.length !== gridColumns.length))
    )
      return 'invalidArgs';
    if (
      op.columnWidthsTwips.some(
        (width) =>
          !Number.isInteger(width) ||
          width < (op.autoFit === 'fixed' ? 1 : MIN_TABLE_COLUMN_WIDTH_TWIPS) ||
          width > 31680
      )
    )
      return 'invalidArgs';
    if (op.columnWidthsTwips.reduce((a, b) => a + b, 0) > 31680) return 'invalidArgs';
  }
  return null;
}

export function applyTableProperties(part: OoxmlPart, op: Op, options?: EditOptions): TreeOpResult {
  const invalid = validateTableProperties(part, op);
  if (invalid) return { ok: false, reason: invalid };
  const result = readEditableTableTopology(part.root, op.tableId);
  if (!result.ok) return { ok: false, reason: 'unknown-table' };
  const { table, grid, rows, gridColumns } = result.topology;
  const nextId = createNodeIdAllocator(part);
  const wml = wmlFreshNamespaceContextAt(part, table);
  const fresh = (
    localName: string,
    attrs: Record<string, string> = {},
    children: readonly OoxmlNode[] = []
  ): OoxmlElement =>
    ({
      id: nextId(),
      kind: 'generic',
      namespaceUri: WML_NAMESPACE_URI,
      localName,
      ...(wml.elementPrefix === undefined ? {} : { prefix: wml.elementPrefix }),
      namespaceBindings: wml.rowBinding ? [wml.rowBinding] : [],
      attributes: Object.entries(attrs).map(([localName, value]) => ({
        kind: 'genericExtension',
        namespaceUri: WML_NAMESPACE_URI,
        localName,
        prefix: wml.attributePrefix,
        value,
      })),
      children,
    }) as OoxmlElement;
  // Replace only the named leaves, preserving foreign attributes and sibling properties.
  const leaf = (
    parent: OoxmlElement,
    name: string,
    attrs: Record<string, string>
  ): OoxmlElement => {
    const old = wmlChildNamed(parent, name);
    if (!old) return fresh(name, attrs);
    return {
      ...old,
      attributes: [
        ...old.attributes.filter(
          (a) => a.namespaceUri !== WML_NAMESPACE_URI || !(a.localName in attrs)
        ),
        ...fresh(name, attrs).attributes,
      ],
    } as OoxmlElement;
  };
  const patch = (
    parent: OoxmlElement,
    containerName: string,
    property: OoxmlElement
  ): OoxmlElement => {
    const old = wmlChildNamed(parent, containerName);
    const container = old ?? fresh(containerName);
    const patched = (
      containerName === 'tblPr'
        ? patchTblPrChild
        : containerName === 'trPr'
          ? patchTrPrChild
          : patchTcPrChild
    )(container, property);
    if (!patched.ok) throw new Error('validated WML property');
    return {
      ...parent,
      children: old
        ? parent.children.map((n) => (n.id === old.id ? patched.container : n))
        : [patched.container, ...parent.children],
    } as OoxmlElement;
  };
  let updated: OoxmlElement = table;
  const updateTableLeaf = (name: string, attrs: Record<string, string> | null) => {
    const props = wmlChildNamed(updated, 'tblPr');
    if (attrs) updated = patch(updated, 'tblPr', leaf(props ?? fresh('tblPr'), name, attrs));
    else if (props) {
      const removed = removeTblPrChild(props, name);
      if (removed.ok)
        updated = {
          ...updated,
          children: updated.children.map((n) => (n.id === props.id ? removed.container : n)),
        } as OoxmlElement;
    }
  };
  if (op.preferredWidth !== undefined)
    updateTableLeaf(
      'tblW',
      op.preferredWidth && { type: op.preferredWidth.type, w: String(op.preferredWidth.value) }
    );
  if (op.justification !== undefined)
    updateTableLeaf('jc', op.justification && { val: op.justification });
  if (op.styleOptions)
    updateTableLeaf(
      'tblLook',
      tableLookAttributes({ ...readTableLook(wmlChildNamed(table, 'tblPr')), ...op.styleOptions })
    );
  const fitContent = op.autoFit === 'contents' || op.autoFit === 'window';
  if (fitContent) {
    updateTableLeaf('tblLayout', { type: 'autofit' });
    updateTableLeaf(
      'tblW',
      op.autoFit === 'contents' ? { type: 'auto', w: '0' } : { type: 'pct', w: '5000' }
    );
  }
  if (op.styleId !== undefined)
    updateTableLeaf('tblStyle', op.styleId ? { val: op.styleId } : null);
  if (op.indentTwips !== undefined)
    updated = patch(
      updated,
      'tblPr',
      leaf(wmlChildNamed(updated, 'tblPr') ?? fresh('tblPr'), 'tblInd', {
        w: String(op.indentTwips),
        type: 'dxa',
      })
    );
  const replacements = new Map<string, OoxmlNode>();
  const cellWidths = new Map<string, number>();
  if (op.columnWidthsTwips) {
    const widths = new Map(op.columnWidthsTwips.map((width, index) => [index, width]));
    for (const { cell, width } of resizedGridCells(result.topology, widths)) {
      cellWidths.set(cell.id, width);
    }
  }
  rows.forEach(({ row, cells }, rowIndex) => {
    let edited: OoxmlElement = row;
    if (op.headerRowCount !== undefined) {
      if (rowIndex < op.headerRowCount)
        edited = patch(
          edited,
          'trPr',
          leaf(wmlChildNamed(edited, 'trPr') ?? fresh('trPr'), 'tblHeader', { val: '1' })
        );
      else {
        const props = wmlChildNamed(edited, 'trPr');
        if (props) {
          const removed = removeTrPrChild(props, 'tblHeader');
          if (removed.ok)
            edited = {
              ...edited,
              children: edited.children.map((n) => (n.id === props.id ? removed.container : n)),
            } as OoxmlElement;
        }
      }
    }
    if (op.columnWidthsTwips || op.autoFit) {
      const cellMap = new Map(
        cells.map((cell) => [
          cell.id,
          patch(
            cell,
            'tcPr',
            leaf(wmlChildNamed(cell, 'tcPr') ?? fresh('tcPr'), 'tcW', {
              w: fitContent ? '0' : String(cellWidths.get(cell.id)),
              type: fitContent ? 'auto' : 'dxa',
            })
          ),
        ])
      );
      edited = {
        ...edited,
        children: edited.children.map((n) => cellMap.get(n.id) ?? n),
      } as OoxmlElement;
    }
    replacements.set(row.id, edited);
  });
  if (op.columnWidthsTwips && grid) {
    const widths = new Map(
      gridColumns.map((column, index) => [column.id, String(op.columnWidthsTwips![index])])
    );
    replacements.set(grid.id, {
      ...grid,
      children: grid.children.map((node) =>
        isWmlElement(node, 'gridCol') && widths.has(node.id)
          ? ({
              ...node,
              attributes: [
                ...node.attributes.filter(
                  (a) => a.namespaceUri !== WML_NAMESPACE_URI || a.localName !== 'w'
                ),
                ...fresh('gridCol', { w: widths.get(node.id)! }).attributes,
              ],
            } as OoxmlElement)
          : node
      ),
    } as OoxmlElement);
    const props = wmlChildNamed(updated, 'tblPr') ?? fresh('tblPr');
    updated = patch(
      updated,
      'tblPr',
      leaf(props, 'tblW', {
        w: String(op.columnWidthsTwips.reduce((a, b) => a + b, 0)),
        type: 'dxa',
      })
    );
    updated = patch(
      updated,
      'tblPr',
      leaf(wmlChildNamed(updated, 'tblPr')!, 'tblLayout', { type: 'fixed' })
    );
  }
  updated = {
    ...updated,
    children: updated.children.map((node) => replacements.get(node.id) ?? node),
  } as OoxmlElement;
  return fromEdit(replaceNode(part, table.id, updated, options), {
    dirty: [table.id],
    created: [],
    deleted: [],
    dependencyKeys: TEXT_DEPS,
    impact: 'flow-structural',
  });
}
