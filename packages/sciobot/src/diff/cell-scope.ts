import {
  findNode,
  parentNodeOf,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure } from '../contracts/document';
import { jsonDataProblem } from '../data/json';

export interface NativeCellScope {
  readonly tableId: string;
  readonly rowId: string;
  readonly cellId: string;
}
export function nearestCellAncestor(
  part: OoxmlPart,
  id: string,
  kind: OoxmlNode['kind']
): OoxmlNode | null {
  let current = parentNodeOf(part, id);
  for (let depth = 0; current && depth < 64; depth++) {
    if (current.kind === kind) return current;
    current = parentNodeOf(part, current.id);
  }
  return null;
}
const fail = (error: string, code: NativeEditFailure['code']): NativeEditFailure => ({
  success: false,
  code,
  error,
});

/** Read live canonical ownership and the authored continuation marker; never infer table geometry. */
export function checkNativeCellScope(
  part: OoxmlPart,
  scope: NativeCellScope
): NativeEditFailure | null {
  const problem = jsonDataProblem(scope, { maxDepth: 2, maxNodes: 10 });
  if (problem) return fail(problem, 'invalid-input');
  if (
    !scope ||
    Object.keys(scope).some((key) => !['tableId', 'rowId', 'cellId'].includes(key)) ||
    [scope.tableId, scope.rowId, scope.cellId].some((id) => typeof id !== 'string' || !id)
  )
    return fail('Invalid native cell scope.', 'invalid-input');
  const cell = findNode(part, scope.cellId),
    row = findNode(part, scope.rowId),
    table = findNode(part, scope.tableId);
  if (
    cell?.kind !== 'tableCell' ||
    row?.kind !== 'tableRow' ||
    table?.kind !== 'table' ||
    nearestCellAncestor(part, cell.id, 'tableRow')?.id !== row.id ||
    nearestCellAncestor(part, row.id, 'table')?.id !== table.id
  )
    return fail(
      'Table, row, and cell do not share the requested native ancestry.',
      'invalid-document'
    );
  const properties = cell.children.find(
    (node) => node.namespaceUri === WML_NAMESPACE_URI && node.localName === 'tcPr'
  );
  const merge = properties?.children.find(
    (node) =>
      node.kind !== 'textValue' &&
      node.namespaceUri === WML_NAMESPACE_URI &&
      node.localName === 'vMerge'
  );
  if (
    merge &&
    merge.kind !== 'textValue' &&
    merge.attributes.find(
      (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'val'
    )?.value !== 'restart'
  )
    return fail(
      'Edit the visible merged cell instead of its hidden continuation.',
      'unsupported-operation'
    );
  return null;
}
