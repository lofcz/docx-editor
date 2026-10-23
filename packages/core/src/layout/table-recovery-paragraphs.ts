// Project independently owned text outside a table row that the display mode removes.
// Keep the source wrappers and offsets. Only the row-owned paragraph mark leaves this view.
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
} from '../store/package/ooxml-tree.ts';
import { retainsNestedTable } from '../store/store/revision-table-preserve-nested.ts';
import { tableChildren } from '../store/store/revision-table-children.ts';
import { currentFieldResultsMode } from '../store/package/field-result-mode.ts';
import { piecesOfParagraph } from './field-projection.ts';
import { rowRevisionOmitted } from './table-cell-revision-view.ts';
import { cacheProjection } from './bounded-projection-cache.ts';
import {
  revisionNodeProjectionMode,
  revisionFilterSelects,
  type RevisionAuthorFilter,
  type RevisionDisplayMode,
} from './revision-projection.ts';

const recoveredByTable = new WeakMap<OoxmlElement, Map<string, readonly OoxmlElement[]>>();
const EMPTY: readonly OoxmlElement[] = [];

/** Foreign content has its own paragraph after the structural decision, in every story. */
export function tableRecoveryParagraphs(
  table: OoxmlElement,
  mode: RevisionDisplayMode,
  filter?: RevisionAuthorFilter
): readonly OoxmlElement[] {
  if (table.kind !== 'table' || (mode === 'all-markup' && !revisionFilterSelects(filter)))
    return EMPTY;
  const key = `${mode}|${filter?.cacheKey ?? ''}|${currentFieldResultsMode()}`;
  const cache = recoveredByTable.get(table);
  const cached = cache?.get(key);
  if (cached) return cached;
  const removes = (node: OoxmlNode): boolean => {
    if (
      node.kind === 'textValue' ||
      node.namespaceUri !== WML_NAMESPACE_URI ||
      !['ins', 'del'].includes(node.localName)
    )
      return false;
    const author =
      node.attributes.find(
        (attribute) =>
          attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'author'
      )?.value ?? '';
    const projected = filter ? revisionNodeProjectionMode(filter, node.id, author, mode) : mode;
    return (
      (node.localName === 'ins' && projected === 'original') ||
      (node.localName === 'del' && projected === 'proposed')
    );
  };
  const independent = (node: OoxmlNode): boolean =>
    node.kind !== 'textValue' && (removes(node) || node.children.some(independent));
  const recovered: OoxmlElement[] = [];
  const collect = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'paragraph') {
      if (!independent(node)) return;
      const projected = {
        ...node,
        children: node.children.map((properties) =>
          properties.kind === 'paragraphProperties'
            ? {
                ...properties,
                children: properties.children.map((holder) =>
                  holder.namespaceUri === WML_NAMESPACE_URI && holder.localName === 'rPr'
                    ? {
                        ...holder,
                        children: holder.children.filter((mark) => !removes(mark)),
                      }
                    : holder
                ),
              }
            : properties
        ),
      } as OoxmlElement;
      const pieces = piecesOfParagraph(
        projected,
        [],
        undefined,
        undefined,
        undefined,
        undefined,
        mode,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        filter
      );
      if (pieces.length) recovered.push(projected);
      return;
    }
    for (const child of node.children) collect(child);
  };
  for (const row of tableChildren(table, 'tableRow')) {
    const properties = row.children.find(
      (node) =>
        node.kind !== 'textValue' &&
        node.namespaceUri === WML_NAMESPACE_URI &&
        node.localName === 'trPr'
    );
    if (
      !rowRevisionOmitted(properties?.kind !== 'textValue' ? properties : undefined, mode, filter)
    )
      continue;
    const rejects =
      properties &&
      properties.kind !== 'textValue' &&
      properties.children.some(
        (mark) => mark.kind !== 'textValue' && mark.localName === 'ins' && removes(mark)
      );
    if (!retainsNestedTable(row, rejects ? 'reject' : 'accept')) collect(row);
  }
  if (cache) cacheProjection(cache, key, recovered);
  else recoveredByTable.set(table, new Map([[key, recovered]]));
  return recovered;
}
