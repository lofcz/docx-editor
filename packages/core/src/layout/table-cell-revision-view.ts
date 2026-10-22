// Which cells a resolved revision view shows.
//
// The Proposed view reads a pending `w:cellDel` as accepted and the Original view reads a pending
// `w:cellIns` as rejected, so each omits that cell. An accepted cell deletion gives its grid span
// to the previous surviving cell in the row (the next one when it was first), which is how the
// store's decision transfers it. The Original view gives nothing away: a tracked column insertion
// records the old grid and old cell properties, and their projection restores the original row.
import { WML_NAMESPACE_URI, type OoxmlElement, type OoxmlNode } from '@docx-editor.dev/core/store';
import { readGridSpan, readVMerge } from './table-row-properties.ts';
import { readPreferredWidth, type PreferredWidth } from './table-widths.ts';
import {
  revisionNodeProjectionMode,
  type RevisionAuthorFilter,
  type RevisionDisplayMode,
} from './revision-projection.ts';

export interface VisibleRowCells {
  readonly cells: readonly OoxmlNode[];
  /** The omitted cells whose grid span each surviving cell received. */
  readonly receivedFrom: ReadonlyMap<OoxmlNode, readonly OoxmlElement[]>;
}
const UNCHANGED_RECEIVED = new Map<OoxmlNode, never>();

const wml = (node: OoxmlNode, name: string): node is OoxmlElement =>
  node.kind !== 'textValue' && node.namespaceUri === WML_NAMESPACE_URI && node.localName === name;
function cellMarkers(cell: OoxmlNode): readonly OoxmlElement[] {
  if (cell.kind === 'textValue') return [];
  const properties = cell.children.find((child) => wml(child, 'tcPr'));
  if (!properties || properties.kind === 'textValue') return [];
  const children: readonly OoxmlNode[] = properties.children;
  return children.filter(
    (child): child is OoxmlElement => wml(child, 'cellIns') || wml(child, 'cellDel')
  );
}
/** Rejected insertion wins: an owner absent from Original donates no geometry. */
function structuralOmission(
  markers: readonly OoxmlElement[],
  mode: RevisionDisplayMode,
  filter: RevisionAuthorFilter | undefined
): 'insertion' | 'deletion' | null {
  let deletion = false;
  for (const marker of markers) {
    const author =
      marker.attributes.find(
        (attribute) =>
          attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'author'
      )?.value ?? '';
    const projected = filter ? revisionNodeProjectionMode(filter, marker.id, author, mode) : mode;
    if (projected === 'original' && (marker.localName === 'ins' || marker.localName === 'cellIns'))
      return 'insertion';
    if (projected === 'proposed' && (marker.localName === 'del' || marker.localName === 'cellDel'))
      deletion = true;
  }
  return deletion ? 'deletion' : null;
}
export function rowRevisionOmitted(
  properties: OoxmlElement | undefined,
  mode: RevisionDisplayMode,
  filter: RevisionAuthorFilter | undefined
): boolean {
  const markers = (properties?.children ?? []).filter(
    (node): node is OoxmlElement => wml(node, 'ins') || wml(node, 'del')
  );
  return structuralOmission(markers, mode, filter) !== null;
}

/** The row's visible cells in `mode`. Returns the input list when the view omits none. */
export function visibleRowCells(
  cells: readonly OoxmlNode[],
  mode: RevisionDisplayMode,
  filter: RevisionAuthorFilter | undefined
): VisibleRowCells {
  if (mode === 'all-markup') return { cells, receivedFrom: UNCHANGED_RECEIVED };
  let visible: OoxmlNode[] | null = null;
  const pending: OoxmlElement[] = [];
  const receivedFrom = new Map<OoxmlNode, OoxmlElement[]>();
  for (let index = 0; index < cells.length; index += 1) {
    const cell = cells[index]!;
    const omission = structuralOmission(
      cell.kind === 'tableCell' ? cellMarkers(cell) : [],
      mode,
      filter
    );
    const omitted = omission !== null;
    if (!omitted) {
      visible?.push(cell);
      if (pending.length && cell.kind === 'tableCell') {
        receivedFrom.set(cell, pending.splice(0));
      }
      continue;
    }
    visible ??= cells.slice(0, index);
    if (omission !== 'deletion') continue;
    const receiver = [...visible].reverse().find((entry) => entry.kind === 'tableCell');
    if (receiver)
      receivedFrom.set(receiver, [...(receivedFrom.get(receiver) ?? []), cell as OoxmlElement]);
    else pending.push(cell as OoxmlElement);
  }
  if (!visible) return { cells, receivedFrom: UNCHANGED_RECEIVED };
  return { cells: visible, receivedFrom };
}

function named(node: OoxmlElement | undefined, name: string): OoxmlElement | undefined {
  for (const child of node?.children ?? []) if (wml(child, name)) return child as OoxmlElement;
  return undefined;
}
const propertiesOf = (cell: OoxmlElement) => named(cell, 'tcPr');

/** Grid columns a surviving cell receives from omitted deleted cells. */
export function receivedGridSpan(donors: readonly OoxmlElement[] | undefined): number {
  let span = 0;
  for (const donor of donors ?? []) span += readGridSpan(propertiesOf(donor));
  return span;
}

/** A surviving cell's preferred width, plus received widths when both are in twips. */
export function withReceivedWidth(
  width: PreferredWidth,
  donors: readonly OoxmlElement[] | undefined
): PreferredWidth {
  let result = width;
  for (const donor of donors ?? []) {
    const properties = propertiesOf(donor);
    const received = readPreferredWidth(named(properties, 'tcW'));
    if (result.type === 'dxa' && received.type === 'dxa')
      result = { type: 'dxa', value: result.value + received.value };
  }
  return result;
}

/**
 * A cell's vertical merge in `mode`. A pending `w:cellMerge` reads its `w:vMerge` in Proposed and
 * its `w:vMergeOrig` in Original; All Markup keeps the authored `w:vMerge`.
 */
export function resolvedVMerge(
  properties: OoxmlElement | undefined,
  mode: RevisionDisplayMode,
  filter: RevisionAuthorFilter | undefined
): 'none' | 'restart' | 'continue' {
  const authored = readVMerge(properties);
  if (mode === 'all-markup') return authored;
  const marker = named(properties, 'cellMerge');
  if (!marker) return authored;
  const read = (name: string) =>
    marker.attributes.find(
      (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === name
    )?.value;
  const projected = filter
    ? revisionNodeProjectionMode(filter, marker.id, read('author') ?? '', mode)
    : mode;
  if (projected === 'all-markup') return authored;
  const value = read(projected === 'proposed' ? 'vMerge' : 'vMergeOrig');
  return value === 'rest' ? 'restart' : value === 'cont' ? 'continue' : 'none';
}
