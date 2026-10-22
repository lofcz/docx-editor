import type { SurfacePaintFrame } from '@docx-editor.dev/core/editor';
import { parentNodeOf } from '@docx-editor.dev/core/store';
import type { IndexedBlock } from '../blocks/block-index';
import type { NativeSlotLease } from './native-slots';

export interface NativeSlotMount {
  readonly block: IndexedBlock;
  readonly pageIndex: number;
  readonly element: HTMLElement;
  readonly slot: NativeSlotLease;
}
const PAINTED =
  '.docx-paragraph-fragment[data-paragraph-id],.docx-table-fragment[data-table-id],' +
  '.docx-table-row[data-row-id],.docx-table-cell[data-cell-id]';

/** Match canonical ancestry to native paint ownership, including nested table cells. */
export function attachNativeSlots(frame: SurfacePaintFrame, mounts: readonly NativeSlotMount[]) {
  const byControl = new Map<string, NativeSlotMount[]>();
  for (const mount of mounts) {
    const list = byControl.get(mount.block.controlId) ?? [];
    list.push(mount);
    byControl.set(mount.block.controlId, list);
  }
  const candidates: { element: HTMLElement; owner: NativeSlotMount | undefined }[] = [];
  const ownerOf = new Map<HTMLElement, NativeSlotMount | undefined>();
  for (const page of frame.pagesLayer.querySelectorAll<HTMLElement>(
    '.docx-page[data-materialized="true"]'
  )) {
    const pageIndex = Number(page.dataset.pageIndex);
    for (const element of page.querySelectorAll<HTMLElement>(PAINTED)) {
      const id =
        element.dataset.paragraphId ??
        element.dataset.tableId ??
        element.dataset.rowId ??
        element.dataset.cellId!;
      let node = parentNodeOf(frame.part, id),
        owner: NativeSlotMount | undefined;
      while (node) {
        owner = byControl.get(node.id)?.find((mount) => mount.pageIndex === pageIndex);
        if (owner) break;
        node = parentNodeOf(frame.part, node.id);
      }
      ownerOf.set(element, owner);
      candidates.push({ element, owner });
    }
  }
  const owned = new Map<NativeSlotMount, HTMLElement[]>();
  for (const { element, owner } of candidates) {
    if (!owner) continue;
    let parent = element.parentElement,
      carried = false;
    while (parent && parent !== frame.pagesLayer) {
      if (ownerOf.get(parent) === owner) {
        carried = true;
        break;
      }
      parent = parent.parentElement;
    }
    if (carried) continue; // A whole table carries its own rows, cells, and paragraphs.
    const list = owned.get(owner) ?? [];
    list.push(element);
    owned.set(owner, list);
  }
  const parentOf = (mount: NativeSlotMount): NativeSlotMount | undefined => {
    let node = parentNodeOf(frame.part, mount.block.controlId);
    while (node) {
      const parent = byControl.get(node.id)?.find((value) => value.pageIndex === mount.pageIndex);
      if (parent) return parent;
      node = parentNodeOf(frame.part, node.id);
    }
    return undefined;
  };
  // An empty outer container still wraps its registered descendants.
  const anchors = new Map<NativeSlotMount, HTMLElement>();
  for (const candidate of candidates) {
    let mount = candidate.owner;
    while (mount) {
      if (!anchors.has(mount)) anchors.set(mount, candidate.element);
      mount = parentOf(mount);
    }
  }
  // Place every root before moving any native element. Wrappers preserve native coordinates.
  for (const mount of mounts) {
    const first = anchors.get(mount);
    if (first?.parentElement) first.before(mount.element);
  }
  for (const mount of mounts) {
    const parent = parentOf(mount);
    if (!parent || mount.element.parentElement !== parent.element.parentElement) continue;
    const list = owned.get(parent) ?? [];
    list.push(mount.element);
    owned.set(parent, list);
  }
  for (const mount of mounts) {
    const elements = owned.get(mount) ?? [];
    elements.sort((a, b) => (a.compareDocumentPosition(b) & 4 ? -1 : 1));
    mount.slot.set(elements);
  }
}
