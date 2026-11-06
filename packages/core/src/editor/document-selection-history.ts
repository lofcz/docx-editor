import type { SemanticSelection } from '../layout/semantic-interaction.ts';
import { findNode, paraIdOf } from '../store/index.ts';
import type { OoxmlPart } from '../store/package/ooxml-tree.ts';
import { undoHistoryPosition } from '../store/store/undo-history-position.ts';

/** Resolve a history mark in the active story without choosing a DOM selection owner. */
export function selectionAfterHistory(
  mark: { paragraphId: string; start: number; end: number } | SemanticSelection | null,
  current: SemanticSelection,
  order: readonly string[],
  shared: boolean,
  clamp: (selection: SemanticSelection) => SemanticSelection
): SemanticSelection {
  if (!mark) return clamp(current);
  if ('anchor' in mark) {
    const local = !shared && [mark.anchor, mark.head].every((p) => order.includes(p.paragraphId));
    return local ? mark : clamp(mark);
  }
  const restored = {
    anchor: { paragraphId: mark.paragraphId, offset: mark.start },
    head: { paragraphId: mark.paragraphId, offset: mark.end },
  };
  // Another story's mark stays outside the active editable region.
  return order.includes(mark.paragraphId) ? restored : clamp(restored);
}

interface SavedPosition {
  readonly nodeId: string;
  readonly stableId: string | null;
  readonly offset: number;
}
interface SavedSelection {
  readonly partName: string;
  readonly anchor: SavedPosition;
  readonly head: SavedPosition;
}
interface Entry {
  readonly before: SavedSelection;
  readonly after: SavedSelection;
}

/** Local selection metadata for the document's existing undo items. */
export class DocumentSelectionHistory {
  private readonly entries = new WeakMap<object, Entry>();
  private readonly position;
  constructor(
    owner: object | undefined,
    private readonly selection: () => SemanticSelection,
    private readonly part: () => OoxmlPart,
    private readonly revision: () => number
  ) {
    this.position = owner ? undoHistoryPosition(owner) : undefined;
  }

  private save(): SavedSelection {
    const part = this.part();
    const position = (value: SemanticSelection['head']): SavedPosition => {
      const node = findNode(part, value.paragraphId);
      return {
        nodeId: value.paragraphId,
        stableId: node && node.kind !== 'textValue' ? paraIdOf(node) : null,
        offset: value.offset,
      };
    };
    const selection = this.selection();
    return {
      partName: part.name,
      anchor: position(selection.anchor),
      head: position(selection.head),
    };
  }

  /** A grouped edit retains its first selection and updates its final selection. */
  begin(): () => void {
    if (!this.position) return () => {};
    const revision = this.revision();
    const before = this.save();
    return () => {
      const key = this.position!.current();
      if (!key || revision === this.revision()) return;
      this.entries.set(key, {
        before: this.entries.get(key)?.before ?? before,
        after: this.save(),
      });
    };
  }

  step(direction: 'undo' | 'redo', run: () => boolean) {
    const undoKey = this.position?.current();
    const changed = run();
    const key = direction === 'undo' ? undoKey : this.position?.current();
    const entry = key ? this.entries.get(key) : undefined;
    const saved = direction === 'undo' ? entry?.before : entry?.after;
    return { changed, selection: saved ? this.resolve(saved) : null };
  }

  private resolve(saved: SavedSelection): SemanticSelection | null {
    const part = this.part();
    if (part.name !== saved.partName) return null;
    const ids = new Map<string, string>();
    const pending = [part.root];
    while (pending.length) {
      const node = pending.pop()!;
      for (const child of node.children) {
        if (child.kind !== 'textValue') pending.push(child);
      }
      if (node.kind !== 'paragraph') continue;
      const stableId = paraIdOf(node);
      if (stableId) ids.set(stableId, node.id);
    }
    const position = (value: SavedPosition) => {
      const paragraphId = value.stableId ? ids.get(value.stableId) : value.nodeId;
      return paragraphId ? { paragraphId, offset: value.offset } : null;
    };
    const anchor = position(saved.anchor);
    const head = position(saved.head);
    return anchor && head ? { anchor, head } : null;
  }
}
