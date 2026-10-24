/** An editor-local view action must stay between the same document undo entries. */
export interface UndoHistoryPosition {
  current(): object | null;
  /** End capture and discard the abandoned redo branch without changing document content. */
  split(): void;
}
const positions = new WeakMap<object, UndoHistoryPosition>();

/** Register a document history authority for editor-local undo actions. @internal */
export function registerUndoHistoryPosition(owner: object, position: UndoHistoryPosition): void {
  positions.set(owner, position);
}

/**
 * Transparent session wrappers can expose their native authority through this local symbol.
 * This permits DOM-free protocol wrappers without an editor runtime import. Nothing is replicated.
 */
const NATIVE_UNDO_AUTHORITY = Symbol.for('@docx-editor.dev/undo-history-authority');
export function undoHistoryPosition(owner: object): UndoHistoryPosition | undefined {
  const seen = new Set<object>();
  for (let depth = 0; depth < 32 && !seen.has(owner); depth++) {
    seen.add(owner);
    const position = positions.get(owner);
    if (position) return position;
    let native: unknown;
    try {
      native = Reflect.get(owner, NATIVE_UNDO_AUTHORITY);
    } catch {
      return undefined;
    }
    if ((!native || typeof native !== 'object') && typeof native !== 'function') return undefined;
    owner = native as object;
  }
  return undefined;
}

export function shareUndoHistoryPosition<T extends object>(owner: T, authority: object): T {
  const position = undoHistoryPosition(authority);
  if (position) positions.set(owner, position);
  return owner;
}
