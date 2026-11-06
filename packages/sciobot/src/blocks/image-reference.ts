import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { findNode } from '@docx-editor.dev/core/store';

/** A disposable address for one exact drawing in one loaded document. */
export interface NativeImageReference {
  /** Null after deletion, document replacement, or disposal. A deleted reference never revives. */
  read(): string | null;
  dispose(): void;
}

/** Retain only replica identity mappings published before the session retires local aliases. */
export function retainNativeImageReference(
  editor: DocxEditorInstance,
  drawingNodeId: string
): NativeImageReference {
  const owner = editor.surface;
  const session = owner?.session;
  const parts = session
    ? [...session.currentPackage().parts.values()].filter(
        (part) => findNode(part, drawingNodeId)?.kind === 'drawing'
      )
    : [];
  const partName = parts.length === 1 ? parts[0]!.name : null;
  let id: string | null = partName ? drawingNodeId : null;
  let offModel = () => {};
  let offEditor = () => {};
  const dispose = () => {
    id = null;
    offModel();
    offEditor();
  };
  const currentOwner = () => !!owner && editor.surface === owner && owner.session === session;
  const validate = () => {
    if (!currentOwner()) dispose();
  };
  if (id && session && owner && partName) {
    offModel = session.subscribe(() => {
      validate();
      if (!id) return;
      const part = session.currentPackage().parts.get(partName);
      // Eager publication can retain the source address even while an alias exists.
      if (part && findNode(part, id)?.kind === 'drawing') return;
      const mapped = owner.collaborationSession()?.resolveNodeId?.(id);
      if (mapped && part && findNode(part, mapped)?.kind === 'drawing') id = mapped;
      else dispose();
    });
    offEditor = editor.on('change', validate);
  }
  return {
    read() {
      validate();
      return id;
    },
    dispose,
  };
}
