import type { ComponentType, ReactNode } from 'react';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { IndexedBlock } from '../blocks/block-index';
import type { ReactBlockInset } from './block-insets';
import type { ReactBlockHostOptions, ReactBlockProviders } from './providers';
import { followEditorSurface } from './follow-editor-surface';
import { mountSurfaceBlockHost, validateBlockViewType } from './surface-host';

export interface ReactBlockFragment {
  readonly ordinal: number;
  readonly pageIndex: number;
  /** Native page-content coordinates, in points. */
  readonly box: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  };
  readonly scale: number;
  /** CSS transform applied by the host to this view. Divide page pixels by this for local chrome. */
  readonly contentScale?: number;
}
export interface ReactBlockViewProps {
  readonly block: IndexedBlock;
  readonly fragment: ReactBlockFragment;
  readonly revision: number;
  /** Read data only. Native text and payloads remain canonical document content. */
  readonly paragraphs: readonly { readonly id: string; readonly text: string }[];
  /** Native paint nodes. Render once inside a transparent (`display: contents`) container. */
  readonly children: ReactNode;
}
export type ReactBlockView = ComponentType<ReactBlockViewProps>;
export interface ReactBlockRegistration {
  readonly view: ReactBlockView;
  /** Borrow block, row, or cell content. Inline native slots use a separate adapter. */
  readonly nativeContent?: boolean;
  /**
   * Reserve native flow for a widget taller than its Word fallback. The widget's bounded
   * height is measured; following native content moves down without a document edit.
   * One opaque widget occupies a bounded scroll viewport on its first native fragment.
   * The complete native fallback remains available for Word export.
   */
  readonly reserveHeight?: boolean;
  /**
   * Padding for a container's native content, in CSS pixels. Its body paragraphs and tables
   * are laid out narrower by this room on each side, without a document edit, so a frame the
   * view paints around them can hold labels and borders. Nested containers add their insets.
   */
  readonly inset?: ReactBlockInset;
}
export type ReactBlockRegistry = ReadonlyMap<string, ReactBlockView | ReactBlockRegistration>;
export interface ReactBlockHost {
  /** Register a renderer without changing document content. Duplicate types refuse. */
  register(
    type: string,
    view: ReactBlockView,
    options?: { nativeContent?: boolean; reserveHeight?: boolean; inset?: ReactBlockInset }
  ): () => void;
  /** Update explicit provider values. Stable provider component types retain widget identity. */
  setProviders(providers: ReactBlockProviders | undefined): void;
  /** Re-render existing materialized fragments after host-only state changes. */
  refresh(): void;
  /** Dispose before destroying the editor. React cleanup completes in the next microtask. */
  dispose(): void;
}
/** Keep renderer registrations across document loads; each document receives fresh widget state. */
export function mountReactBlockHost(
  editor: DocxEditorInstance,
  initial: ReactBlockRegistry = new Map(),
  options: ReactBlockHostOptions = {}
): ReactBlockHost {
  if (!editor.surface) throw new Error('Attach the editor before mounting React blocks.');
  const registry = new Map(initial);
  let providers = options.providers;
  let disposed = false;
  const lifetime = followEditorSurface(editor, () =>
    mountSurfaceBlockHost(editor, registry, { ...options, providers })
  );
  const assertLive = () => {
    if (disposed) throw new Error('The React block host is disposed.');
  };
  return {
    register(type, view, registration = {}) {
      assertLive();
      validateBlockViewType(type);
      if (registry.has(type)) throw new Error(`React block type is already registered: ${type}`);
      lifetime.current()?.register(type, view, registration);
      const entry = { view, ...registration };
      registry.set(type, entry);
      return () => {
        if (disposed || registry.get(type) !== entry) return;
        registry.delete(type);
        lifetime.current()?.unregister(type);
      };
    },
    setProviders(next) {
      assertLive();
      providers = next;
      lifetime.current()?.setProviders(next);
    },
    refresh() {
      if (!disposed) lifetime.current()?.refresh();
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      lifetime.dispose();
      registry.clear();
      providers = undefined;
    },
  };
}
