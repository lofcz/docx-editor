import { createElement, useLayoutEffect, useRef, type ComponentType, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type { DocxEditorInstance, SurfacePaintFrame } from '@docx-editor.dev/core/editor';
import {
  findNode,
  contentControlContentNodeOf,
  paragraphTextOf,
  storyParagraphs,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex, type BlockIndex, type IndexedBlock } from '../blocks/block-index';
import { isolateWidgetInput } from './input-boundary';
import { inlineLayoutBoxes } from './inline-layout-boxes';
import { BlockMeasurement, lastBodyParagraphIn, type MeasuredFragment } from './measurement';
import { NativeSlotLease, preserveNativeSelection } from './native-slots';
import { attachNativeSlots, type NativeSlotMount } from './native-slot-plan';
import {
  BlockRootProviders,
  type ReactBlockHostOptions,
  type ReactBlockProviders,
} from './providers';

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
   * Reserve native flow for a widget taller than its Word fallback. The widget's intrinsic
   * height is measured; following native content moves down without a document edit.
   */
  readonly reserveHeight?: boolean;
}
export type ReactBlockRegistry = ReadonlyMap<string, ReactBlockView | ReactBlockRegistration>;
export interface ReactBlockHost {
  /** Register a renderer without changing document content. Duplicate types refuse. */
  register(
    type: string,
    view: ReactBlockView,
    options?: { nativeContent?: boolean; reserveHeight?: boolean }
  ): () => void;
  /** Update explicit provider values. Stable provider component types retain widget identity. */
  setProviders(providers: ReactBlockProviders | undefined): void;
  /** Re-render existing materialized fragments after host-only state changes. */
  refresh(): void;
  /** Dispose before destroying the editor. React cleanup completes in the next microtask. */
  dispose(): void;
}
interface Mount {
  readonly element: HTMLDivElement;
  readonly root: Root;
  readonly slot: NativeSlotLease;
  readonly release: () => void;
}
/**
 * Independent React roots, keyed by persisted block ID and native fragment ordinal.
 * Page moves and canonical node-ID changes retain each root and its local React state.
 * Native slots borrow the original paint nodes. Container wrappers must use `display: contents`.
 * Unknown, invalid, or unregistered payloads retain their native Word rendering.
 */
export function mountReactBlockHost(
  editor: DocxEditorInstance,
  initial: ReactBlockRegistry = new Map(),
  options: ReactBlockHostOptions = {}
): ReactBlockHost {
  const surface = editor.surface;
  if (!surface) throw new Error('Attach the editor before mounting React blocks.');
  const registry = new Map<string, ReactBlockRegistration>();
  const mounts = new Map<string, Mount>();
  let disposed = false;
  let frame: SurfacePaintFrame | null = null;
  let focused: HTMLElement | null = null;
  let indexedPackage: OoxmlPackage | null = null;
  let index: BlockIndex | null = null;
  let providers = options.providers;
  const measurement = options.inlineOnly ? null : new BlockMeasurement(surface, options.measure);
  const validateType = (type: string) => {
    if (!/^[A-Za-z][A-Za-z0-9_.:-]{0,127}$/.test(type))
      throw new TypeError('Invalid React block type.');
  };
  for (const [type, view] of initial) {
    validateType(type);
    registry.set(
      type,
      typeof view === 'object' && view !== null && 'view' in view ? view : { view }
    );
  }
  function park(): void {
    for (const mount of mounts.values()) mount.slot.reset();
    for (const mount of mounts.values()) mount.element.remove();
  }
  function discard(key: string, mount: Mount): void {
    mounts.delete(key);
    measurement?.release(key);
    mount.slot.reset();
    mount.element.remove();
    mount.release();
    // Hosts can be disposed from an enclosing React effect cleanup.
    queueMicrotask(() => mount.root.unmount());
  }
  function paint(next: SurfacePaintFrame): void {
    if (disposed) return;
    frame = next;
    const pkg = surface!.session.currentPackage();
    // A pending native paint will publish its own matching frame.
    if (pkg.parts.get(next.part.name) !== next.part) return;
    if (pkg !== indexedPackage) {
      indexedPackage = pkg;
      index = buildBlockIndex(pkg);
    }
    const controls = new Map(
      (next.layout.contentControls ?? []).map((boundary) => [boundary.id, boundary])
    );
    const pages = new Map<number, HTMLElement>();
    for (const page of next.pagesLayer.querySelectorAll<HTMLElement>(
      '.docx-page[data-materialized="true"]'
    )) {
      const content = page.querySelector<HTMLElement>(':scope > .docx-page-content');
      if (content) pages.set(Number(page.dataset.pageIndex), content);
    }
    const valid = new Set<string>();
    const inlineBoxes = options.inlineOnly ? inlineLayoutBoxes(next.layout) : null;
    const native: NativeSlotMount[] = [];
    const render: (() => void)[] = [];
    for (const block of index!.byId.values()) {
      if (block.partName !== next.part.name) continue;
      const registration = registry.get(block.payload.type),
        boundary = controls.get(block.controlId);
      if (!registration || !boundary || (options.inlineOnly && boundary.level !== 'inline'))
        continue;
      const View = registration.view;
      const nativeContent = registration.nativeContent === true && boundary.level !== 'inline';
      const paragraphs = Object.freeze(
        boundary.range
          ? [
              {
                id: boundary.range.paragraphId,
                text: paragraphTextOf(next.part, boundary.range.paragraphId) ?? '',
              },
            ]
          : collectParagraphs(block.controlId, next)
      );
      const fragments = options.inlineOnly ? boundary.fragments.slice(0, 1) : boundary.fragments;
      fragments.forEach((fragment, ordinal) => {
        const key = `${block.id}:${ordinal}`;
        valid.add(key);
        const page = pages.get(fragment.pageIndex);
        if (!page) return; // Keep roots parked for virtual pages.
        let mount = mounts.get(key);
        if (!mount) {
          const element = page.ownerDocument.createElement('div');
          element.dataset.sciobotBlockHost = key;
          element.style.position = 'absolute';
          element.style.pointerEvents = 'none';
          if (options.inlineOnly) element.style.visibility = 'hidden';
          const root = createRoot(element);
          const release = isolateWidgetInput(element);
          mount = { element, root, slot: new NativeSlotLease(key, next.pagesLayer), release };
          mounts.set(key, mount);
        }
        const box = inlineBoxes?.get(`${fragment.pageIndex}:${block.controlId}`) ?? fragment.box;
        Object.assign(mount.element.style, {
          display: nativeContent ? 'contents' : 'block',
          position: nativeContent ? 'static' : 'absolute',
          pointerEvents: nativeContent ? 'auto' : 'none',
          left: `${box.x * next.scale}px`,
          top: `${box.y * next.scale}px`,
          width: `${box.width * next.scale}px`,
          height: `${box.height * next.scale}px`,
        });
        if (nativeContent)
          native.push({
            block,
            pageIndex: fragment.pageIndex,
            element: mount.element,
            slot: mount.slot,
          });
        else page.append(mount.element);
        const currentMount = mount;
        const geometry = Object.freeze({
          ordinal,
          pageIndex: fragment.pageIndex,
          box: Object.freeze({ ...box }),
          scale: next.scale,
        });
        const view = createElement(View, {
          block,
          fragment: geometry,
          paragraphs,
          revision: surface!.session.packageRevision(),
          children: nativeContent ? currentMount.slot.content() : null,
        });
        // Every fragment of a measured block keeps the same wrapper type, so a block that
        // splits or rejoins across pages never remounts its widget. Only the last fragment
        // reserves: native flow grows after the block's last line.
        const measured =
          measurement !== null &&
          registration.reserveHeight === true &&
          !nativeContent &&
          boundary.level !== 'inline';
        const target =
          measured && ordinal === boundary.fragments.length - 1
            ? measuredFragment(block.controlId, box.height, next)
            : null;
        render.push(() =>
          currentMount.root.render(
            createElement(BlockRootProviders, {
              render: providers,
              scope: { editor, block, fragment: geometry },
              children: measured
                ? createElement(MeasuredRoot, {
                    measurement: measurement!,
                    mountKey: key,
                    target,
                    children: view,
                  })
                : view,
            })
          )
        );
      });
    }
    for (const [key, mount] of mounts) if (!valid.has(key)) discard(key, mount);
    attachNativeSlots(next, native);
    for (const callback of render) callback();
    measurement?.schedule();
    if (focused?.isConnected && focused.ownerDocument.activeElement !== focused)
      focused.focus({ preventScroll: true });
    focused = null;
  }
  const stop = surface.observePaint({
    beforePaint() {
      const active = surface.hostElement().ownerDocument.activeElement;
      focused = [...mounts.values()].some((mount) => active && mount.element.contains(active))
        ? (active as HTMLElement)
        : null;
      park();
    },
    afterPaint: paint,
  });
  function refresh(): void {
    if (disposed || !frame) return;
    preserveNativeSelection(frame.pagesLayer, () => {
      const active = surface!.hostElement().ownerDocument.activeElement;
      focused = [...mounts.values()].some((mount) => active && mount.element.contains(active))
        ? (active as HTMLElement)
        : null;
      park();
      paint(frame!);
    });
  }
  return {
    setProviders(next) {
      if (disposed) throw new Error('The React block host is disposed.');
      providers = next;
      refresh();
    },
    register(type, view, options = {}) {
      if (disposed) throw new Error('The React block host is disposed.');
      validateType(type);
      if (registry.has(type)) throw new Error(`React block type is already registered: ${type}`);
      registry.set(type, { view, ...options });
      refresh();
      let registered = true;
      return () => {
        if (!registered || disposed) return;
        registered = false;
        registry.delete(type);
        refresh();
      };
    },
    refresh,
    dispose() {
      if (disposed) return;
      disposed = true;
      if (frame) preserveNativeSelection(frame.pagesLayer, stop);
      else stop();
      for (const [key, mount] of mounts) discard(key, mount);
      measurement?.dispose();
      registry.clear();
      frame = null;
      index = null;
      indexedPackage = null;
      focused = null;
      providers = undefined;
    },
  };
}

function measuredFragment(
  controlId: string,
  nativeHeight: number,
  frame: SurfacePaintFrame
): MeasuredFragment | null {
  const control = findNode(frame.part, controlId);
  const paragraphId = control ? lastBodyParagraphIn(control) : null;
  return paragraphId ? { paragraphId, nativeHeight, scale: frame.scale } : null;
}

/**
 * The widget's intrinsic box. Its height follows the widget content, never the native
 * fragment box, so a reservation cannot change the next measurement.
 */
function MeasuredRoot({
  measurement,
  mountKey,
  target,
  children,
}: {
  readonly measurement: BlockMeasurement;
  readonly mountKey: string;
  readonly target: MeasuredFragment | null;
  readonly children: ReactNode;
}) {
  const ref = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    if (target && ref.current) measurement.track(mountKey, ref.current, target);
    else measurement.release(mountKey);
  });
  useLayoutEffect(() => () => measurement.release(mountKey), [measurement, mountKey]);
  return createElement('div', { ref, 'data-sciobot-block-measure': '' }, children);
}

function collectParagraphs(
  controlId: string,
  frame: SurfacePaintFrame
): { readonly id: string; readonly text: string }[] {
  // The canonical index owns traversal. UI roots receive detached text, never editable DOM.
  const out: { readonly id: string; readonly text: string }[] = [];
  const node = findNode(frame.part, controlId);
  if (node)
    for (const paragraph of storyParagraphs(contentControlContentNodeOf(node) ?? node))
      out.push(
        Object.freeze({ id: paragraph.id, text: paragraphTextOf(frame.part, paragraph.id) ?? '' })
      );
  return out;
}
