import { createElement, useLayoutEffect, useRef, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type { DocxEditorInstance, SurfacePaintFrame } from '@docx-editor.dev/core/editor';
import {
  findNode,
  contentControlContentNodeOf,
  paragraphTextOf,
  storyParagraphs,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildBlockIndex, type BlockIndex } from '../blocks/block-index';
import { isolateWidgetInput } from './input-boundary';
import { inlineLayoutBoxes } from './inline-layout-boxes';
import { BlockMeasurement, lastBodyParagraphIn, type MeasuredFragment } from './measurement';
import { blockViewportHeight } from './overflow';
import { containerInsets } from './block-insets';
import { NativeSlotLease, preserveNativeSelection } from './native-slots';
import { attachNativeSlots, type NativeSlotMount } from './native-slot-plan';
import { BlockRootProviders, type ReactBlockHostOptions } from './providers';

import type { ReactBlockHost, ReactBlockRegistration, ReactBlockRegistry } from './host';

export interface SurfaceBlockHost extends ReactBlockHost {
  unregister(type: string): void;
}

export function validateBlockViewType(type: string): void {
  if (!/^[A-Za-z][A-Za-z0-9_.:-]{0,127}$/.test(type))
    throw new TypeError('Invalid React block type.');
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
export function mountSurfaceBlockHost(
  editor: DocxEditorInstance,
  initial: ReactBlockRegistry = new Map(),
  options: ReactBlockHostOptions = {}
): SurfaceBlockHost {
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
  // Insets are computed again when the package or the registry changes.
  let insetsFor: OoxmlPackage | null = null;
  const measurement = options.inlineOnly ? null : new BlockMeasurement(surface, options.measure);
  for (const [type, view] of initial) {
    validateBlockViewType(type);
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
    if (!options.inlineOnly && insetsFor !== pkg) {
      insetsFor = pkg;
      const insets = containerInsets(next.part, index!, (type) => registry.get(type)?.inset);
      // Outside the paint: a changed inset lays the document out again.
      queueMicrotask(() => {
        if (!disposed) surface!.setBlockInsets(insets);
      });
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
    const rooms = measurement ? widgetRooms(index!, next, registry, measurement) : null;
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
      const measured =
        measurement !== null &&
        registration.reserveHeight === true &&
        !nativeContent &&
        boundary.level !== 'inline';
      const fragments =
        options.inlineOnly || measured ? boundary.fragments.slice(0, 1) : boundary.fragments;
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
        const nativeBox =
          inlineBoxes?.get(`${fragment.pageIndex}:${block.controlId}`) ?? fragment.box;
        // A container's frame covers the room its measured widgets reserve below their text.
        const roomBottom = rooms?.get(`${block.id}:${fragment.pageIndex}`);
        const box =
          roomBottom !== undefined && roomBottom > nativeBox.y + nativeBox.height
            ? { ...nativeBox, height: roomBottom - nativeBox.y }
            : nativeBox;
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
        // One opaque widget keeps its identity when the Word fallback splits across pages.
        // Reserve only excess viewport height after the complete fallback's last paragraph.
        // Bounded by a whole page, not the rest of this one: a widget that fits an empty page
        // reserves its full height, and the layout moves its paragraph to the next page.
        const maxHeight = blockViewportHeight(next, fragment.pageIndex, 0);
        const target = measured
          ? measuredFragment(
              block.controlId,
              boundary.fragments.reduce((sum, item) => sum + item.box.height, 0),
              next,
              maxHeight
            )
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
                    maxHeight,
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
    destroyed: dispose,
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
  function dispose(): void {
    if (disposed) return;
    disposed = true;
    if (frame) preserveNativeSelection(frame.pagesLayer, stop);
    else stop();
    for (const [key, mount] of mounts) discard(key, mount);
    measurement?.dispose();
    if (!options.inlineOnly) surface!.setBlockInsets(new Map());
    registry.clear();
    frame = null;
    index = null;
    indexedPackage = null;
    focused = null;
    providers = undefined;
  }
  return {
    setProviders(next) {
      if (disposed) throw new Error('The React block host is disposed.');
      providers = next;
      refresh();
    },
    register(type, view, options = {}) {
      if (disposed) throw new Error('The React block host is disposed.');
      validateBlockViewType(type);
      if (registry.has(type)) throw new Error(`React block type is already registered: ${type}`);
      registry.set(type, { view, ...options });
      insetsFor = null;
      refresh();
      let registered = true;
      return () => {
        if (!registered || disposed) return;
        registered = false;
        registry.delete(type);
        insetsFor = null;
        refresh();
      };
    },
    unregister(type) {
      if (disposed || !registry.delete(type)) return;
      insetsFor = null;
      refresh();
    },
    refresh,
    dispose,
  };
}

function measuredFragment(
  controlId: string,
  nativeHeight: number,
  frame: SurfacePaintFrame,
  maxHeight: number
): MeasuredFragment | null {
  const control = findNode(frame.part, controlId);
  const paragraphId = control ? lastBodyParagraphIn(control) : null;
  return paragraphId ? { paragraphId, nativeHeight, scale: frame.scale, maxHeight } : null;
}

/**
 * The widget's bounded box. Its height follows its content up to the page limit, never
 * the native reservation, so a reservation cannot change the next measurement.
 */
function MeasuredRoot({
  measurement,
  mountKey,
  target,
  maxHeight,
  children,
}: {
  readonly measurement: BlockMeasurement;
  readonly mountKey: string;
  readonly target: MeasuredFragment | null;
  readonly maxHeight: number;
  readonly children: ReactNode;
}) {
  const ref = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    if (target && ref.current) measurement.track(mountKey, ref.current, target);
    else measurement.release(mountKey);
  });
  useLayoutEffect(() => () => measurement.release(mountKey), [measurement, mountKey]);
  return createElement(
    'div',
    {
      ref,
      'data-sciobot-block-measure': '',
      'data-sciobot-block-ui': '',
      'data-docx-host-ui': '',
      contentEditable: false,
      style: {
        maxHeight,
        // The widget covers its Word fallback text, so the fallback never shows through it.
        background: 'var(--doc-page-bg-rendered, #fff)',
        overflowY: 'auto',
        overscrollBehaviorY: 'contain',
        pointerEvents: 'auto',
      },
    },
    children
  );
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

/**
 * Where the reserved room of each measured widget ends, per enclosing block and page: the
 * bottom of the widget's last native fragment plus its reservation, clamped to the page.
 * Keys are `blockId:pageIndex`; the widget's own block is not a key for its room.
 */
function widgetRooms(
  index: BlockIndex,
  frame: SurfacePaintFrame,
  registry: ReadonlyMap<string, ReactBlockRegistration>,
  measurement: BlockMeasurement
): Map<string, number> {
  const rooms = new Map<string, number>();
  const boundaries = new Map(
    (frame.layout.contentControls ?? []).map((boundary) => [boundary.id, boundary])
  );
  for (const block of index.byId.values()) {
    if (block.partName !== frame.part.name || !registry.get(block.payload.type)?.reserveHeight)
      continue;
    const control = findNode(frame.part, block.controlId);
    const paragraph = control ? lastBodyParagraphIn(control) : null;
    const room = paragraph ? measurement.reservedAfter(paragraph) : 0;
    const last = boundaries.get(block.controlId)?.fragments.at(-1);
    if (!(room > 0) || !last) continue;
    const pageBottom = frame.layout.pages[last.pageIndex]?.contentBox.height ?? Infinity;
    const bottom = Math.min(last.box.y + last.box.height + room, pageBottom);
    for (
      let parent: string | null | undefined = block.parentId;
      parent;
      parent = index.byId.get(parent)?.parentId
    ) {
      const key = `${parent}:${last.pageIndex}`;
      rooms.set(key, Math.max(rooms.get(key) ?? 0, bottom));
    }
  }
  return rooms;
}
