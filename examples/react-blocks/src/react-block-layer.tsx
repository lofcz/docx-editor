import { createElement, useLayoutEffect, useRef, type ComponentType, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type { DocxEditorInstance, SurfacePaintFrame } from '@docx-editor.dev/core/editor';
import type { OoxmlNode } from '@docx-editor.dev/core/store';
import { readBlocks, type DocumentBlock, type BlockData } from './document';
import type { BlockBridge } from './bridge';

export interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}
export interface BlockViewProps {
  block: DocumentBlock;
  box: Box;
  header?: Box;
  children: ReactNode;
  update(data: BlockData): void;
  resize(height: number): void;
}
export type BlockRegistry = Readonly<Record<string, ComponentType<BlockViewProps>>>;
interface Borrowed {
  element: HTMLElement;
  parent: HTMLElement;
}
interface Fragment {
  key: string;
  block: DocumentBlock;
  box: Box;
  header?: Box;
  native: HTMLElement[];
  children: Fragment[];
  order: number;
  first: HTMLElement;
}
interface BlockMount {
  layer: HTMLElement;
  root: Root;
  slot?: HTMLElement;
  borrowed: Borrowed[];
}

function union(boxes: Box[]): Box {
  const x = Math.min(...boxes.map((b) => b.x)),
    y = Math.min(...boxes.map((b) => b.y));
  return {
    x,
    y,
    width: Math.max(...boxes.map((b) => b.x + b.width)) - x,
    height: Math.max(...boxes.map((b) => b.y + b.height)) - y,
  };
}
function boxOf(element: HTMLElement, origin: DOMRect): Box {
  const box = element.getBoundingClientRect();
  return {
    x: box.left - origin.left,
    y: box.top - origin.top,
    width: box.width,
    height: box.height,
  };
}
function attach(slot: HTMLElement, elements: HTMLElement[]) {
  for (const old of [...slot.children]) if (!elements.includes(old as HTMLElement)) old.remove();
  for (const element of elements) if (element.parentElement !== slot) slot.append(element);
}
function NativeContent({
  mount,
  elements,
  tag,
}: {
  mount: BlockMount;
  elements: HTMLElement[];
  tag: string;
}) {
  const ref = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    mount.slot = ref.current!;
    attach(mount.slot, elements);
    return () => {
      mount.slot = undefined;
    };
  }, [mount, elements]);
  return <div ref={ref} data-native-slot={tag} style={{ display: 'contents' }} />;
}

// React's delegated listeners run first. Stop UI events at this root before
// the enclosing editable surface can interpret them as document edits.
function stopControlEvent(event: Event) {
  const target = event.target as Node | null;
  const element = target?.nodeType === 1 ? (target as Element) : target?.parentElement;
  if (element?.closest('[data-react-header]')) event.stopPropagation();
}
const UI_EVENTS = [
  'beforeinput',
  'input',
  'keydown',
  'keyup',
  'keypress',
  'compositionstart',
  'compositionupdate',
  'compositionend',
  'paste',
  'copy',
  'cut',
];

/** React roots follow canonical block identity and fragment ordinal across page moves. */
export function mountReactBlocks(
  editor: DocxEditorInstance,
  registry: BlockRegistry,
  bridge: BlockBridge
) {
  const surface = editor.surface!;
  const mounts = new Map<string, BlockMount>();
  let focused: HTMLElement | null = null;
  function restore() {
    for (const mount of mounts.values())
      for (const { element, parent } of mount.borrowed) {
        delete element.dataset.reactFallback;
        if (parent.isConnected && element.parentElement !== parent) parent.append(element);
      }
    // Park custom hosts while the engine reconciles its own page children.
    for (const mount of mounts.values()) mount.layer.remove();
  }
  function dispose(key: string, mount: BlockMount) {
    mounts.delete(key);
    mount.layer.remove();
    // A document switch can dispose this adapter inside another React root's cleanup.
    queueMicrotask(() => mount.root.unmount());
  }
  function paint(frame: SurfacePaintFrame) {
    // Unregistered document tags remain ordinary native Word content.
    const registered = (items: DocumentBlock[]): DocumentBlock[] =>
      items.flatMap((block) =>
        Object.hasOwn(registry, block.type)
          ? [{ ...block, children: registered(block.children) }]
          : registered(block.children)
      );
    const blocks = registered(readBlocks(frame.part));
    const owners = new Map<string, DocumentBlock>();
    const defined = new Set<string>();
    const index = (block: DocumentBlock) => {
      defined.add(block.id);
      for (const id of block.paragraphs) owners.set(id, block);
      block.children.forEach(index);
    };
    blocks.forEach(index);
    const boundaries = new Map((frame.layout.contentControls ?? []).map((c) => [c.id, c]));
    const valid = new Set<string>();
    for (const [id, boundary] of boundaries)
      if (defined.has(id)) boundary.fragments.forEach((_, i) => valid.add(`${id}:${i}`));
    const order = new Map<string, number>();
    const visit = (node: OoxmlNode) => {
      if (node.kind === 'textValue') return;
      if (node.kind === 'paragraph') order.set(node.id, order.size);
      for (const child of node.children) visit(child);
    };
    visit(frame.part.root);
    const elementOrder = (element: HTMLElement) =>
      order.get(
        element.dataset.paragraphId ??
          element.querySelector<HTMLElement>('[data-paragraph-id]')?.dataset.paragraphId ??
          ''
      ) ?? Number.MAX_SAFE_INTEGER;
    for (const page of frame.pagesLayer.querySelectorAll<HTMLElement>(
      '.docx-page[data-materialized="true"]'
    )) {
      const content = page.querySelector<HTMLElement>(':scope > .docx-page-content');
      if (!content) continue;
      const native = new Map<string, HTMLElement[]>();
      for (const element of [...content.children] as HTMLElement[]) {
        const id =
          element.dataset.paragraphId ??
          element.querySelector<HTMLElement>('[data-paragraph-id]')?.dataset.paragraphId;
        const owner = id && owners.get(id);
        if (!owner) continue;
        const list = native.get(owner.id) ?? [];
        list.push(element);
        native.set(owner.id, list);
      }
      if (!native.size) continue;
      const origin = content.getBoundingClientRect();
      const fragments = (block: DocumentBlock): Fragment | null => {
        const elements = native.get(block.id) ?? [];
        const children = block.children.map(fragments).filter((f): f is Fragment => f !== null);
        const boxes = [...elements.map((e) => boxOf(e, origin)), ...children.map((c) => c.box)];
        if (!boxes.length) return null;
        const first = [...elements, ...children.map((c) => c.first)].sort(
          (a, b) => elementOrder(a) - elementOrder(b)
        )[0]!;
        const header = elements.find((e) => e.dataset.paragraphId === block.headerId);
        if (header) header.dataset.reactFallback = '';
        const ordinal =
          boundaries
            .get(block.id)
            ?.fragments.findIndex((f) => f.pageIndex === Number(page.dataset.pageIndex)) ?? 0;
        const key = `${block.id}:${Math.max(0, ordinal)}`;
        valid.add(key);
        return {
          key,
          block,
          box: union(boxes),
          header: header ? boxOf(header, origin) : undefined,
          native: elements,
          children,
          order: elementOrder(first),
          first,
        };
      };
      const roots = blocks.map(fragments).filter((f): f is Fragment => f !== null);
      function mountFor(fragment: Fragment): BlockMount {
        let mount = mounts.get(fragment.key);
        if (!mount) {
          const layer = content!.ownerDocument.createElement('div');
          layer.dataset.reactBlockHost = fragment.key;
          layer.style.display = 'contents';
          const root = createRoot(layer);
          UI_EVENTS.forEach((type) => layer.addEventListener(type, stopControlEvent));
          mount = { layer, root, borrowed: [] };
          mounts.set(fragment.key, mount);
        }
        return mount;
      }
      // Insert the outer hosts in document order before borrowing their text nodes.
      for (const fragment of roots) content.insertBefore(mountFor(fragment).layer, fragment.first);
      function render(fragment: Fragment) {
        const mount = mountFor(fragment);
        const childNodes = fragment.children.map((child) => ({
          element: mountFor(child).layer,
          order: child.order,
        }));
        const elements = [
          ...fragment.native.map((element) => ({ element, order: elementOrder(element) })),
          ...childNodes,
        ]
          .sort((a, b) => a.order - b.order)
          .map((x) => x.element);
        mount.borrowed = fragment.native.map((element) => ({ element, parent: content! }));
        if (mount.slot) attach(mount.slot, elements);
        const View = registry[fragment.block.type];
        if (View)
          mount.root.render(
            createElement(View, {
              block: fragment.block,
              box: fragment.box,
              header: fragment.header,
              update: (data: BlockData) => bridge.update(fragment.block, data),
              resize: (height: number) => bridge.resize(fragment.block, height, frame.scale),
              children: (
                <NativeContent mount={mount} elements={elements} tag={fragment.block.tag} />
              ),
            })
          );
        fragment.children.forEach(render);
      }
      roots.forEach(render);
    }
    // Keep roots parked for virtual pages, so scrolling retains local React state.
    for (const [key, mount] of mounts) if (!valid.has(key)) dispose(key, mount);
    if (focused?.isConnected && focused.ownerDocument.activeElement !== focused)
      focused.focus({ preventScroll: true });
  }
  const stop = surface.observePaint({
    beforePaint() {
      focused = surface.hostElement().ownerDocument.activeElement as HTMLElement | null;
      restore();
    },
    afterPaint: paint,
  });
  return () => {
    stop();
    restore();
    for (const [key, mount] of mounts) dispose(key, mount);
  };
}
