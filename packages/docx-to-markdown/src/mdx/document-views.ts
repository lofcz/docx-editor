import type { DocxEditorInstance, SurfacePaintFrame } from '@docx-editor.dev/core/editor';
import {
  findNode,
  type OoxmlContentControlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { describeComponent, indexedComponents, type MdxDocumentComponent } from './document';
import { COMPONENT_TAG, type WordComponent } from './document-components';
import { paintedComponentFragments, paintedPageContents } from './document-painted-pages';

/** The screen-only host of one component fragment on a visible Word page. @public */
export interface MdxComponentView {
  readonly element: HTMLElement;
  readonly component: MdxDocumentComponent;
  readonly editor: DocxEditorInstance;
  readonly pageIndex: number;
  /** Zero-based fragment of this control in document order. */
  readonly fragmentIndex: number;
  /**
   * Read this Word control immediately and after a paint that changes its subtree.
   * The node includes the editable children and nested controls. Return value stops observation.
   * Removing the view also stops observation.
   */
  readonly observeContent: (listener: (control: OoxmlContentControlNode) => void) => () => void;
}

/** Mount a trusted project view. Return a cleanup function for listeners or framework roots. @public */
export type MdxComponentRenderer = (view: MdxComponentView) => void | (() => void);

/**
 * Attach trusted project views to the native geometry of registered component names.
 * The host can add buttons, menus, charts, or animations. Native paragraphs stay editable
 * underneath. Enable pointer events on interactive children. The saved DOCX uses its
 * readable Word fallback. Offscreen views release their resources.
 * @public
 */
export function mountMdxComponents(
  editor: DocxEditorInstance,
  renderers: ReadonlyMap<string, MdxComponentRenderer>
): () => void {
  const surface = editor.surface;
  if (!surface) throw new Error('Attach the document editor before mounting components.');
  type Listener = (control: OoxmlContentControlNode) => void;
  type Mount = {
    element: HTMLElement;
    signature: string;
    node: OoxmlContentControlNode;
    listeners: Set<Listener>;
    stop?: () => void;
  };
  const mounts = new Map<string, Mount>();
  const descriptions = new WeakMap<
    OoxmlContentControlNode,
    { record: WordComponent; component: MdxDocumentComponent; signature: string }
  >();
  let previous: OoxmlPackage | undefined;
  let index: ReadonlyMap<string, WordComponent> = new Map();
  let disposed = false;
  function remove(key: string, mount: Mount) {
    mounts.delete(key);
    mount.listeners.clear();
    try {
      mount.stop?.();
    } finally {
      mount.element.remove();
    }
  }
  function paint(frame: SurfacePaintFrame) {
    const pkg = surface!.session.currentPackage();
    if (disposed || pkg.parts.get(frame.part.name) !== frame.part) return;
    if (pkg !== previous) {
      index = indexedComponents(pkg);
      previous = pkg;
    }
    if (!index.size || !renderers.size) {
      for (const [key, mount] of mounts) remove(key, mount);
      return;
    }
    const pages = paintedPageContents(frame);
    const visible = new Set<string>();
    for (const { control, fragmentIndex, fragment, page } of paintedComponentFragments(
      frame.layout,
      pages
    )) {
      const { pageIndex, box } = fragment;
      if (!control.tag?.startsWith(COMPONENT_TAG)) continue;
      const record = index.get(control.tag.slice(COMPONENT_TAG.length));
      if (!record) throw new Error('Missing MDX component metadata for ' + control.tag + '.');
      const render = renderers.get(record.node.name ?? record.node.type);
      if (!render) continue;
      const node = findNode(frame.part, control.id);
      if (node?.kind !== 'contentControl') continue;
      let entry = descriptions.get(node);
      if (!entry || entry.record !== record) {
        const component = describeComponent(frame.part, node, record, index);
        entry = { record, component, signature: JSON.stringify(component) };
        descriptions.set(node, entry);
      }
      const key = control.id + '/' + pageIndex + '/' + fragmentIndex;
      visible.add(key);
      let mount = mounts.get(key);
      if (mount && mount.signature !== entry.signature) {
        remove(key, mount);
        mount = undefined;
      }
      if (!mount) {
        const element = page.ownerDocument.createElement('div');
        element.className = 'docx-mdx-component-view';
        element.contentEditable = 'false';
        element.setAttribute('data-docx-host-ui', '');
        element.dataset.mdxComponent = entry.component.name;
        mount = { element, signature: entry.signature, node, listeners: new Set() };
        mounts.set(key, mount);
      }
      const { element } = mount;
      const first = !element.parentElement;
      const dimensions = [box.x, box.y, box.width, box.height].map(
        (value) => value * frame.scale + 'px'
      );
      for (const [i, property] of (['left', 'top', 'width', 'height'] as const).entries())
        if (element.style[property] !== dimensions[i]) element.style[property] = dimensions[i]!;
      if (element.parentElement !== page) page.append(element);
      if (first && !mount.stop) {
        try {
          mount.stop =
            render({
              element,
              component: entry.component,
              editor,
              pageIndex,
              fragmentIndex,
              observeContent(listener) {
                if (mounts.get(key) !== mount) return () => {};
                listener(mount.node);
                mount.listeners.add(listener);
                return () => {
                  mount.listeners.delete(listener);
                };
              },
            }) ?? (() => {});
        } catch (error) {
          remove(key, mount);
          throw error;
        }
      } else if (mount.node !== node) {
        mount.node = node;
        for (const listener of mount.listeners) listener(node);
      }
    }
    for (const [key, mount] of mounts) if (!visible.has(key)) remove(key, mount);
  }
  function clear() {
    disposed = true;
    for (const [key, mount] of mounts) remove(key, mount);
  }
  const stop = surface.observePaint({ afterPaint: paint, destroyed: clear });
  return () => {
    stop();
    clear();
  };
}
