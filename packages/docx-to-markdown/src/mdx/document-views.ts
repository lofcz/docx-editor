import type { DocxEditorInstance, SurfacePaintFrame } from '@docx-editor.dev/core/editor';
import type { OoxmlPackage } from '@docx-editor.dev/core/store';
import { readMdxComponents, type MdxDocumentComponent } from './document';

/** The screen-only host of one component fragment on a visible Word page. @public */
export interface MdxComponentView {
  readonly element: HTMLElement;
  readonly component: MdxDocumentComponent;
  readonly editor: DocxEditorInstance;
  readonly pageIndex: number;
  /** Zero-based fragment of this control in document order. */
  readonly fragmentIndex: number;
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
  type Mount = { element: HTMLElement; signature: string; stop?: () => void };
  const mounts = new Map<string, Mount>();
  let previous: OoxmlPackage | undefined;
  let byControl = new Map<
    string,
    { component: MdxDocumentComponent; signature: string; render: MdxComponentRenderer }
  >();
  let disposed = false;
  function remove(key: string, mount: Mount) {
    mounts.delete(key);
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
      previous = pkg;
      byControl = new Map();
      for (const component of readMdxComponents(pkg)) {
        const render = renderers.get(component.name);
        if (render)
          byControl.set(component.controlId, {
            component,
            render,
            signature: JSON.stringify(component),
          });
      }
    }
    const pages = new Map<number, HTMLElement>();
    for (const content of frame.pagesLayer.querySelectorAll<HTMLElement>('.docx-page-content')) {
      const page = content.closest<HTMLElement>('.docx-page');
      if (page) pages.set(Number(page.dataset.pageIndex), content);
    }
    const visible = new Set<string>();
    for (const control of frame.layout.contentControls ?? []) {
      const entry = byControl.get(control.id);
      if (!entry) continue;
      for (const [fragmentIndex, fragment] of control.fragments.entries()) {
        const { box, pageIndex } = fragment;
        const page = pages.get(pageIndex);
        if (!page) continue;
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
          mount = { element, signature: entry.signature };
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
              entry.render({
                element,
                component: entry.component,
                editor,
                pageIndex,
                fragmentIndex,
              }) ?? (() => {});
          } catch (error) {
            remove(key, mount);
            throw error;
          }
        }
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
