import type { SurfacePaintFrame } from '@docx-editor.dev/core/editor';

/** Opaque widgets stay connected while the native painter replaces page descendants. */
export class OpaqueBlockLayer {
  private element: HTMLDivElement | null = null;

  contains(element: HTMLElement): boolean {
    return element.parentElement === this.element;
  }

  place(
    frame: SurfacePaintFrame,
    content: HTMLElement,
    root: HTMLElement,
    box: { x: number; y: number; width: number; height: number }
  ): void {
    const parent = frame.pagesLayer.parentElement;
    const page = content.parentElement;
    if (!parent || !page) return;
    if (!this.element) {
      this.element = parent.ownerDocument.createElement('div');
      this.element.dataset.sciobotBlockLayer = '';
      this.element.dataset.docxHostUi = '';
      this.element.contentEditable = 'false';
      Object.assign(this.element.style, { position: 'absolute', pointerEvents: 'none' });
    }
    if (this.element.parentElement !== parent) parent.append(this.element);
    Object.assign(this.element.style, {
      left: `${frame.pagesLayer.offsetLeft}px`,
      top: `${frame.pagesLayer.offsetTop}px`,
      width: frame.pagesLayer.style.width,
      height: frame.pagesLayer.style.height,
    });
    // The painter supplies these pixel offsets, including facing-page placement.
    const pageX = Number.parseFloat(page.style.left) || 0;
    const pageY = Number.parseFloat(page.style.top) || 0;
    const x = (Number.parseFloat(content.style.left) || 0) + box.x * frame.scale;
    const y = (Number.parseFloat(content.style.top) || 0) + box.y * frame.scale;
    const width = box.width * frame.scale;
    const height = box.height * frame.scale;
    const pageWidth = Number.parseFloat(page.style.width) || 0;
    const pageHeight = Number.parseFloat(page.style.height) || 0;
    Object.assign(root.style, {
      left: `${pageX + x}px`,
      top: `${pageY + y}px`,
      // Include the page margins for handles, but never paint across another sheet.
      clipPath: `inset(${-y}px ${x + width - pageWidth}px ${y + height - pageHeight}px ${-x}px)`,
    });
    if (root.parentElement !== this.element) this.element.append(root);
  }

  dispose(): void {
    this.element?.remove();
    this.element = null;
  }
}
