import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { expect, test } from 'bun:test';
import { mountPaginatedSurface } from '../paginated-surface.ts';
import { docx, paragraph } from './paginated-surface-fixtures.ts';

test('application controls clear table hover furniture and keep their own pointer gestures', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const mounted = mountPaginatedSurface(
    container,
    docx(
      '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
        `<w:tr><w:tc>${paragraph('Left')}</w:tc><w:tc>${paragraph('Right')}</w:tc></w:tr></w:tbl>`
    ),
    { scale: 1 }
  );
  if (!mounted.ok) throw new Error(mounted.reason);
  const surface = mounted.surface;
  try {
    const pages = container.querySelector<HTMLElement>('.docx-pages')!;
    const furniture = container.querySelector<HTMLElement>('.docx-table-furniture')!;
    const move = () => {
      const page = surface.layout().pages[0]!;
      const table = page.fragments.find((block) => block.kind === 'table')!;
      return new PointerEvent('pointermove', {
        bubbles: true,
        clientX: page.contentBox.x + table.box.x + table.columnEdges[1]!,
        clientY: page.contentBox.y + table.rows[0]!.box.y + table.rows[0]!.box.height / 2,
      });
    };
    const divider = () => furniture.querySelector('.docx-table-divider-handle');
    pages.dispatchEvent(move());
    expect(divider()).not.toBeNull();
    const ui = document.createElement('div');
    ui.dataset.docxHostUi = '';
    const control = document.createElement('button');
    ui.append(control);
    pages.append(ui);
    control.dispatchEvent(move());
    expect(divider()).toBeNull();
    surface.type('text');
    expect(divider()).toBeNull();
    control.dispatchEvent(move());
    expect(divider()).toBeNull();
    pages.dispatchEvent(move());
    expect(divider()).not.toBeNull();
  } finally {
    surface.destroy();
    container.remove();
  }
});
