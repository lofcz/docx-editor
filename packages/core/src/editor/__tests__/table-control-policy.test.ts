import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { buildTableCommandPlannerInput } from '../docx-editor-derive.ts';
import { planTableCommand } from '../table-command-plan.ts';
import { moveByTabInTable } from '../surface-table-tab.ts';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { docx, paragraph } from './paginated-surface-fixtures.ts';

const table = (left: string, right: string) =>
  '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr><w:tc>${left}</w:tc><w:tc>${right}</w:tc></w:tr></w:tbl>`;

test('host table ownership gates native commands and furniture while preserving content and nested tables', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const owned = new Set<string>();
  const reason = 'Use the application column controls.';
  const editor = createDocxEditor({
    container,
    document: docx(
      table(
        paragraph('Left') + table(paragraph('Inner A'), paragraph('Inner B')),
        paragraph('Right')
      ) + paragraph('After')
    ),
    tableControlRefusal: (id, pkg) => {
      expect(pkg.parts.has(pkg.mainDocumentPart)).toBe(true);
      return owned.has(id) ? reason : null;
    },
  });
  try {
    const surface = editor.surface!;
    const place = (text: string) => {
      const paragraphId = surface.session
        .paragraphIds()
        .find((id) => paragraphTextOf(surface.session.part(), id) === text)!;
      surface.setSelection({
        anchor: { paragraphId, offset: 0 },
        head: { paragraphId, offset: 0 },
      });
      return paragraphId;
    };
    place('Right');
    const raw = planTableCommand(
      buildTableCommandPlannerInput({ type: 'insertRow', where: 'below' }, surface)
    );
    expect(raw.ok).toBe(true);
    const outer = surface.layout().pages[0]!.fragments.find((block) => block.kind === 'table')!;
    owned.add(outer.tableId);
    place('Right');
    expect(editor.can({ type: 'insertRow', where: 'below' })).toMatchObject({ ok: false, reason });
    expect(editor.exec({ type: 'deleteTable' })).toMatchObject({ ok: false, reason });
    expect(surface.applyTableCommandPlan(raw)).toMatchObject({ ok: false, reason });
    const pkg = surface.session.currentPackage();
    expect(moveByTabInTable(surface, 'next')).toBe(true);
    expect(surface.session.currentPackage()).toBe(pkg);
    expect(container.querySelector('.docx-table-select-handle')).toBeNull();
    const pages = container.querySelector<HTMLElement>('.docx-pages')!;
    const page = surface.layout().pages[0]!;
    pages.dispatchEvent(
      new PointerEvent('pointermove', {
        bubbles: true,
        clientX: ((page.contentBox.x + outer.box.x + outer.columnEdges[1]!) * 4) / 3,
        clientY: ((page.contentBox.y + outer.rows[0]!.box.y + 5) * 4) / 3,
      })
    );
    expect(container.querySelector('.docx-table-divider-handle')).toBeNull();
    const right = place('Right');
    surface.type('Typed ');
    expect(paragraphTextOf(surface.session.part(), right)).toBe('Typed Right');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(paragraphTextOf(surface.session.part(), right)).toBe('Right');
    place('Inner B');
    expect(editor.can({ type: 'insertRow', where: 'below' }).ok).toBe(true);
    expect(editor.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
  } finally {
    editor.destroy();
    container.remove();
  }
});
