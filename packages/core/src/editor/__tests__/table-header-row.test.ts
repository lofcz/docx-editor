import { afterEach, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor } from '../docx-editor.ts';
import { docx, paragraph as p } from './paginated-surface-fixtures.ts';
import { paragraphTextFromLayout } from '../../layout/semantic-interaction.ts';

const cleanups: (() => void)[] = [];
afterEach(() => cleanups.splice(0).forEach((cleanup) => cleanup()));
const table = (headers = 0) =>
  '<w:tbl><w:tblPr><w:tblStyle w:val="Imported"/><w:tblLook w:val="04A0"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  Array.from(
    { length: 16 },
    (_, i) =>
      `<w:tr><w:trPr><w:trHeight w:val="1400" w:hRule="atLeast"/>${i < headers ? '<w:tblHeader/>' : ''}</w:trPr><w:tc><w:tcPr><w:shd w:fill="FFEEDD"/></w:tcPr>${p(`Row ${i}`)}</w:tc></w:tr>`
  ).join('') +
  '</w:tbl>';
function mount(headers = 0, owned = false) {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(p('Before') + table(headers) + p('After')),
    tableControlRefusal: owned ? () => 'Owned table' : undefined,
  });
  cleanups.push(() => {
    editor.destroy();
    container.remove();
  });
  const surface = editor.surface!;
  const id = surface.session
    .paragraphIds()
    .find((id) => paragraphTextFromLayout(surface.layout(), id) === 'Row 2')!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 0 },
  });
  const rows = () =>
    surface
      .layout()
      .pages.flatMap((page) =>
        page.fragments.flatMap((block) => (block.kind === 'table' ? block.rows : []))
      );
  return { editor, surface, rows };
}
const toggle = { type: 'toggleHeaderRow' } as const;

test('header toggle repeats the first row across pages and preserves imported styling', async () => {
  const { editor, surface, rows } = mount();
  const selection = surface.state().selection;
  expect(editor.can(toggle).ok).toBe(true);
  expect(editor.exec(toggle).ok).toBe(true);
  expect(surface.state().selection).toEqual(selection);
  expect(rows().some((row) => row.isHeaderRepeat)).toBe(true);
  expect(
    new Set(
      rows()
        .filter((row) => row.isHeaderRow)
        .map((row) => row.id)
    ).size
  ).toBe(1);
  const xml = strFromU8(unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!);
  expect(xml).toContain('w:val="Imported"');
  expect(xml).toContain('w:val="04A0"');
  expect(xml).toContain('w:fill="FFEEDD"');
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(rows().some((row) => row.isHeaderRow)).toBe(false);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  expect(rows().some((row) => row.isHeaderRepeat)).toBe(true);
});

test('stop repeating clears imported multiple-row headers and undo restores all of them', () => {
  const { editor, rows } = mount(2);
  const count = () =>
    new Set(
      rows()
        .filter((row) => row.isHeaderRow)
        .map((row) => row.id)
    ).size;
  expect(count()).toBe(2);
  expect(editor.exec(toggle).ok).toBe(true);
  expect(count()).toBe(0);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(count()).toBe(2);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  expect(count()).toBe(0);
});

test('header toggle respects viewing mode and application-owned tables', () => {
  expect(mount(0, true).editor.can(toggle).ok).toBe(false);
  const { editor } = mount();
  editor.setEditingMode('viewing');
  expect(editor.can(toggle).ok).toBe(false);
});
