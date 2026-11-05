import { afterEach, expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { docx, paragraph as p } from './paginated-surface-fixtures.ts';
import { paragraphTextFromLayout } from '../../layout/semantic-interaction.ts';

const grid = '<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>';
const cell = (text: string) => `<w:tc>${p(text)}</w:tc>`;
const table = `<w:tbl>${grid}<w:tr>${cell('A')}${cell('B')}</w:tr><w:tr>${cell('C')}${cell('D')}</w:tr></w:tbl>`;
const cleanups: (() => void)[] = [];
afterEach(() => {
  cleanups.splice(0).forEach((cleanup) => cleanup());
});
function mount(owned = false, markup = table) {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(p('Before') + markup + p('After')),
    tableControlRefusal: owned ? () => 'Owned table' : undefined,
  });
  cleanups.push(() => {
    editor.destroy();
    container.remove();
  });
  const surface = editor.surface!;
  const id = surface.session
    .paragraphIds()
    .find((id) => paragraphTextFromLayout(surface.layout(), id) === 'A')!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 0 },
  });
  const control = (kind: string) =>
    container.querySelector<HTMLButtonElement>(`[data-table-append="${kind}"]`)!;
  const shape = () => {
    const table = surface.layout().pages[0]!.fragments.find((item) => item.kind === 'table');
    if (table?.kind !== 'table') throw Error('Missing table');
    return [table.rows.length, table.rows[0]!.cells.length];
  };
  return { editor, surface, control, shape, container };
}

test('the corner appends a row and column in one undo transaction', () => {
  const { editor, control, shape } = mount();
  control('both').click();
  expect(shape()).toEqual([3, 3]);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(shape()).toEqual([2, 2]);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  expect(shape()).toEqual([3, 3]);
});
test('append buttons retain keyboard focus across repeated operations', () => {
  const { control, shape } = mount();
  const button = control('row');
  button.focus();
  button.click();
  button.click();
  expect(shape()).toEqual([4, 2]);
  expect(document.activeElement).toBe(button);
  control('column').click();
  expect(shape()).toEqual([4, 3]);
});
test('owned tables and viewing mode hide append buttons', () => {
  expect(mount(true).control('row')).toBeNull();
  const { editor, control } = mount();
  editor.setEditingMode('viewing');
  expect(control('row')).toBeNull();
});

test('append controls support authored tables without a column grid', () => {
  const { editor, control, shape } = mount(false, table.replace(grid, ''));
  expect(control('both').disabled).toBe(false);
  control('both').click();
  expect(shape()).toEqual([3, 3]);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(shape()).toEqual([2, 2]);
});
