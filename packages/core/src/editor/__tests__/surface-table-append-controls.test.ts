import { afterEach, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import {
  resetTableCommandPlannerCallCount,
  tableCommandPlannerCallCount,
} from '../table-command-plan.ts';
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
async function mount(owned = false, markup = table, author?: string) {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    author,
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
  await Promise.resolve();
  const control = (kind: string) =>
    container.querySelector<HTMLButtonElement>(`[data-table-append="${kind}"]`)!;
  const shape = () => {
    const table = surface.layout().pages[0]!.fragments.find((item) => item.kind === 'table');
    if (table?.kind !== 'table') throw Error('Missing table');
    return [table.rows.length, table.rows[0]!.cells.length];
  };
  return { editor, surface, control, shape, container };
}

test('the corner appends a row and column in one undo transaction', async () => {
  const { editor, control, shape } = await mount();
  control('both').click();
  expect(shape()).toEqual([3, 3]);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(shape()).toEqual([2, 2]);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  expect(shape()).toEqual([3, 3]);
});
test('append buttons retain keyboard focus across repeated operations', async () => {
  const { control, shape } = await mount();
  const button = control('row');
  button.focus();
  button.click();
  await Promise.resolve();
  button.click();
  await Promise.resolve();
  expect(shape()).toEqual([4, 2]);
  expect(document.activeElement).toBe(button);
  control('column').click();
  expect(shape()).toEqual([4, 3]);
});
test('owned tables and viewing mode hide append buttons', async () => {
  expect((await mount(true)).control('row')).toBeNull();
  const { editor, control } = await mount();
  editor.setEditingMode('viewing');
  expect(control('row')).toBeNull();
});

test('append controls support authored tables without a column grid', async () => {
  const { editor, control, shape } = await mount(false, table.replace(grid, ''));
  expect(control('both').disabled).toBe(false);
  control('both').click();
  expect(shape()).toEqual([3, 3]);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(shape()).toEqual([2, 2]);
});

test('edge appends preserve locked cell content and obey enclosing control locks', async () => {
  const math =
    '<w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:r><w:t>2+2</w:t></w:r></w:sdtContent></w:sdt>';
  const worksheet = table.replace(p('D'), `<w:p>${math}</w:p>`);
  const { editor, control, shape } = await mount(false, worksheet);
  expect(control('both').disabled).toBe(false);
  control('both').click();
  expect(shape()).toEqual([3, 3]);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(shape()).toEqual([2, 2]);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  const xml = strFromU8(unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!);
  expect(xml.match(/2\+2/g)).toHaveLength(1);
  expect(xml).toContain('sdtContentLocked');
  const enclosed = `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>${worksheet}</w:sdtContent></w:sdt>`;
  expect((await mount(false, enclosed)).control('both').disabled).toBe(true);
});

test('append buttons follow the bottom and right edges and refresh visible labels', async () => {
  const { surface, control } = await mount();
  const row = control('row');
  expect(parseFloat(row.style.width)).toBeGreaterThan(0);
  expect(row.style.height).toBe('24px');
  const column = control('column');
  expect(column.style.width).toBe('24px');
  expect(parseFloat(column.style.height)).toBeGreaterThan(0);
  expect(parseFloat(column.style.left)).toBeCloseTo(
    parseFloat(row.style.left) + parseFloat(row.style.width)
  );
  expect(parseFloat(row.style.top)).toBeCloseTo(
    parseFloat(column.style.top) + parseFloat(column.style.height)
  );
  surface.setTableInteractionLabel((key) => `localized:${key}`);
  expect(row.title).toBe('localized:table.insertRowBelow');
  expect(row.getAttribute('aria-label')).toBe(row.title);
});

test('append controls create tracked row and column changes while suggesting', async () => {
  const { editor, control, shape } = await mount(false, table, 'Writer');
  editor.surface!.setEditingMode('suggest');
  await Promise.resolve();
  expect(control('both').disabled).toBe(false);
  control('both').click();
  expect(shape()).toEqual([3, 3]);
  const xml = strFromU8(unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!);
  expect(xml).toContain('w:author="Writer"');
  expect(xml).toContain('w:ins');
  expect(xml).toContain('w:cellIns');
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(shape()).toEqual([2, 2]);
});

test('append controls refuse suggesting without an author before activation', async () => {
  const { editor, control, shape } = await mount();
  editor.surface!.setEditingMode('suggest');
  await Promise.resolve();
  expect(control('both').disabled).toBe(true);
  expect(control('both').title).toContain('author');
  control('both').click();
  expect(shape()).toEqual([2, 2]);
});

test('append permission probes share row and column plans after a document command', async () => {
  const { editor, surface, control } = await mount();
  resetTableCommandPlannerCallCount();
  expect(editor.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
  expect(tableCommandPlannerCallCount()).toBe(1);
  await Promise.resolve();
  expect(tableCommandPlannerCallCount()).toBe(3);
  resetTableCommandPlannerCallCount();
  const paragraphId = surface.session
    .paragraphIds()
    .find((id) => paragraphTextFromLayout(surface.layout(), id) === 'A')!;
  const position = { paragraphId, offset: 1 };
  surface.setSelection({ anchor: position, head: position });
  await Promise.resolve();
  expect(tableCommandPlannerCallCount()).toBe(0);
  control('both').click();
  expect(tableCommandPlannerCallCount()).toBe(2);
});
