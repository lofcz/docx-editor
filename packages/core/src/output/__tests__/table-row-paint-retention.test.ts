import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
import { expect, test } from 'bun:test';
import { applyTreeOp } from '../../store/store/tree-ops.ts';
import type { OoxmlElement } from '@docx-editor.dev/core/store';
import { createLayoutSession } from '../../layout/layout-session.ts';
import { createParagraphLayoutCache } from '../../layout/layout-cache.ts';
import { readTableStructure } from '../../layout/semantic-table.ts';
import { load, lay, para, row, table } from '../../layout/__tests__/table-row-keep-fixtures.ts';
import { paintSemanticLayout, paintSemanticLayoutWithAuthorSlots } from '../semantic-paint.ts';

test('typing preserves unchanged row elements and matches a cold paint', () => {
  const part = load(table(Array.from({ length: 30 }, (_, i) => row(`Row${i}`))));
  const session = createLayoutSession();
  const cache = createParagraphLayoutCache();
  const layout = lay(part, 15, { session, cache });
  const container = document.createElement('div');
  paintSemanticLayout(container, layout, { scale: 1 });
  const before = Array.from(container.querySelectorAll('.docx-table-row'));
  const paragraphs = (rowElement: Element) =>
    Array.from(rowElement.querySelectorAll('.docx-paragraph-fragment'));
  const [edited, untouched] = paragraphs(before[3]!);
  const editedText = edited!.querySelector('.layout-run-text')!.firstChild;
  const node = part.root.children
    .find((n) => n.kind === 'body')!
    .children.find((n) => n.kind === 'table')! as OoxmlElement;
  const id = readTableStructure(node, 310, 0)!.rows[3]!.cells[0]!.blocks[0]!.id;
  const op = applyTreeOp(part, { op: 'insertText', paragraphId: id, offset: 0, text: 'X' });
  if (!op.ok) throw Error(op.reason);
  const updated = lay(op.part, 15, { session, cache });
  paintSemanticLayout(container, updated, { scale: 1 });
  const after = Array.from(container.querySelectorAll('.docx-table-row'));
  expect(after[0]).toBe(before[0]);
  // Preserve the text node inside the edited cell as well as the neighboring cells.
  expect(after[3]).toBe(before[3]);
  expect(paragraphs(after[3]!)[0]).toBe(edited);
  expect(edited!.querySelector('.layout-run-text')!.firstChild).toBe(editedText);
  expect(editedText!.textContent).toContain('X');
  expect(paragraphs(after[3]!)[1]).toBe(untouched);
  expect(after[4]).toBe(before[4]);
  const cold = document.createElement('div');
  paintSemanticLayout(cold, lay(op.part), { scale: 1 });
  expect(container.innerHTML).toBe(cold.innerHTML);
  paintSemanticLayout(container, updated, { scale: 2 });
  expect(container.querySelector('.docx-table-row')).not.toBe(after[0]);
});

/** Painted rows by row id; a row split across pages keeps its first fragment's element. */
function rowElements(container: HTMLElement): Map<string, Element> {
  const rows = new Map<string, Element>();
  for (const element of container.querySelectorAll('.docx-table-row')) {
    const id = (element as HTMLElement).dataset.rowId!;
    if (!rows.has(id)) rows.set(id, element);
  }
  return rows;
}

/** The page index each row is painted on. */
function rowPages(container: HTMLElement): Map<string, number> {
  const pages = new Map<string, number>();
  Array.from(container.children).forEach((page, index) => {
    for (const element of page.querySelectorAll('.docx-table-row'))
      pages.set((element as HTMLElement).dataset.rowId!, index);
  });
  return pages;
}

function rowEditSession(body: string) {
  let part = load(body);
  const session = createLayoutSession();
  const cache = createParagraphLayoutCache();
  const container = document.createElement('div');
  paintSemanticLayout(container, lay(part, 15, { session, cache }), { scale: 1 });
  return {
    container,
    get part() {
      return part;
    },
    apply(op: Parameters<typeof applyTreeOp>[1]) {
      const edit = applyTreeOp(part, op);
      if (!edit.ok) throw Error(edit.reason);
      part = edit.part;
      paintSemanticLayout(container, lay(part, 15, { session, cache }), { scale: 1 });
      const cold = document.createElement('div');
      paintSemanticLayout(cold, lay(part), { scale: 1 });
      expect(container.innerHTML).toBe(cold.innerHTML);
    },
  };
}

const tableOf = (part: Parameters<typeof applyTreeOp>[0]) =>
  part.root.children
    .find((n) => n.kind === 'body')!
    .children.find((n) => n.kind === 'table')! as OoxmlElement;

test('a row insert keeps every later row element, including rows that move to the next page', () => {
  const run = rowEditSession(table(Array.from({ length: 40 }, (_, i) => row(`Row${i}`))));
  const before = rowElements(run.container);
  const pagesBefore = rowPages(run.container);
  const node = tableOf(run.part);
  const rows = node.children.filter((n) => n.kind === 'tableRow');
  run.apply({ op: 'insertTableRow', tableId: node.id, rowId: rows[2]!.id, where: 'below' });
  const after = rowElements(run.container);
  const pagesAfter = rowPages(run.container);
  let crossed = 0;
  for (const [id, element] of before) {
    expect(after.get(id)).toBe(element);
    if (pagesAfter.get(id) !== pagesBefore.get(id)) crossed += 1;
  }
  // The insert pushed one row over each later page break.
  expect(crossed).toBeGreaterThan(2);
});

test('a page break above a table moves its rows to the next page with their elements', () => {
  const run = rowEditSession(
    '<w:p><w:r><w:t>Before</w:t></w:r></w:p>' +
      table(Array.from({ length: 30 }, (_, i) => row(`Row${i}`)))
  );
  const before = rowElements(run.container);
  const body = run.part.root.children.find((n) => n.kind === 'body')! as OoxmlElement;
  const preamble = body.children.find((n) => n.kind === 'paragraph')!;
  run.apply({ op: 'insertPageBreak', paragraphId: preamble.id, offset: 6 });
  const after = rowElements(run.container);
  let kept = 0;
  for (const [id, element] of before) if (after.get(id) === element) kept += 1;
  // A row now opening a page takes its top border there, but keeps its element.
  expect(kept).toBe(before.size);
});

test('a borrowed table keeps its rows when its continuation becomes the first fragment', () => {
  const grid = table(Array.from({ length: 8 }, (_, i) => row(`Row${i}`)));
  const part = load(para('Before', 8) + grid);
  const layout = lay(part);
  const container = document.createElement('div');
  paintSemanticLayout(container, layout, { scale: 1 });
  const sheet = container.children[1]!;
  const old = sheet.querySelector<HTMLElement>('.docx-table-fragment')!;
  expect(old.dataset.fragmentIndex).toBe('1');
  const cellText = old.querySelector('.layout-run-text')!;
  const rows = [...old.querySelectorAll('.docx-table-row')];
  const marker = document.createComment('table home');
  const slot = document.createElement('section');
  old.before(marker, slot);
  slot.append(old);
  const next = lay(load(para('Before', 12) + grid));
  let offered = false;
  paintSemanticLayoutWithAuthorSlots(container, next, { scale: 1 }, undefined, (_kept, reserve) => {
    offered = reserve(
      1,
      (changed) => changed.has(old),
      (root) => (root === old ? marker : root)
    );
  });
  expect(offered).toBe(true);
  expect(slot.firstElementChild).toBe(old);
  expect(old.dataset.fragmentIndex).toBe('0');
  expect(old.contains(cellText)).toBe(true);
  expect([...old.querySelectorAll('.docx-table-row')].slice(-rows.length)).toEqual(rows);
  marker.replaceWith(old);
  slot.remove();
  const cold = document.createElement('div');
  paintSemanticLayout(cold, next, { scale: 1 });
  expect(container.innerHTML).toBe(cold.innerHTML);
});
