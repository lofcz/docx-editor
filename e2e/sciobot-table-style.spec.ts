import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/tables';
for (const direction of ['horizontal', 'vertical', 'rectangle'] as const)
  test(`${direction} native cell selection merges, restores unequal slots, types and keeps exact history and Word output`, async ({
    page,
  }) => {
    await page.goto('/tables.html?styles=1');
    await page.waitForFunction(() => !!window.tableControls);
    const rowEnd = direction === 'horizontal' ? 1 : 2,
      colEnd = direction === 'vertical' ? 0 : 1;
    await page.evaluate(
      ({ rowEnd, colEnd }) => window.tableControls.selectCells([1, rowEnd], [0, colEnd]),
      { rowEnd, colEnd }
    );
    const before = await page.evaluate(() => window.tableControls.state());
    expect(before.cellSelection!.cellIds).toHaveLength(rowEnd * (colEnd + 1));
    const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Table cells' });
    await toolbar.getByRole('button', { name: 'Merge cells', exact: true }).click();
    await page.evaluate(() => window.tableControls.settle());
    const merged = await page.evaluate(() => window.tableControls.state());
    expect((await page.evaluate(() => window.tableControls.columns())).widths).toEqual([
      1200, 1800, 2400,
    ]);
    expect(await page.evaluate(() => window.tableControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
      before.fingerprint
    );
    expect(await page.evaluate(() => window.tableControls.redo())).toBe(true);
    expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
      merged.fingerprint
    );
    const owner = await page.evaluate(
      () => window.tableControls.columns().rows[1]!.cells[0]!.paragraphs[0]!.id
    );
    const index = await page.evaluate(
      (id) => window.tableControls.state().paragraphs.findIndex((p) => p.id === id),
      owner
    );
    await page.evaluate((index) => window.tableControls.select(index, 0), index);
    await toolbar.getByRole('button', { name: 'Split cell', exact: true }).click();
    await page.evaluate(() => window.tableControls.settle());
    const split = await page.evaluate(() => window.tableControls.state()),
      shape = await page.evaluate(() => window.tableControls.columns());
    expect(shape.widths).toEqual([1200, 1800, 2400]);
    expect(shape.rows[1]!.cells[0]!.width).toBe(1200);
    expect(shape.rows[1]!.cells[0]!.span).toBe(1);
    expect(await page.evaluate(() => window.tableControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
      merged.fingerprint
    );
    expect(await page.evaluate(() => window.tableControls.redo())).toBe(true);
    expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
      split.fingerprint
    );
    await page.keyboard.type('Live');
    const typed = await page.evaluate(() => window.tableControls.state());
    expect(typed.paragraphs.find((p) => p.id === owner)!.text).toContain('Live');
    await page.evaluate(() => window.tableControls.settle());
    const saved = await page.evaluate(() => window.tableControls.saved());
    expect(saved[0]).toEqual(saved[1]);
    expect(saved[0]).toEqual(typed.fingerprint);
  });
for (const merge of ['none', 'horizontal', 'vertical'] as const)
  test(`${merge} table appearance pointer controls retain text, native geometry, exact history and Word output`, async ({
    page,
  }) => {
    await page.goto(`/tables.html?styles=1&merge=${merge}`);
    await page.waitForFunction(() => !!window.tableControls);
    const cell = await page.evaluate(
      () => window.tableControls.columns().rows[1]!.cells[1]!.paragraphs[0]!
    );
    const index = await page.evaluate(
      (id) => window.tableControls.state().paragraphs.findIndex((p) => p.id === id),
      cell.id
    );
    await page.evaluate((index) => window.tableControls.select(index, 2), index);
    const before = await page.evaluate(() => window.tableControls.state());
    const toolbar = page
      .locator('[data-pane="0"]')
      .getByRole('toolbar', { name: 'Table appearance' });
    await toolbar.locator('input[type="color"]').fill('#1a2b3c');
    await toolbar.getByRole('spinbutton', { name: 'Column width (pt)', exact: true }).fill('105');
    await toolbar.getByRole('spinbutton', { name: 'Row height (pt)', exact: true }).fill('36');
    await toolbar.getByRole('spinbutton', { name: 'Header rows', exact: true }).fill('2');
    const snapshots = [before.fingerprint];
    for (const name of [
      'Apply cell background',
      'Apply column width',
      'Apply row height',
      'All borders',
      'Apply header rows',
    ]) {
      await toolbar.getByRole('button', { name, exact: true }).click();
      await page.evaluate(() => window.tableControls.settle());
      snapshots.push((await page.evaluate(() => window.tableControls.state())).fingerprint);
    }
    const after = await page.evaluate(() => window.tableControls.columns());
    expect(after.widths).toEqual([1200, 2100, 2400]);
    expect(after.rows.map((r) => r.header)).toEqual([true, true, false]);
    const texts = await page.evaluate(() => window.tableControls.state());
    for (const p of before.paragraphs)
      expect(texts.paragraphs.find((q) => q.id === p.id)!.text).toBe(p.text);
    const saved = await page.evaluate(() => window.tableControls.saved());
    expect(saved[0]).toEqual(saved[1]);
    expect(saved[0]).toEqual(snapshots.at(-1)!);
    for (let i = snapshots.length - 2; i >= 0; i--) {
      expect(await page.evaluate(() => window.tableControls.undo())).toBe(true);
      expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
        snapshots[i]!
      );
    }
    for (let i = 1; i < snapshots.length; i++) {
      expect(await page.evaluate(() => window.tableControls.redo())).toBe(true);
      expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
        snapshots[i]!
      );
    }
    await page.keyboard.type('Live');
    const typed = await page.evaluate(() => window.tableControls.state());
    expect(typed.paragraphs.find((p) => p.id === cell.id)!.text).toContain('Live');
    await page.evaluate(() => window.tableControls.settle());
    const final = await page.evaluate(() => window.tableControls.saved());
    expect(final[0]).toEqual(final[1]);
    expect(final[0]).toEqual(typed.fingerprint);
  });

test('a real drag across native cells supplies the rectangle consumed by the merge control', async ({
  page,
}) => {
  await page.goto('/tables.html?styles=1');
  await page.waitForFunction(() => !!window.tableControls);
  const shape = await page.evaluate(() => window.tableControls.columns());
  const first = shape.rows[1]!.cells[0]!.id,
    last = shape.rows[2]!.cells[1]!.id;
  const startCell = page.locator('[data-pane="0"]').locator(`[data-cell-id="${first}"]`).first();
  const endCell = page.locator('[data-pane="0"]').locator(`[data-cell-id="${last}"]`).first();
  await endCell.scrollIntoViewIfNeeded();
  const start = await startCell.boundingBox(),
    end = await endCell.boundingBox();
  expect(start).not.toBeNull();
  expect(end).not.toBeNull();
  await page.mouse.move(start!.x + 12, start!.y + start!.height / 2);
  await page.mouse.down();
  await page.mouse.move(end!.x + end!.width / 2, end!.y + end!.height / 2, { steps: 12 });
  await page.mouse.up();
  await expect
    .poll(async () =>
      page.evaluate(() => window.tableControls.state().cellSelection?.cellIds.length ?? 0)
    )
    .toBe(4);
  const before = await page.evaluate(() => window.tableControls.state());
  await page
    .locator('[data-pane="0"]')
    .getByRole('button', { name: 'Merge cells', exact: true })
    .click();
  await page.evaluate(() => window.tableControls.settle());
  const merged = await page.evaluate(() => window.tableControls.columns());
  expect(merged.rows[1]!.cells[0]!.span).toBe(2);
  expect(await page.evaluate(() => window.tableControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
    before.fingerprint
  );
});
