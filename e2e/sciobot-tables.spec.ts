import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/tables';

test('the 2x3 grid pick adopts native cell one and one undo restores exactly the original document', async ({
  page,
}) => {
  await page.goto('/tables.html');
  await page.waitForFunction(() => !!window.tableControls);
  await page.evaluate(() => window.tableControls.select(0, 5));
  const before = await page.evaluate(() => window.tableControls.state());
  const picker = page.locator('[data-pane="0"]').getByRole('grid', { name: 'Insert table' });
  await expect(picker.getByRole('gridcell')).toHaveCount(64);
  await picker.getByRole('button', { name: 'Insert 2 by 3 table', exact: true }).hover();
  await expect(picker.getByLabel('Table size')).toHaveText('2 × 3');
  await picker.getByRole('button', { name: 'Insert 2 by 3 table', exact: true }).click();
  await page.evaluate(() => window.tableControls.settle());
  const after = await page.evaluate(() => window.tableControls.state());
  expect(after.tables).toHaveLength(1);
  expect(after.tables[0]!.cells).toHaveLength(6);
  expect(after.anchor!.cellId).toBe(after.tables[0]!.cells[0]);
  expect(after.selection.head.offset).toBe(0);
  expect(after.paragraphs.filter((p) => p.text !== '')).toEqual(before.paragraphs);
  const saved = await page.evaluate(() => window.tableControls.saved());
  expect(saved[0]).toEqual(saved[1]);
  expect(saved[0]).toEqual(after.fingerprint);
  expect(await page.evaluate(() => window.tableControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
    before.fingerprint
  );
  expect(await page.evaluate(() => window.tableControls.redo())).toBe(true);
  expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
    after.fingerprint
  );
});

test('real Tab and Shift+Tab move through native cells, typing lands there, and last-cell Tab adds one row', async ({
  page,
}) => {
  await page.goto('/tables.html');
  await page.waitForFunction(() => !!window.tableControls);
  await page.evaluate(() => window.tableControls.select(0));
  await page
    .locator('[data-pane="0"]')
    .getByRole('button', { name: 'Insert 2 by 3 table', exact: true })
    .click();
  const inserted = await page.evaluate(() => window.tableControls.state());
  await page.keyboard.type('First');
  await page.keyboard.press('Tab');
  expect((await page.evaluate(() => window.tableControls.state())).anchor!.cellId).toBe(
    inserted.tables[0]!.cells[1]
  );
  await page.keyboard.type('Second');
  await page.keyboard.press('Shift+Tab');
  expect((await page.evaluate(() => window.tableControls.state())).anchor!.cellId).toBe(
    inserted.tables[0]!.cells[0]
  );
  await page.evaluate(() => window.tableControls.selectCell(5));
  const before = await page.evaluate(() => window.tableControls.state());
  await page.keyboard.press('Tab');
  const grown = await page.evaluate(() => window.tableControls.state());
  expect(grown.tables[0]!.cells).toHaveLength(9);
  expect(grown.anchor!.cellId).toBe(grown.tables[0]!.cells[6]);
  expect(await page.evaluate(() => window.tableControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
    before.fingerprint
  );
  await page.evaluate(() => window.tableControls.settle());
  const saved = await page.evaluate(() => window.tableControls.saved());
  expect(saved[0]).toEqual(saved[1]);
});
