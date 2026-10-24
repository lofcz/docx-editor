import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/tables';

for (const action of ['above', 'below', 'delete'] as const)
  test(`row ${action} pointer action keeps Word rows, native caret, one-step undo and shared output`, async ({
    page,
  }) => {
    await page.goto('/tables.html?rows=1');
    await page.waitForFunction(() => !!window.tableControls);
    await page.evaluate(() => window.tableControls.select(2, 3));
    const before = await page.evaluate(() => window.tableControls.state());
    const label = action === 'delete' ? 'Delete row' : `Insert row ${action}`;
    await page
      .locator('[data-pane="0"]')
      .getByRole('toolbar', { name: 'Table rows' })
      .getByRole('button', { name: label, exact: true })
      .click();
    await page.evaluate(() => window.tableControls.settle());
    const after = await page.evaluate(() => window.tableControls.state());
    expect(after.tables[0]!.cells).toHaveLength(action === 'delete' ? 4 : 8);
    expect(after.anchor!.tableId).toBe(before.tables[0]!.id);
    if (action !== 'delete')
      expect(after.anchor!.cellId).toBe(after.tables[0]!.cells[action === 'above' ? 2 : 4]);
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
    await page.keyboard.type('Live');
    const typed = await page.evaluate(() => window.tableControls.state());
    expect(typed.paragraphs.find((p) => p.id === typed.selection.head.paragraphId)!.text).toContain(
      'Live'
    );
    await page.evaluate(() => window.tableControls.settle());
    const final = await page.evaluate(() => window.tableControls.saved());
    expect(final[0]).toEqual(final[1]);
  });

test('row controls remain editable from peer source recovered inside cell one after rejecting AI content', async ({
  page,
}) => {
  await page.goto('/tables.html?rows=1&recovered=1');
  await page.waitForFunction(() => !!window.tableControls);
  await page.evaluate(() => window.tableControls.selectRecovered());
  const before = await page.evaluate(() => window.tableControls.state());
  expect(before.paragraphs.find((p) => p.id === before.selection.head.paragraphId)!.text).toBe(
    'Peer'
  );
  await page
    .locator('[data-pane="0"]')
    .getByRole('button', { name: 'Insert row below', exact: true })
    .click();
  const inserted = await page.evaluate(() => window.tableControls.state());
  expect(inserted.tables[0]!.cells).toHaveLength(8);
  expect(inserted.anchor!.cellId).toBe(inserted.tables[0]!.cells[2]);
  await page.keyboard.type('Live');
  const typed = await page.evaluate(() => window.tableControls.state());
  expect(typed.paragraphs.find((p) => p.id === typed.selection.head.paragraphId)!.text).toBe(
    'Live'
  );
  expect(typed.paragraphs.find((p) => p.id === before.selection.head.paragraphId)!.text).toBe(
    'Peer'
  );
  expect(await page.evaluate(() => window.tableControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
    inserted.fingerprint
  );
  expect(await page.evaluate(() => window.tableControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
    before.fingerprint
  );
  expect(await page.evaluate(() => window.tableControls.redo())).toBe(true);
  expect(await page.evaluate(() => window.tableControls.redo())).toBe(true);
  expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
    typed.fingerprint
  );
  const saved = await page.evaluate(() => window.tableControls.saved());
  expect(saved[0]).toEqual(saved[1]);
  // Word export removes the one internal recovery marker. Every other native value is exact.
  const marker = JSON.stringify([
    'element',
    'urn:docx-editor.dev:review:retired-paragraph:1',
    'retired',
    [],
    [],
  ]);
  expect(typed.fingerprint.split(marker)).toHaveLength(2);
  expect(saved[0]).toEqual(typed.fingerprint.replace(`${marker},`, ''));
  expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
    typed.fingerprint
  );
});
