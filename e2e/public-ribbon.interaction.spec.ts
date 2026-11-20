import { expect, test } from '@playwright/test';

for (const adapter of ['react', 'vue'] as const) {
  test(`${adapter} public ribbon preserves the table caret and document scroll`, async ({
    page,
  }, info) => {
    const port = info.config.metadata[adapter + 'Port'];
    await page.setViewportSize({ width: 1360, height: 800 });
    await page.goto(`http://localhost:${port}/?fixture=repeated-table-header.docx`, {
      waitUntil: 'commit',
    });
    const ribbon = page.locator('[data-ribbon]');
    const tab = (name: string) => ribbon.getByRole('tab', { name, exact: true });
    const body = page.locator('.demo-viewport');
    const rows = page.locator('.docx-table-row');
    const cells = rows.locator('.docx-table-cell [data-start]');
    await expect(cells.first()).toBeVisible();
    await expect(ribbon).toHaveCount(1);
    await tab('Home').click();
    await expect(ribbon.getByRole('group', { name: 'Paragraph', exact: true })).toBeVisible();
    await tab('Home').focus();
    await page.keyboard.press('ArrowRight');
    await expect(tab('Insert')).toBeFocused();
    await expect(ribbon.locator('[data-slot="table.insert"]')).toBeVisible();
    await ribbon.getByRole('button', { name: 'Collapse ribbon' }).click();
    await expect(ribbon.getByRole('tabpanel')).toHaveCount(0);
    await ribbon.getByRole('button', { name: 'Expand ribbon' }).click();

    await cells.first().click();
    await tab('Table Design').click();
    await ribbon.getByRole('button', { name: 'Banded Grid Blue' }).click();
    await tab('Table Layout').click();
    const before = await rows.count();
    await ribbon.getByRole('button', { name: 'Insert row below' }).click();
    await expect.poll(() => rows.count()).toBeGreaterThan(before);

    await page.setViewportSize({ width: 390, height: 720 });
    await body.evaluate((el) => {
      el.scrollTop = 230;
    });
    const scroll = await body.evaluate((el) => el.scrollTop);
    expect(scroll).toBeGreaterThan(0);
    await tab('Table Layout').focus();
    await page.keyboard.press('Home');
    await expect(tab('Home')).toBeFocused();
    await page.keyboard.press('End');
    await expect(tab('Table Layout')).toBeFocused();
    await expect(tab('Table Layout')).toHaveAttribute('aria-selected', 'true');
    const strip = ribbon.getByRole('tablist');
    expect(await strip.evaluate((el) => el.scrollLeft)).toBeGreaterThan(0);
    const panel = ribbon.getByRole('tabpanel');
    expect(await panel.evaluate((el) => el.scrollWidth - el.clientWidth)).toBeLessThanOrEqual(1);
    const more = ribbon.getByRole('button', { name: 'More', exact: true });
    await more.focus();
    await page.keyboard.press('ArrowDown');
    await expect(ribbon.getByTestId('toolbar-overflow-panel')).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(ribbon.getByTestId('toolbar-overflow-panel')).toHaveCount(0);
    expect(await body.evaluate((el) => el.scrollTop)).toBe(scroll);
  });
}
