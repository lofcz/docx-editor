import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/line-height';
for (const value of [1, 1.2, 1.5, 2, 3])
  test(`line height ${value} retains exact spacing, browser selection, typing and shared saved output`, async ({
    page,
  }) => {
    await page.goto('/line-height.html');
    await page.waitForFunction(() => !!window.lineHeightControls);
    await page.evaluate(() => window.lineHeightControls.select(0, 0));
    await expect(page.locator('[data-pane="0"] output')).toHaveText('Exactly 12.55 pt');
    await page.evaluate(() => window.lineHeightControls.select(1, 0, 5));
    const before = await page.evaluate(() => window.lineHeightControls.state());
    await page
      .locator('[data-pane="0"]')
      .getByRole('button', { name: String(value), exact: true })
      .click();
    await page.evaluate(() => window.lineHeightControls.settle());
    const after = await page.evaluate(() => window.lineHeightControls.state());
    expect(after.selection).toEqual(before.selection);
    expect(after.spacing).toEqual({ rule: 'multiple', value });
    expect(after.paragraphs[0]).toEqual(before.paragraphs[0]);
    const saved = await page.evaluate(() => window.lineHeightControls.saved());
    expect(saved[0]).toEqual(saved[1]);
    expect(saved[0]).toEqual(after.fingerprint);
    expect(await page.evaluate(() => window.lineHeightControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.lineHeightControls.state())).fingerprint).toEqual(
      before.fingerprint
    );
    await page.keyboard.type('Hello');
    expect((await page.evaluate(() => window.lineHeightControls.state())).paragraphs[1]!.text).toBe(
      'Hello beta'
    );
    await page.evaluate(() => window.lineHeightControls.settle());
    const final = await page.evaluate(() => window.lineHeightControls.saved());
    expect(final[0]).toEqual(final[1]);
    expect((await page.evaluate(() => window.lineHeightControls.state())).paragraphs[0]).toEqual(
      before.paragraphs[0]
    );
  });
