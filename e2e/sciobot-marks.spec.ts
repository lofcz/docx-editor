import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/marks';

const cases = [
  ['bold', 'Bold', 'Control+b'],
  ['italic', 'Italic', 'Control+i'],
  ['underline', 'Underline', 'Control+u'],
  ['strikethrough', 'Strikethrough', 'Control+Shift+x'],
  ['subscript', 'Subscript', 'Control+,'],
  ['superscript', 'Superscript', 'Control+.'],
  ['code', 'Inline code', 'Control+e'],
] as const;
for (const [mark, label, shortcut] of cases)
  test(`${mark} toolbar, mixed state, shortcut, real typing and shared saved native properties`, async ({
    page,
  }) => {
    await page.goto('/marks.html');
    await page.waitForFunction(() => !!window.marksControls);
    await page.evaluate(() => window.marksControls.select(0, 5));
    const before = await page.evaluate(() => window.marksControls.state());
    const button = page
      .locator('[data-pane="0"]')
      .getByRole('button', { name: label, exact: true });
    await button.click();
    await page.evaluate(() => window.marksControls.settle());
    const applied = await page.evaluate(() => window.marksControls.state());
    expect(applied.selection).toEqual(before.selection);
    expect(applied.marks[mark]).toBe('on');
    await expect(button).toHaveAttribute('aria-pressed', 'true');
    const saved = await page.evaluate(() => window.marksControls.saved());
    expect(saved[0]).toEqual(saved[1]);
    expect(await page.evaluate(() => window.marksControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.marksControls.state())).marks[mark]).toBe('off');
    await page.evaluate(() => window.marksControls.select(0, 5));
    await page.keyboard.press(shortcut);
    expect((await page.evaluate(() => window.marksControls.state())).selection).toEqual(
      before.selection
    );
    expect((await page.evaluate(() => window.marksControls.state())).marks[mark]).toBe('on');
    await page.evaluate(() => window.marksControls.select(0, 10));
    await expect(button).toHaveAttribute('aria-pressed', 'mixed');
    await page.evaluate(() => window.marksControls.select(10));
    await page.keyboard.press(shortcut);
    await page.keyboard.type('!');
    expect((await page.evaluate(() => window.marksControls.state())).text).toBe('Alpha beta!');
    await page.evaluate(() => window.marksControls.select(10, 11));
    expect((await page.evaluate(() => window.marksControls.state())).marks[mark]).toBe('on');
    await page.evaluate(() => window.marksControls.settle());
    const final = await page.evaluate(() => window.marksControls.saved());
    expect(final[0]).toEqual(final[1]);
  });
