import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/colors';
for (const channel of ['foreground', 'background'] as const)
  for (const clear of [false, true])
    test(`${channel} ${clear ? 'clear' : 'RGB'} preserves native selection, painted text, independent colors, history and shared saving`, async ({
      page,
    }) => {
      await page.goto('/colors.html');
      await page.waitForFunction(() => !!window.colorControls);
      await page.evaluate(() => window.colorControls.select(1, 0, 5));
      const before = await page.evaluate(() => window.colorControls.state());
      const label = `${channel} ${clear ? 'clear' : '#17A2B8'}`;
      await page
        .locator('[data-pane="0"]')
        .getByRole('button', { name: label, exact: true })
        .click();
      await page.evaluate(() => window.colorControls.settle());
      const after = await page.evaluate(() => window.colorControls.state());
      expect(after.selection).toEqual(before.selection);
      expect(after.paragraphs[0]).toEqual(before.paragraphs[0]);
      const other = channel === 'foreground' ? 'background' : 'foreground';
      expect(after[other]).toEqual(before[other]);
      if (!clear) {
        expect(after[channel]).toEqual({ value: '#17A2B8', mixed: false });
        const paint = await page
          .locator('[data-pane="0"] [data-paragraph-id]')
          .filter({ hasText: 'Alpha beta' })
          .first()
          .evaluate(
            (element, kind) =>
              Array.from(element.querySelectorAll('span'))
                .filter((node) => node.textContent === 'Alpha')
                .map((node) => {
                  const style = node.ownerDocument.defaultView!.getComputedStyle(node);
                  return kind === 'foreground' ? style.color : style.backgroundColor;
                }),
            channel
          );
        expect(paint).toContain('rgb(23, 162, 184)');
      }
      const saved = await page.evaluate(() => window.colorControls.saved());
      expect(saved[0]).toEqual(saved[1]);
      expect(saved[0]).toEqual(after.fingerprint);
      expect(await page.evaluate(() => window.colorControls.undo())).toBe(true);
      expect((await page.evaluate(() => window.colorControls.state())).fingerprint).toEqual(
        before.fingerprint
      );
      await page.keyboard.type('Hello');
      expect((await page.evaluate(() => window.colorControls.state())).paragraphs[1]!.text).toBe(
        'Hello beta'
      );
      await page.evaluate(() => window.colorControls.settle());
      const final = await page.evaluate(() => window.colorControls.saved());
      expect(final[0]).toEqual(final[1]);
    });
