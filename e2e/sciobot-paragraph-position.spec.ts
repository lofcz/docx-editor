import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/paragraph-position';
for (const action of ['start', 'left', 'center', 'right', 'end', 'justify', 'increase', 'decrease'])
  test(`${action} preserves browser selection, native history, imported neighbor and saved convergence`, async ({
    page,
  }) => {
    await page.goto('/paragraph-position.html');
    await page.waitForFunction(() => !!window.paragraphPositionControls);
    await page.evaluate(() => window.paragraphPositionControls.select(1, 0, 5));
    const before = await page.evaluate(() => window.paragraphPositionControls.state());
    const label =
      action === 'increase'
        ? 'Increase indent'
        : action === 'decrease'
          ? 'Decrease indent'
          : `Align ${action}`;
    await page.locator('[data-pane="0"]').getByRole('button', { name: label, exact: true }).click();
    await page.evaluate(() => window.paragraphPositionControls.settle());
    const after = await page.evaluate(() => window.paragraphPositionControls.state());
    expect(after.selection).toEqual(before.selection);
    expect(after.paragraphs[0]).toEqual(before.paragraphs[0]);
    if (action === 'increase' || action === 'decrease')
      expect(after.indent?.left).toBe(action === 'increase' ? 360 : 0);
    else
      expect(after.alignment).toBe(
        action === 'start'
          ? 'left'
          : action === 'end'
            ? 'right'
            : action === 'justify'
              ? 'both'
              : action
      );
    const saved = await page.evaluate(() => window.paragraphPositionControls.saved());
    expect(saved[0]).toEqual(saved[1]);
    expect(saved[0]).toEqual(after.fingerprint);
    expect(await page.evaluate(() => window.paragraphPositionControls.undo())).toBe(true);
    expect(
      (await page.evaluate(() => window.paragraphPositionControls.state())).fingerprint
    ).toEqual(before.fingerprint);
    await page.keyboard.type('Hello');
    expect(
      (await page.evaluate(() => window.paragraphPositionControls.state())).paragraphs[1]!.text
    ).toBe('Hello beta');
    await page.evaluate(() => window.paragraphPositionControls.settle());
    const final = await page.evaluate(() => window.paragraphPositionControls.saved());
    expect(final[0]).toEqual(final[1]);
  });
