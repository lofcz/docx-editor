import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/retention';

test('twenty scroll and document-switch cycles retain parked state and release roots and observers', async ({
  page,
}) => {
  test.setTimeout(120_000);
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/retention.html');
  const input = page.locator('[data-pane="0"]').getByLabel('Retained note');
  await expect(input).toBeVisible();
  await input.fill('Saved locally');
  await page.locator('[data-pane="0"]').getByRole('button', { name: 'Count 0' }).click();
  const identity = await input.elementHandle();
  const baseline = await page.evaluate(() => window.blockRetention.stats());
  expect(baseline.roots).toBe(2);
  expect(baseline.listenerRoots).toBe(2);
  for (let cycle = 0; cycle < 20; cycle++) {
    await page.evaluate(() => window.blockRetention.scroll(true));
    await expect.poll(() => page.evaluate(() => window.blockRetention.parked())).toBe(true);
    expect(await identity!.evaluate((node) => node.isConnected)).toBe(false);
    await page.evaluate(() => window.blockRetention.scroll(false));
    await expect(input).toBeVisible();
    await expect(input).toHaveValue('Saved locally');
    expect(await input.evaluate((node, old) => node === old, identity)).toBe(true);
    await expect(
      page.locator('[data-pane="0"]').getByRole('button', { name: 'Count 1' })
    ).toBeVisible();
    expect(await page.evaluate(() => window.blockRetention.stats())).toEqual(baseline);
  }
  for (let cycle = 0; cycle < 20; cycle++) {
    expect(await page.evaluate(() => window.blockRetention.cycle())).toEqual({
      roots: 0,
      observers: 0,
      observations: 0,
      listenerRoots: 0,
    });
    expect(await page.evaluate(() => window.blockRetention.stats())).toEqual(baseline);
  }
  expect(await page.evaluate(() => window.blockRetention.stop())).toEqual({
    roots: 0,
    observers: 0,
    observations: 0,
    listenerRoots: 0,
  });
  expect(errors).toEqual([]);
});

for (const receiver of [0, 1] as const)
  test(`parked roots survive peer input, deletion, history, and saved convergence ${receiver}`, async ({
    page,
  }) => {
    await page.goto('/retention.html');
    const notes = page.locator('[data-pane="0"]').getByLabel('Retained note');
    await notes.fill('Local');
    await page.evaluate(() => window.blockRetention.scroll(true));
    await expect.poll(() => page.evaluate(() => window.blockRetention.parked())).toBe(true);
    await page.evaluate(() => window.blockRetention.pause());
    await page.evaluate(() => window.blockRetention.insert(' peer'));
    await page.evaluate((receiver) => window.blockRetention.reconnect(receiver), receiver);
    await page.evaluate(() => window.blockRetention.scroll(false));
    await expect(notes).toHaveValue('Local');
    expect(await page.evaluate(() => window.blockRetention.text())).toBe(
      await page.evaluate(() => window.blockRetention.text(1))
    );
    expect(await page.evaluate(() => window.blockRetention.remote())).not.toBeNull();
    await page.evaluate(() => window.blockRetention.pause());
    expect(await page.evaluate(() => window.blockRetention.remove())).toBe(true);
    await page.evaluate(() => window.blockRetention.insert(' later'));
    await page.evaluate((receiver) => window.blockRetention.reconnect(receiver), receiver);
    expect(await page.evaluate(() => window.blockRetention.stats().roots)).toBe(0);
    expect(await page.evaluate(() => window.blockRetention.stats().listenerRoots)).toBe(0);
    expect(await page.evaluate(() => window.blockRetention.undo())).toBe(true);
    await expect(notes).toHaveValue('');
    expect(await page.evaluate(() => window.blockRetention.text())).toContain(
      'Editable body peer later'
    );
    expect(await page.evaluate(() => window.blockRetention.redo())).toBe(true);
    expect(await page.evaluate(() => window.blockRetention.stats().roots)).toBe(0);
    expect(await page.evaluate(() => window.blockRetention.undo())).toBe(true);
    await expect(notes).toHaveValue('');
    const saved = await page.evaluate(() => window.blockRetention.saved());
    expect(saved[0]).toEqual(saved[1]);
    expect(await page.evaluate(() => window.blockRetention.stop())).toEqual({
      roots: 0,
      observers: 0,
      observations: 0,
      listenerRoots: 0,
    });
  });
