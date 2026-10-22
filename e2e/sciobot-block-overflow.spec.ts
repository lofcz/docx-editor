import { expect, test, type Page } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/measurement';

const api = <T>(page: Page, run: (m: Window['blockMeasurement']) => T | Promise<T>) =>
  page.evaluate(`(${run.toString()})(window.blockMeasurement)`) as Promise<T>;
const settle = async (page: Page) => {
  await api(page, (m) => m.settle());
  await page.waitForTimeout(150);
};

test('a short fallback reserves only a page-bounded viewport, including after growth and shrink', async ({
  page,
}) => {
  await page.goto('/measurement.html?overflow=short');
  await expect(page.locator('[data-sciobot-block-measure]')).toHaveCount(2);
  await api(page, (m) => m.setRows(90));
  await settle(page);
  for (const peer of [0, 1]) {
    const viewport = await page.evaluate((peer) => window.blockMeasurement.viewport(peer), peer);
    expect(viewport.scrollHeight).toBeGreaterThan(viewport.height * 2);
    expect(viewport.bottom).toBeLessThanOrEqual(viewport.pageBottom + 1);
    const geometry = await page.evaluate((peer) => window.blockMeasurement.geometry(peer), peer);
    expect(geometry.tailTop!).toBeGreaterThanOrEqual(geometry.widgetBottom! - 0.5);
  }
  await api(page, (m) => m.setRows(2));
  await settle(page);
  expect((await api(page, (m) => m.viewport())).scrollHeight).toBe(
    (await api(page, (m) => m.viewport())).height
  );
});

for (const receiver of [0, 1] as const)
  test(`long listening transcript scrolls, edits, exports and converges; receiver ${receiver}`, async ({
    page,
  }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.goto('/measurement.html?overflow=long');
    await expect(page.locator('[data-sciobot-block-measure]')).toHaveCount(2);
    await settle(page);
    for (const peer of [0, 1]) {
      const box = await page.evaluate((peer) => window.blockMeasurement.viewport(peer), peer);
      expect(box.roots).toBe(1);
      expect(box.nativePages).toBeGreaterThan(1);
      expect(box.scrollHeight).toBeGreaterThan(box.height * 2);
      expect(box.bottom).toBeLessThanOrEqual(box.pageBottom + 1);
      const geometry = await page.evaluate((peer) => window.blockMeasurement.geometry(peer), peer);
      expect(geometry.tailTop!).toBeGreaterThan(geometry.widgetBottom!);
    }
    const alice = page.locator('[data-pane="0"]');
    const beforeScroll = await api(page, (m) => m.savedDigests());
    const last = alice.getByRole('textbox', { name: 'Turn 90', exact: true });
    // Playwright scrolls both the widget and the editor pane. Real input edits its last turn.
    await last.click();
    expect(await api(page, (m) => m.savedDigests())).toEqual(beforeScroll);
    await page.keyboard.press('End');
    await page.keyboard.type(' Alice');
    await api(page, (m) => m.pause());
    await last.press('Tab');
    expect(await api(page, (m) => m.insert(' Bob', 'Tail', 1))).toBe(true);
    await page.evaluate((first) => window.blockMeasurement.reconnect(first), receiver);
    await settle(page);
    const expected = Array.from(
      { length: 90 },
      (_, i) => `Speaker ${i + 1}: Listening transcript turn ${i + 1}.${i === 89 ? ' Alice' : ''}`
    );
    for (const peer of [0, 1]) {
      const saved = await page.evaluate((peer) => window.blockMeasurement.savedTexts(peer), peer);
      expect(saved).toEqual(['Intro', ...expected, 'Tail Bob', 'After']);
    }
    const digests = await api(page, (m) => m.savedDigests());
    expect(digests[0]).toEqual(digests[1]);
    await api(page, (m) => m.select('Tail Bob', 2));
    const remote = await api(page, (m) => m.remoteY());
    expect(remote![0]).not.toBeNull();
    expect(remote![0]).toBeCloseTo(remote![1]!, 3);
    expect(await api(page, (m) => m.exec('undo'))).toBe(true);
    await settle(page);
    expect((await api(page, (m) => m.savedTexts()))[90]).toBe(expected[89]!.replace(' Alice', ''));
    expect(await api(page, (m) => m.textOfTail())).toBe('Tail Bob');
    expect(await api(page, (m) => m.exec('redo'))).toBe(true);
    await settle(page);
    expect((await api(page, (m) => m.savedTexts()))[90]).toBe(expected[89]);

    // Native typing still targets following painted content after widget scrolling.
    const tail = alice.locator(
      `.docx-paragraph-fragment[data-paragraph-id="${await api(page, (m) => m.idOf('Tail Bob'))}"]`
    );
    await tail.click({ position: { x: 4, y: 6 } });
    await page.keyboard.press('End');
    await page.keyboard.type('!');
    await settle(page);
    expect(await api(page, (m) => m.textOfTail(1))).toBe('Tail Bob!');

    await api(page, (m) => m.pause());
    expect(await api(page, (m) => m.removeBlock(1))).toBe(true);
    expect(
      await page.evaluate((text) => window.blockMeasurement.insert(' peer', text, 0), expected[89]!)
    ).toBe(true);
    await page.evaluate((first) => window.blockMeasurement.reconnect(first), receiver);
    await settle(page);
    await expect(page.locator('[data-sciobot-block-measure]')).toHaveCount(0);
    expect(await api(page, (m) => m.statuses())).not.toContain('error');
    const final = await api(page, (m) => m.savedDigests());
    expect(final[0]).toEqual(final[1]);
    expect(await api(page, (m) => m.savedTexts())).toEqual(['Intro', 'Tail Bob!', 'After']);
    expect(errors).toEqual([]);
  });
