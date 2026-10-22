import { expect, test, type Page } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/measurement';

/** Evaluate a closure-free callback against the page API. */
const api = <T>(page: Page, run: (measurement: Window['blockMeasurement']) => T | Promise<T>) =>
  page.evaluate(`(${run.toString()})(window.blockMeasurement)`) as Promise<T>;
/** Consecutive distinct values, rounded to 1/1000 pt. */
const moves = (values: number[]) =>
  values
    .map((value) => Math.round(value * 1000) / 1000)
    .filter((value, index, all) => index === 0 || value !== all[index - 1]);
const stable = async (page: Page) => {
  await api(page, (m) => m.settle());
  await page.waitForTimeout(150);
  await api(page, (m) => m.settle());
};

test('an expanding widget moves following native text once per edit, and undo needs no measurement steps', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/measurement.html');
  for (const pane of [0, 1])
    await expect(page.locator(`[data-pane="${pane}"] [data-sciobot-block-measure]`)).toHaveCount(1);
  await stable(page);
  const base = await api(page, (m) => m.lineY('Tail'));
  expect(await api(page, (m) => m.lineY('Tail', 1))).toBe(base);
  expect(await api(page, (m) => m.canUndo())).toBe(false);

  // Real clicks on the widget's own control: each is one payload edit.
  const alice = page.locator('[data-pane="0"]');
  const ys: number[] = [base];
  for (const rows of [1, 2]) {
    await api(page, (m) => m.resetHistory());
    await alice.getByRole('button', { name: 'Add row' }).click();
    await stable(page);
    expect(await api(page, (m) => m.rows())).toBe(rows);
    expect(await api(page, (m) => m.rows(1))).toBe(rows);
    const y = await api(page, (m) => m.lineY('Tail'));
    expect(y).toBeGreaterThan(ys.at(-1)!);
    // Moved once: the old position, then the new one, never back and forth.
    const history = moves(await api(page, (m) => m.tailHistory()));
    expect(history.at(-1)).toBe(Math.round(y * 1000) / 1000);
    expect(history.length).toBeLessThanOrEqual(2);
    if (history.length === 2) expect(history[0]).toBe(Math.round(ys.at(-1)! * 1000) / 1000);
    ys.push(y);
    // The peer measured its own widget and reserved the same room.
    expect(await api(page, (m) => m.lineY('Tail', 1))).toBeCloseTo(y, 3);
    for (const peer of [0, 1]) {
      const { widgetBottom, tailTop } = await page.evaluate(
        (peer) => window.blockMeasurement.geometry(peer),
        peer
      );
      expect(widgetBottom).not.toBeNull();
      expect(tailTop!).toBeGreaterThanOrEqual(widgetBottom! - 0.5);
    }
  }
  // Settled: later turns publish no layout.
  await api(page, (m) => m.resetHistory());
  await stable(page);
  expect(await api(page, (m) => m.tailHistory())).toHaveLength(0);

  // Undo from the document with the keyboard: one step per click, none for measurement.
  const after = alice.locator(
    `.docx-paragraph-fragment[data-paragraph-id="${await api(page, (m) => m.idOf('After'))}"]`
  );
  await after.click({ position: { x: 4, y: 6 } });
  await page.keyboard.press('Control+z');
  await stable(page);
  expect(await api(page, (m) => m.rows())).toBe(1);
  expect(await api(page, (m) => m.lineY('Tail'))).toBeCloseTo(ys[1]!, 3);
  await page.keyboard.press('Control+z');
  await stable(page);
  expect(await api(page, (m) => m.rows())).toBe(0);
  expect(await api(page, (m) => m.lineY('Tail'))).toBeCloseTo(base, 3);
  expect(await api(page, (m) => m.canUndo())).toBe(false);
  await page.keyboard.press('Control+Shift+z');
  await page.keyboard.press('Control+Shift+z');
  await stable(page);
  expect(await api(page, (m) => m.rows())).toBe(2);
  expect(await api(page, (m) => m.lineY('Tail'))).toBeCloseTo(ys[2]!, 3);
  expect(await api(page, (m) => m.lineY('Tail', 1))).toBeCloseTo(ys[2]!, 3);

  // Native typing in the moved paragraph lands where it is painted.
  const tail = alice.locator(
    `.docx-paragraph-fragment[data-paragraph-id="${await api(page, (m) => m.idOf('Tail'))}"]`
  );
  await tail.click({ position: { x: 2, y: 6 } });
  await page.keyboard.press('End');
  await page.keyboard.type('!');
  await stable(page);
  expect(await api(page, (m) => m.textOfTail(1))).toBe('Tail!');
  expect(errors).toEqual([]);
});

for (const receiver of [0, 1] as const)
  test(`concurrent widget growth, peer text, remote carets, and deletion converge; receiver ${receiver}`, async ({
    page,
  }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.goto('/measurement.html');
    for (const pane of [0, 1])
      await expect(page.locator(`[data-pane="${pane}"] [data-sciobot-block-measure]`)).toHaveCount(
        1
      );
    await stable(page);
    const base = await api(page, (m) => m.lineY('Tail'));

    await api(page, (m) => m.pause());
    const alice = page.locator('[data-pane="0"]');
    await alice.getByRole('button', { name: 'Add row' }).click();
    await alice.getByRole('button', { name: 'Add row' }).click();
    expect(await api(page, (m) => m.insert(' Bob', 'Tail', 1))).toBe(true);
    await stable(page);
    expect(await api(page, (m) => m.lineY('Tail'))).toBeGreaterThan(base);
    expect(await api(page, (m) => m.lineY('Tail Bob', 1))).toBe(base);
    await page.evaluate((first) => window.blockMeasurement.reconnect(first), receiver);
    await stable(page);
    expect(await api(page, (m) => m.statuses())).not.toContain('error');
    const grown = await api(page, (m) => m.lineY('Tail Bob'));
    expect(grown).toBeGreaterThan(base);
    expect(await api(page, (m) => m.lineY('Tail Bob', 1))).toBeCloseTo(grown, 3);
    expect(await api(page, (m) => [m.rows(), m.rows(1)])).toEqual([2, 2]);
    const [savedAlice, savedBob] = await api(page, (m) => m.savedDigests());
    expect(savedAlice).toEqual(savedBob);

    // Bob's caret paints on the moved line in both sessions.
    await api(page, (m) => m.select('Tail Bob', 2));
    await stable(page);
    const remote = await api(page, (m) => m.remoteY());
    expect(remote![0]).toBeCloseTo(grown, 3);
    expect(remote![1]).toBeCloseTo(grown, 3);
    await expect(alice.locator('.docx-remote-caret').first()).toBeVisible();

    // Alice's history keeps Bob's text and both reservations follow it.
    expect(await api(page, (m) => m.exec('undo'))).toBe(true);
    await stable(page);
    const shrunk = await api(page, (m) => m.lineY('Tail Bob'));
    expect(shrunk).toBeLessThan(grown);
    expect(await api(page, (m) => m.lineY('Tail Bob', 1))).toBeCloseTo(shrunk, 3);
    expect(await api(page, (m) => m.exec('redo'))).toBe(true);
    await stable(page);
    for (const peer of [0, 1])
      expect(
        await page.evaluate((peer) => window.blockMeasurement.lineY('Tail Bob', peer), peer)
      ).toBeCloseTo(grown, 3);

    // Deletion during Alice's growth: the widget and its room disappear in both sessions.
    await api(page, (m) => m.pause());
    expect(await api(page, (m) => m.removeBlock(1))).toBe(true);
    expect(await api(page, (m) => m.setRows(4))).toBe(true);
    await page.evaluate((first) => window.blockMeasurement.reconnect(first), receiver);
    await stable(page);
    expect(await api(page, (m) => m.statuses())).not.toContain('error');
    expect(await api(page, (m) => [m.rows(), m.rows(1)])).toEqual([null, null]);
    await expect(page.locator('[data-sciobot-block-measure]')).toHaveCount(0);
    const final = await api(page, (m) => m.lineY('Tail Bob'));
    expect(final).toBeLessThan(base);
    expect(await api(page, (m) => m.lineY('Tail Bob', 1))).toBeCloseTo(final, 3);
    const digests = await api(page, (m) => m.savedDigests());
    expect(digests[0]).toEqual(digests[1]);
    expect(errors).toEqual([]);
  });
