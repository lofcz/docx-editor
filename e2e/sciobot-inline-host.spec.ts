import { expect, test, type Page } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/inline-host';

const api = <T>(page: Page, run: (host: Window['inlineHost']) => T | Promise<T>) =>
  page.evaluate(`(${run.toString()})(window.inlineHost)`) as Promise<T>;
async function stable(page: Page) {
  await api(page, (h) => h.settle());
  await page.waitForTimeout(100);
  await api(page, (h) => h.settle());
}
async function geometry(page: Page) {
  const result = await api(page, (h) => h.geometry());
  expect(result).not.toBeNull();
  for (const key of ['x', 'y', 'width', 'height', 'baseline'] as const)
    expect(result![key]).toBeCloseTo(result!.expected[key], 1);
  expect(result!.hits).toEqual([result!.range.start, result!.range.end]);
}
async function mouseEdges(page: Page) {
  const box = (await api(page, (h) => h.geometry()))!;
  for (const [edge, offset] of [
    ['left', box.range.start],
    ['right', box.range.end],
  ] as const) {
    expect(
      await page.evaluate(
        ({ x, y }) => Boolean(document.elementFromPoint(x, y)?.closest('[data-docx-host-ui]')),
        box.mouse[edge]
      )
    ).toBe(false);
    await page.mouse.click(box.mouse[edge].x, box.mouse[edge].y);
    await stable(page);
    const selection = await api(page, (h) => h.selection());
    expect(selection).toMatchObject({
      anchor: { paragraphId: box.range.paragraphId, offset },
      head: { paragraphId: box.range.paragraphId, offset },
    });
  }
}
test('source chip retains identity, focused input, native selection, and exact geometry across providers and zoom', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/inline-host.html');
  const alice = page.locator('[data-pane="0"]'),
    chip = alice.locator('[data-sciobot-inline-measure]');
  await expect(chip).toBeVisible();
  await stable(page);
  await geometry(page);
  await mouseEdges(page);
  const revision = await api(page, (h) => h.revision());
  await chip.getByRole('button').click();
  await chip.getByRole('textbox').fill('local');
  const input = await chip.getByRole('textbox').elementHandle();
  await api(page, (h) => {
    h.providers('citation');
    h.zoom(1.5);
    h.refresh();
  });
  await stable(page);
  await expect(chip.getByRole('button')).toHaveText('citation 1');
  await expect(chip.getByRole('textbox')).toHaveValue('local');
  expect(
    await input!.evaluate((element) => element.isConnected && document.activeElement === element)
  ).toBe(true);
  expect(await api(page, (h) => h.events[0])).toEqual({ mounted: 1, unmounted: 0 });
  expect(await api(page, (h) => h.revision())).toBe(revision);
  await geometry(page);
  await mouseEdges(page);
  await api(page, (h) => h.select(7));
  await page.keyboard.type('left ');
  await stable(page);
  expect(await api(page, (h) => h.text())).toBe('Before left Source after');
  const selection = await api(page, (h) => h.selection());
  await api(page, (h) => {
    h.providers('updated source');
    h.refresh();
  });
  await stable(page);
  expect(await api(page, (h) => h.selection())).toEqual(selection);
  await page.keyboard.type('next ');
  await stable(page);
  expect(await api(page, (h) => h.text())).toBe('Before left next Source after');
  await api(page, (h) => h.select(23));
  await page.keyboard.type(' right');
  await stable(page);
  expect(await api(page, (h) => h.text())).toBe('Before left next Source right after');
  await geometry(page);
  const reopened = await api(page, (h) => h.reopen());
  expect(reopened.saved).toEqual(reopened.reopened);
  expect(reopened.measured).toBe(true);
  expect(reopened.text).toBe('Before left next Source right after');
  await expect(page.locator('[data-sciobot-inline-measure]')).toHaveCount(3);
  await api(page, (h) => h.dispose());
  await stable(page);
  await expect(alice.locator('[data-sciobot-block-host]')).toHaveCount(0);
  await expect(alice.locator('.docx-paragraph-fragment').first()).toContainText(
    'Before left next Source right after'
  );
  expect(errors).toEqual([]);
});

for (const reverse of [false, true])
  for (const receiver of [0, 1] as const)
    test(`local boxes survive shared payload, native edits, history, reconnect, and deletion ${reverse}/${receiver}`, async ({
      page,
    }) => {
      const errors: string[] = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(`/inline-host.html${reverse ? '?reverse' : ''}`);
      await expect(page.locator('[data-sciobot-inline-measure]')).toHaveCount(2);
      await stable(page);
      const widths = await api(page, (h) => [h.geometry()!.width, h.geometry(1)!.width]);
      expect(widths[0]).not.toBe(widths[1]);
      await api(page, (h) => h.pause());
      expect(await api(page, (h) => h.rename('Shared'))).toBe(true);
      expect(await api(page, (h) => h.insert(' Bob'))).toBe(true);
      await page.evaluate((receiver) => window.inlineHost.reconnect(receiver), receiver);
      await stable(page);
      for (const peer of [0, 1]) {
        await expect(page.locator(`[data-pane="${peer}"] [data-payload-name]`)).toHaveText(
          'Shared'
        );
        expect(await page.evaluate((peer) => window.inlineHost.text(peer), peer)).toBe(
          'Before Source after Bob'
        );
      }
      await api(page, (h) => h.select(22, 1));
      await stable(page);
      const remote = await api(page, (h) => h.remoteGeometry());
      expect(remote).not.toBeNull();
      for (const key of ['x', 'y', 'height'] as const)
        expect(remote![key]).toBeCloseTo(remote!.expected[key], 1);
      await expect(page.locator('[data-pane="0"] .docx-remote-caret').first()).toBeVisible();
      expect(await api(page, (h) => h.exec('undo'))).toBe(true);
      await stable(page);
      expect(await api(page, (h) => h.text())).toBe('Before Source after Bob');
      await expect(page.locator('[data-pane="0"] [data-payload-name]')).toHaveText('Initial');
      expect(await api(page, (h) => h.exec('redo'))).toBe(true);
      await stable(page);
      await geometry(page);
      expect(await api(page, (h) => h.events)).toEqual([
        { mounted: 1, unmounted: 0 },
        { mounted: 1, unmounted: 0 },
      ]);
      const saved = await api(page, (h) => h.savedDigests());
      expect(saved[0]).toEqual(saved[1]);
      await api(page, (h) => h.pause());
      expect(await api(page, (h) => h.remove())).toBe(true);
      expect(await api(page, (h) => h.insert('!'))).toBe(true);
      await page.evaluate((receiver) => window.inlineHost.reconnect(receiver), receiver);
      await stable(page);
      await expect(page.locator('[data-sciobot-inline-measure]')).toHaveCount(0);
      expect(await api(page, (h) => h.text())).toBe('Before  after Bob!');
      expect(await api(page, (h) => h.text(1))).toBe('Before  after Bob!');
      expect(await api(page, (h) => h.events)).toEqual([
        { mounted: 1, unmounted: 1 },
        { mounted: 1, unmounted: 1 },
      ]);
      const deleted = await api(page, (h) => h.savedDigests());
      expect(deleted[0]).toEqual(deleted[1]);
      expect(errors).toEqual([]);
    });
