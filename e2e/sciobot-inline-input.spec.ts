import { expect, test, type Page } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/inline-host';

const errors = new WeakMap<Page, string[]>();
test.beforeEach(({ page }) => {
  const messages: string[] = [];
  errors.set(page, messages);
  page.on('pageerror', (error) => messages.push(error.message));
});
test.afterEach(({ page }) => expect(errors.get(page)).toEqual([]));

const api = <T>(page: Page, run: (host: Window['inlineHost']) => T | Promise<T>) =>
  page.evaluate(`(${run.toString()})(window.inlineHost)`) as Promise<T>;
async function ready(page: Page, reverse = false) {
  await page.goto(`/inline-host.html${reverse ? '?reverse' : ''}`);
  await expect(page.locator('[data-sciobot-inline-measure]')).toHaveCount(2);
  await api(page, (h) => h.settle());
}
async function copy(page: Page) {
  await api(page, (h) => h.select(7));
  await page.keyboard.press('Shift+ArrowRight');
  expect(await api(page, (h) => h.selection())).toMatchObject({
    anchor: { offset: 7 },
    head: { offset: 13 },
  });
  await page.keyboard.press('Control+c');
  const transfer = await api(page, (h) => h.transfers.at(-1)!);
  expect(transfer['text/plain']).toBe('Source');
  expect(
    JSON.parse(transfer['application/x-sciobot-blocks+json']!).blocks[0].payload.attributes
  ).toEqual({ name: 'Initial' });
}

test('real arrows, clipboard, replacement, and history keep complete native inline atoms', async ({
  page,
  context,
}) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await ready(page);
  const original = await api(page, (h) => h.digest());
  await api(page, (h) => h.select(7));
  await page.keyboard.press('ArrowRight');
  expect(await api(page, (h) => h.selection())).toMatchObject({ head: { offset: 13 } });
  await page.keyboard.press('ArrowLeft');
  expect(await api(page, (h) => h.selection())).toMatchObject({ head: { offset: 7 } });
  await copy(page);
  await api(page, (h) => h.select(19));
  await page.keyboard.press('Control+v');
  await api(page, (h) => h.settle());
  expect(await api(page, (h) => h.text())).toBe('Before Source afterSource');
  const blocks = await api(page, (h) => h.blocks());
  expect(blocks).toHaveLength(3);
  expect(blocks[1]!.id).not.toBe('source-1');
  expect(blocks[1]!.attributes).toEqual({ name: 'Initial' });
  const pasted = await api(page, (h) => h.digest());
  await page.keyboard.press('Control+z');
  await api(page, (h) => h.settle());
  expect(await api(page, (h) => h.digest())).toEqual(original);
  await page.keyboard.press('Control+Shift+z');
  await api(page, (h) => h.settle());
  expect(await api(page, (h) => h.digest())).toEqual(pasted);
  // A stale API selection inside the fallback still replaces the entire displayed chip.
  await api(page, (h) => h.range(8, 11));
  await page.keyboard.type('X');
  await api(page, (h) => h.settle());
  expect(await api(page, (h) => h.text())).toBe('Before X afterSource');
  expect(await api(page, (h) => h.blocks())).toEqual([blocks[1], blocks[2]]);
  await page.keyboard.press('Control+z');
  await api(page, (h) => h.settle());
  expect(await api(page, (h) => h.digest())).toEqual(pasted);
  await api(page, (h) => h.range(11, 8));
  await page.keyboard.press('Control+x');
  await api(page, (h) => h.settle());
  expect(await api(page, (h) => h.text())).toBe('Before  afterSource');
  expect((await api(page, (h) => h.transfers.at(-1)!))['text/plain']).toBe('Source');
  await page.keyboard.press('Control+z');
  await api(page, (h) => h.settle());
  expect(await api(page, (h) => h.digest())).toEqual(pasted);
  const input = page.locator('[data-pane="0"] input').first();
  await input.fill('local');
  await input.press('Control+a');
  await input.press('Control+x');
  await expect(input).toHaveValue('');
  await input.press('Control+v');
  await expect(input).toHaveValue('local');
  expect(await api(page, (h) => h.digest())).toEqual(pasted);
  const saved = await api(page, (h) => h.reopen());
  expect(saved.saved).toEqual(saved.reopened);
});

for (const reverse of [false, true])
  for (const receiver of [0, 1] as const)
    test(`inline clipboard preserves concurrent peer typing, deletion, history, and saved content ${reverse}/${receiver}`, async ({
      page,
      context,
    }) => {
      await context.grantPermissions(['clipboard-read', 'clipboard-write']);
      await ready(page, reverse);
      await copy(page);
      await api(page, (h) => h.pause());
      await api(page, (h) => h.select(19));
      await page.keyboard.press('Control+v');
      await api(page, (h) => h.select(0, 1));
      await page.keyboard.type('Bob ');
      await page.evaluate((receiver) => window.inlineHost.reconnect(receiver), receiver);
      const pasted = await api(page, (h) => h.blocks());
      expect(pasted).toHaveLength(3);
      for (const peer of [0, 1]) {
        expect(await page.evaluate((peer) => window.inlineHost.text(peer), peer)).toBe(
          'Bob Before Source afterSource'
        );
        expect(await page.evaluate((peer) => window.inlineHost.blocks(peer), peer)).toEqual(pasted);
      }
      expect(await api(page, (h) => h.exec('undo'))).toBe(true);
      expect(await api(page, (h) => h.text())).toBe('Bob Before Source after');
      expect(await api(page, (h) => h.blocks())).toEqual([pasted[0], pasted[2]]);
      expect(await api(page, (h) => h.exec('redo'))).toBe(true);
      await api(page, (h) => h.pause());
      await api(page, (h) => h.select(17));
      await page.keyboard.press('Backspace');
      await api(page, (h) => h.select(h.text(1)!.length, 1));
      await page.keyboard.type(' peer');
      await page.evaluate((receiver) => window.inlineHost.reconnect(receiver), receiver);
      for (const peer of [0, 1]) {
        expect(await page.evaluate((peer) => window.inlineHost.text(peer), peer)).toBe(
          'Bob Before  afterSource peer'
        );
        expect(await page.evaluate((peer) => window.inlineHost.blocks(peer), peer)).toEqual([
          pasted[1],
          pasted[2],
        ]);
      }
      expect(await api(page, (h) => h.exec('undo'))).toBe(true);
      expect(await api(page, (h) => h.text())).toBe('Bob Before Source afterSource peer');
      expect(await api(page, (h) => h.blocks())).toEqual(pasted);
      expect(await api(page, (h) => h.exec('redo'))).toBe(true);
      const saved = await api(page, (h) => h.savedDigests());
      expect(saved[0]).toEqual(saved[1]);
      const remote = await api(page, (h) => h.remoteGeometry());
      expect(remote).not.toBeNull();
      for (const key of ['x', 'y', 'height'] as const)
        expect(remote![key]).toBeCloseTo(remote!.expected[key], 1);
    });
