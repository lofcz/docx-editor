import { expect, test, type Page } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/fragments';

const original = [
  'Intro',
  'Outer start',
  'Inner text',
  ...Array.from({ length: 56 }, (_, index) => `Outer line ${index + 1}`),
  'Outer end',
  'Form fallback',
  'Tail',
];
const painted = (page: Page, peer = 0) =>
  page.locator(`[data-pane="${peer}"] .docx-paragraph-fragment`).allTextContents();
const formPage = (page: Page, peer = 0) =>
  page.locator(`[data-pane="${peer}"] [data-form-page]`).getAttribute('data-form-page');
const saved = async (page: Page) => {
  const digests = await page.evaluate(() => window.blockFragments.savedDigests());
  expect(digests[0]).toEqual(digests[1]);
};

for (const first of [0, 1] as const) {
  test(`page moves retain focused widget state and native continuations through reconnect ${first}`, async ({
    page,
  }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.goto('/fragments.html');
    const alice = page.locator('[data-pane="0"]');
    const notes = alice.getByLabel('Local notes');
    await expect(notes).toHaveCount(1);
    await page.evaluate(() => window.blockFragments.settle());
    for (const peer of [0, 1]) {
      expect(await painted(page, peer)).toEqual(original);
      const pages = await page
        .locator(`[data-pane="${peer}"] [data-container="outer"]`)
        .evaluateAll((nodes) =>
          nodes.map((node) => Number((node as HTMLElement).dataset.nativePage))
        );
      expect(pages).toEqual([0, 1, 2]);
    }
    await alice.getByRole('button', { name: 'Count 0' }).click();
    await notes.fill('Local state');
    await page.keyboard.press('ArrowLeft');
    const input = await notes.elementHandle();
    const before = await formPage(page);
    await page.evaluate(() => window.blockFragments.pause());
    expect(await page.evaluate(() => window.blockFragments.push())).toBe(true);
    await expect(notes).toBeFocused();
    await expect(notes).toHaveValue('Local state');
    await expect(alice.getByRole('button', { name: 'Count 1' })).toHaveCount(1);
    await expect.poll(() => formPage(page)).not.toBe(before);
    expect(await formPage(page, 1)).toBe(before);
    expect(await notes.evaluate((node, previous) => node === previous, input)).toBe(true);
    expect(
      await notes.evaluate((node: HTMLInputElement) => [node.selectionStart, node.selectionEnd])
    ).toEqual([10, 10]);
    await page.keyboard.type('!');
    await expect(notes).toHaveValue('Local stat!e');
    expect(await page.evaluate(() => window.blockFragments.insert(' peer', 'Inner text', 1))).toBe(
      true
    );
    await page.evaluate((first) => window.blockFragments.reconnect(first), first);
    await expect(notes).toBeFocused();
    expect(await formPage(page, 1)).toBe(await formPage(page));
    const changed = original.map((text) => (text === 'Inner text' ? 'Inner text peer' : text));
    for (const peer of [0, 1]) expect(await painted(page, peer)).toEqual(changed);
    expect(await page.evaluate(() => window.blockFragments.events)).toEqual([
      { mounted: 1, unmounted: 0 },
      { mounted: 1, unmounted: 0 },
    ]);
    await expect(page.locator('[data-pane="1"]').getByLabel('Local notes')).toHaveValue('');
    await saved(page);

    // Document history restores the document caret and preserves the widget state.
    await alice
      .locator('.docx-paragraph-fragment')
      .filter({ hasText: /^Tail$/ })
      .click();
    await page.keyboard.press('Control+z');
    await expect.poll(() => formPage(page)).toBe(before);
    await expect(notes).toHaveValue('Local stat!e');
    for (const peer of [0, 1]) expect(await painted(page, peer)).toEqual(changed);
    await page.keyboard.press('Control+Shift+z');
    await expect.poll(() => formPage(page)).not.toBe(before);
    expect(await notes.evaluate((node, previous) => node === previous, input)).toBe(true);
    await expect(alice.getByRole('button', { name: 'Count 1' })).toHaveCount(1);

    await page.evaluate(() =>
      window.blockFragments.select(['Inner text peer', 6], ['Inner text peer', 10])
    );
    const remote = await page.evaluate(() => window.blockFragments.remoteGeometry());
    expect(remote).not.toBeNull();
    expect(remote![0]).toEqual(remote![1]);
    for (const caret of remote![0]!)
      expect(caret).toEqual({
        x: expect.any(Number),
        y: expect.any(Number),
        page: expect.any(Number),
      });
    await saved(page);

    // A peer deletes the moved widget while Alice edits its local input.
    await notes.click();
    await page.keyboard.press('End');
    await page.evaluate(() => window.blockFragments.pause());
    await page.keyboard.type(' more');
    expect(await page.evaluate(() => window.blockFragments.removeBlock('form', 1))).toBe(true);
    await page.evaluate((first) => window.blockFragments.reconnect(first), first);
    await expect(page.getByLabel('Local notes')).toHaveCount(0);
    for (const peer of [0, 1])
      expect(await painted(page, peer)).toEqual(changed.filter((text) => text !== 'Form fallback'));
    expect(await page.evaluate(() => window.blockFragments.events)).toEqual([
      { mounted: 1, unmounted: 1 },
      { mounted: 1, unmounted: 1 },
    ]);
    await saved(page);
    expect(errors).toEqual([]);
  });
}
