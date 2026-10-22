import { expect, test, type CDPSession, type Page } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/selection';

const compose = async (cdp: CDPSession, commit = true) => {
  await cdp.send('Input.imeSetComposition', { text: 'か', selectionStart: 1, selectionEnd: 1 });
  await cdp.send('Input.imeSetComposition', { text: 'かな', selectionStart: 2, selectionEnd: 2 });
  if (commit) await cdp.send('Input.insertText', { text: '仮名' });
};
const textAt = (page: Page, id: string, peer = 0) =>
  page.evaluate(([id, peer]) => window.blockSelection.textAt({ paragraphId: id }, peer), [
    id,
    peer,
  ] as const);
const ready = async (page: Page) => {
  await page.goto('/selection.html');
  await expect(page.locator('[data-pane="0"] [data-sciobot-native-slot]')).toHaveCount(4);
  return page.evaluate(() => window.blockSelection.idOf('Inner text'));
};
const nativeEnd = async (page: Page, id: string) => {
  await page.locator(`[data-pane="0"] .docx-paragraph-fragment[data-paragraph-id="${id}"]`).click();
  await page.keyboard.press('End');
};
const saved = async (page: Page) => {
  const digests = await page.evaluate(() => window.blockSelection.savedDigests());
  expect(digests[0]).toEqual(digests[1]);
};

test('native slot composition commits once and retains its caret through a remote repaint', async ({
  page,
  context,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  const id = await ready(page);
  await nativeEnd(page, id);
  const cdp = await context.newCDPSession(page);
  await compose(cdp);
  await expect.poll(() => textAt(page, id)).toBe('Inner text仮名');
  await page.evaluate(() => window.blockSelection.insert(' peer', 'Tail', 1));
  for (const peer of [0, 1]) expect(await textAt(page, id, peer)).toBe('Inner text仮名');
  const position = { paragraphId: id, offset: 12 };
  expect(await page.evaluate(() => window.blockSelection.model())).toEqual({
    anchor: position,
    head: position,
  });
  await page.keyboard.type('!');
  expect(await textAt(page, id)).toBe('Inner text仮名!');
  await page.evaluate(() =>
    window.blockSelection.select(['Inner text仮名!', 10], ['Inner text仮名!', 12])
  );
  const geometry = await page.evaluate(() => window.blockSelection.remoteGeometry());
  expect(geometry).not.toBeNull();
  expect(geometry![0]).toEqual(geometry![1]);
  for (const caret of geometry![0]!)
    expect(caret).toEqual({
      x: expect.any(Number),
      y: expect.any(Number),
      page: expect.any(Number),
    });
  await saved(page);
  expect(errors).toEqual([]);
});

for (const first of [0, 1] as const) {
  test(`composition survives concurrent edits, history, reconnect ${first}, and deletion`, async ({
    page,
    context,
  }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    const id = await ready(page);
    const cdp = await context.newCDPSession(page);
    await nativeEnd(page, id);
    await page.evaluate(() => window.blockSelection.pause());
    await compose(cdp);
    expect(await textAt(page, id)).toBe('Inner text仮名');
    expect(await textAt(page, id, 1)).toBe('Inner text');
    await page.evaluate(() => window.blockSelection.insert(' peer', 'Tail', 1));
    await page.evaluate((first) => window.blockSelection.reconnect(first), first);
    for (const peer of [0, 1]) expect(await textAt(page, id, peer)).toBe('Inner text仮名');
    expect(await page.evaluate(() => window.blockSelection.exec('undo', 0))).toBe(true);
    for (const peer of [0, 1]) expect(await textAt(page, id, peer)).toBe('Inner text');
    expect(await page.evaluate(() => window.blockSelection.exec('redo', 0))).toBe(true);
    for (const peer of [0, 1]) {
      expect(await textAt(page, id, peer)).toBe('Inner text仮名');
      expect(
        await page.evaluate((peer) => window.blockSelection.text('Tail peer', peer), peer)
      ).toBe('Tail peer');
    }
    await saved(page);

    const notes = page.locator('[data-pane="0"]').getByLabel('Widget notes');
    await notes.click();
    await page.keyboard.press('End');
    const revisions = await page.evaluate(() => [
      window.blockSelection.revision(0),
      window.blockSelection.revision(1),
    ]);
    await compose(cdp);
    await expect(notes).toHaveValue('Widget words仮名');
    expect(
      await page.evaluate(() => [
        window.blockSelection.revision(0),
        window.blockSelection.revision(1),
      ])
    ).toEqual(revisions);
    await page.evaluate(() => window.blockSelection.insert(' repaint', 'Tail peer', 1));
    await expect(notes).toBeFocused();
    await expect(notes).toHaveValue('Widget words仮名');
    expect(
      await notes.evaluate((node: HTMLTextAreaElement) => [node.selectionStart, node.selectionEnd])
    ).toEqual([14, 14]);
    await page.keyboard.type('!');
    await expect(notes).toHaveValue('Widget words仮名!');
    await expect(page.locator('[data-pane="1"]').getByLabel('Widget notes')).toHaveValue(
      'Widget words'
    );

    // Delete the widget while its browser composition is active.
    await page.evaluate(() => window.blockSelection.pause());
    await compose(cdp, false);
    expect(await page.evaluate(() => window.blockSelection.removeBlock('form', 1))).toBe(true);
    await page.evaluate((first) => window.blockSelection.reconnect(first), first);
    await expect(page.getByLabel('Widget notes')).toHaveCount(0);
    await cdp.send('Input.imeSetComposition', { text: '', selectionStart: 0, selectionEnd: 0 });
    for (const peer of [0, 1]) {
      expect(await textAt(page, id, peer)).toBe('Inner text仮名');
      expect(
        await page.evaluate((peer) => window.blockSelection.text('Tail peer repaint', peer), peer)
      ).toBe('Tail peer repaint');
    }
    await saved(page);
    expect(errors).toEqual([]);
  });
}
