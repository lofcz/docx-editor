import { expect, test, type Page } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/selection';

const ready = async (page: Page) => {
  await page.goto('/selection.html');
  await expect(page.locator('[data-pane="0"] [data-sciobot-native-slot]')).toHaveCount(4);
  return page.evaluate(() => window.blockSelection.idOf('Tail'));
};
const documentState = (page: Page, tailId: string, peer = 0) =>
  page.evaluate(
    ([tailId, peer]) => ({
      revision: window.blockSelection.revision(peer as number),
      tail: window.blockSelection.textAt({ paragraphId: tailId as string }, peer as number),
    }),
    [tailId, peer] as const
  );

test('typing, paste, cut, history, keys, and composition in a widget textarea never modify Word text', async ({
  page,
  context,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  const tailId = await ready(page);
  const alice = page.locator('[data-pane="0"]');
  const tail = alice.locator(`.docx-paragraph-fragment[data-paragraph-id="${tailId}"]`);
  await tail.click({ position: { x: 4, y: 8 } });
  await page.keyboard.press('End');
  await page.keyboard.type('!');
  const native = await documentState(page, tailId);
  expect(native.tail).toBe('Tail!');
  await expect(alice.locator('[data-docx-caret]')).toHaveCount(1);

  const notes = alice.getByLabel('Widget notes');
  await notes.click();
  await expect(notes).toBeFocused();
  // The painted document caret yields to the textarea's own native caret.
  await expect(alice.locator('[data-docx-caret]')).toHaveCount(0);
  expect(await notes.evaluate((node) => getComputedStyle(node).caretColor)).not.toBe(
    'rgba(0, 0, 0, 0)'
  );
  await page.keyboard.press('End');
  await page.keyboard.type(' typed');
  await expect(notes).toHaveValue('Widget words typed');
  await page.evaluate(() => navigator.clipboard.writeText('PASTE'));
  await page.keyboard.press('Control+v');
  await expect(notes).toHaveValue('Widget words typedPASTE');
  await page.keyboard.press('Control+a');
  await page.keyboard.press('Control+x');
  await expect(notes).toHaveValue('');
  expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('Widget words typedPASTE');
  await page.keyboard.press('Control+z');
  await expect(notes).toHaveValue('Widget words typedPASTE');
  await page.keyboard.press('Control+z');
  await expect(notes).toHaveValue('Widget words typed');
  await page.keyboard.press('Control+Shift+z');
  await expect(notes).toHaveValue('Widget words typedPASTE');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Control+b');
  await expect(notes).toHaveValue('Widget words typedPAST\n');
  const cdp = await context.newCDPSession(page);
  await cdp.send('Input.imeSetComposition', { text: 'か', selectionStart: 1, selectionEnd: 1 });
  await cdp.send('Input.insertText', { text: '漢' });
  await expect(notes).toHaveValue('Widget words typedPAST\n漢');
  await expect(notes).toBeFocused();
  expect(await documentState(page, tailId)).toEqual(native);

  // Native input still reaches the editor, and its history holds only document edits.
  await tail.click({ position: { x: 4, y: 8 } });
  await expect(alice.locator('[data-docx-caret]')).toHaveCount(1);
  await page.keyboard.press('End');
  await page.keyboard.type('?');
  expect((await documentState(page, tailId)).tail).toBe('Tail!?');
  await page.keyboard.press('Control+z');
  expect((await documentState(page, tailId)).tail).toBe('Tail!');
  await page.keyboard.press('Control+z');
  expect((await documentState(page, tailId)).tail).toBe('Tail');
  await expect(notes).toHaveValue('Widget words typedPAST\n漢');
  expect(errors).toEqual([]);
});

test('two real sessions keep widget input local through concurrent edits, history, reconnect, deletion, and saved convergence', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  const tailId = await ready(page);
  const alice = page.locator('[data-pane="0"]');
  const notes = alice.getByLabel('Widget notes');
  const bobBefore = await documentState(page, tailId, 1);
  await page.evaluate(() => window.blockSelection.pause());
  await notes.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' from Alice');
  expect(await page.evaluate(() => window.blockSelection.insert(' Bob', 'Inner text', 1))).toBe(
    true
  );
  await page.evaluate(() => window.blockSelection.reconnect(0));
  await expect(notes).toBeFocused();
  await expect(notes).toHaveValue('Widget words from Alice');
  for (const peer of [0, 1]) expect((await documentState(page, tailId, peer)).tail).toBe('Tail');
  for (const peer of [0, 1])
    expect(
      await page.evaluate((peer) => window.blockSelection.text('Inner text Bob', peer), peer)
    ).toBe('Inner text Bob');
  expect((await documentState(page, tailId, 1)).revision).toBeGreaterThan(bobBefore.revision);
  // Alice's widget typing created no document history; her undo has nothing to revert.
  expect(await page.evaluate(() => window.blockSelection.exec('undo', 0))).toBe(false);
  expect(await page.evaluate(() => window.blockSelection.text('Inner text Bob', 1))).toBe(
    'Inner text Bob'
  );
  expect(await page.evaluate(() => window.blockSelection.exec('undo', 1))).toBe(true);
  expect(await page.evaluate(() => window.blockSelection.text('Inner text', 0))).toBe('Inner text');
  expect(await page.evaluate(() => window.blockSelection.exec('redo', 1))).toBe(true);
  await expect(notes).toHaveValue('Widget words from Alice');

  // Bob removes the widget block while Alice is typing into its textarea.
  await notes.click();
  await page.keyboard.press('End');
  await page.evaluate(() => window.blockSelection.pause());
  await page.keyboard.type(' more');
  await page.evaluate(async () => {
    if (!(await window.blockSelection.removeBlock('form', 1))) throw new Error('Removal refused');
    await window.blockSelection.reconnect(1);
  });
  for (const pane of [0, 1])
    await expect(page.locator(`[data-pane="${pane}"] [aria-label="Widget notes"]`)).toHaveCount(0);
  for (const peer of [0, 1]) expect((await documentState(page, tailId, peer)).tail).toBe('Tail');
  const digests = await page.evaluate(() => window.blockSelection.savedDigests());
  expect(digests[0]).toEqual(digests[1]);
  expect(errors).toEqual([]);
});
