import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/main';

test('native typing and payload edits retain one widget root and arbitrary local React state', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/');
  const widget = page.getByTestId('widget-0');
  await expect(widget.getByRole('button', { name: 'Count 0' })).toBeVisible();
  await widget.getByRole('button', { name: 'Count 0' }).click();
  await expect(widget.getByRole('button', { name: 'Count 1' })).toBeVisible();
  await widget.getByLabel('Notes 0').fill('Unstored local notes');
  const paragraph = page
    .locator('[data-pane="0"] .docx-paragraph-fragment')
    .filter({ hasText: 'Editable body' });
  await paragraph.click({ position: { x: 40, y: 8 } });
  await page.keyboard.press('End');
  await page.keyboard.type(' Typed');
  await expect(paragraph).toContainText('Editable body Typed');
  await expect(widget.getByRole('button', { name: 'Count 1' })).toBeVisible();
  await expect(widget.getByLabel('Notes 0')).toHaveValue('Unstored local notes');
  await expect(widget.locator('[data-native-text]')).toContainText('Editable body Typed');
  expect(await page.evaluate(() => window.blockHostProof.rename('Changed'))).toBe(true);
  await expect(widget.locator('[data-name]')).toHaveText('Changed');
  expect(await page.evaluate(() => window.blockHostProof.events[0])).toEqual({
    mounted: 1,
    unmounted: 0,
  });
  await expect(
    page
      .locator('[data-pane="0"] .docx-paragraph-fragment')
      .filter({ hasText: 'Unregistered fallback' })
  ).toBeVisible();
  await page.evaluate(() => window.blockHostProof.dispose());
  await expect(widget).toHaveCount(0);
  expect(await page.evaluate(() => window.blockHostProof.events[0])).toEqual({
    mounted: 1,
    unmounted: 1,
  });
  await expect(paragraph).toContainText('Editable body Typed');
  expect(errors).toEqual([]);
});

test('a native page move retains the widget root, local count, and focused form value', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/');
  const widget = page.getByTestId('widget-0');
  await widget.getByRole('button', { name: 'Count 0' }).click();
  await expect(widget.getByRole('button', { name: 'Count 1' })).toBeVisible();
  await widget.getByLabel('Notes 0').fill('Across pages');
  const before = await widget.locator('[data-page]').textContent();
  expect(await page.evaluate(() => window.blockHostProof.push())).toBe(true);
  await page.locator('[data-pane="0"]').evaluate((node) => {
    node.scrollTop = node.scrollHeight;
  });
  await expect(widget.locator('[data-page]')).not.toHaveText(before!);
  await expect(widget.getByRole('button', { name: 'Count 1' })).toBeVisible();
  await expect(widget.getByLabel('Notes 0')).toHaveValue('Across pages');
  expect(await page.evaluate(() => window.blockHostProof.events[0])).toEqual({
    mounted: 1,
    unmounted: 0,
  });
  expect(errors).toEqual([]);
});

test('two sessions synchronize payloads and native text through reconnect, history, deletion, and saved reopen', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/');
  const alice = page.getByTestId('widget-0'),
    bob = page.getByTestId('widget-1');
  await alice.getByRole('button', { name: 'Count 0' }).click();
  await expect(alice.getByRole('button', { name: 'Count 1' })).toBeVisible();
  await page.evaluate(async () => {
    window.blockHostProof.pause();
    if (!window.blockHostProof.rename('Offline')) throw new Error('Rename refused');
    await window.blockHostProof.append(' Bob');
    await window.blockHostProof.reconnect(1);
  });
  await expect(alice.locator('[data-name]')).toHaveText('Offline');
  await expect(bob.locator('[data-name]')).toHaveText('Offline');
  await expect(alice.locator('[data-native-text]')).toContainText('Editable body Bob');
  expect(await page.evaluate(() => window.blockHostProof.undo())).toBe(true);
  await expect(alice.locator('[data-name]')).toHaveText('Initial');
  await expect(bob.locator('[data-name]')).toHaveText('Initial');
  await expect(bob.locator('[data-native-text]')).toContainText('Editable body Bob');
  expect(await page.evaluate(() => window.blockHostProof.redo())).toBe(true);
  await expect(alice.getByRole('button', { name: 'Count 1' })).toBeVisible();
  const digests = await page.evaluate(() => window.blockHostProof.savedDigests());
  expect(digests[0]).toEqual(digests[1]);
  await page.evaluate(() => window.blockHostProof.reopen());
  await page.locator('[data-reopened]').scrollIntoViewIfNeeded();
  await expect(page.getByTestId('widget-2').locator('[data-name]')).toHaveText('Offline');
  await expect(page.getByTestId('widget-2').locator('[data-native-text]')).toContainText(
    'Editable body Bob'
  );
  expect(await page.evaluate(() => window.blockHostProof.remove())).toBe(true);
  await expect(alice).toHaveCount(0);
  await expect(bob).toHaveCount(0);
  expect(await page.evaluate(() => window.blockHostProof.events.slice(0, 2))).toEqual([
    { mounted: 1, unmounted: 1 },
    { mounted: 1, unmounted: 1 },
  ]);
  expect(errors).toEqual([]);
});

test('locale and application services update per editor while React form state and native selection survive', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/');
  const alice = page.getByTestId('widget-0'),
    bob = page.getByTestId('widget-1');
  await expect(alice.locator('[data-provider]')).toHaveText('en|Alice:Initial');
  await expect(bob.locator('[data-provider]')).toHaveText('cs|Bob:Initial');
  await alice.getByRole('button', { name: 'Count 0' }).click();
  await alice.getByLabel('Notes 0').fill('Local notes');
  await page.evaluate(() => window.blockHostProof.providers('fr', 'Alice updated'));
  await expect(alice.locator('[data-provider]')).toHaveText('fr|Alice updated:Initial');
  await expect(bob.locator('[data-provider]')).toHaveText('cs|Bob:Initial');
  await expect(alice.getByRole('button', { name: 'Count 1' })).toBeVisible();
  await expect(alice.getByLabel('Notes 0')).toBeFocused();
  await expect(alice.getByLabel('Notes 0')).toHaveValue('Local notes');
  const paragraph = page
    .locator('[data-pane="0"] .docx-paragraph-fragment')
    .filter({ hasText: 'Editable body' });
  await paragraph.click({ position: { x: 25, y: 8 } });
  await page.keyboard.press('End');
  await page.keyboard.type(' Provider');
  await expect(paragraph).toHaveText('Editable body Provider');
  const selection = await page.evaluate(() => window.blockHostProof.selection());
  expect(selection.model!.head.offset).toBe(22);
  await page.evaluate(() => window.blockHostProof.providers('de', 'Alice live'));
  await expect(alice.locator('[data-provider]')).toHaveText('de|Alice live:Initial');
  expect(await page.evaluate(() => window.blockHostProof.selection())).toEqual(selection);
  await page.keyboard.type('!');
  await expect(paragraph).toHaveText('Editable body Provider!');
  await page.evaluate(async () => {
    if (!window.blockHostProof.rename('Shared')) throw new Error('Rename refused');
    await window.blockHostProof.settle();
  });
  await expect(alice.locator('[data-provider]')).toHaveText('de|Alice live:Shared');
  await expect(bob.locator('[data-provider]')).toHaveText('cs|Bob:Shared');
  expect(await page.evaluate(() => window.blockHostProof.events.slice(0, 2))).toEqual([
    { mounted: 1, unmounted: 0 },
    { mounted: 1, unmounted: 0 },
  ]);
  const saved = await page.evaluate(() => window.blockHostProof.savedDigests());
  expect(saved[0]).toEqual(saved[1]);
  expect(errors).toEqual([]);
});
