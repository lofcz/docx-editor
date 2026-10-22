import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/native-slots';

test('real typing inside nested React containers and table cells retains native caret and one painted paragraph', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/native-slots.html');
  const alice = page.locator('[data-pane="0"]');
  await expect(alice.locator('[data-sciobot-native-slot]')).toHaveCount(7);
  await alice.getByRole('button', { name: 'Count 0' }).click();
  for (const [block, text] of [
    ['inner', 'Inner text'],
    ['inside-cell', 'Cell text'],
    ['cell', 'Row cell text'],
  ]) {
    const paragraph = alice.locator(`[data-container="${block}"] .docx-paragraph-fragment`);
    await expect(paragraph).toHaveText(text!);
    await paragraph.click({ position: { x: 15, y: 8 } });
    await page.keyboard.press('End');
    await page.keyboard.type(' Typed');
    await expect(paragraph).toHaveText(text + ' Typed');
    const before = await page.evaluate(() => window.nativeSlotProof.selection());
    expect(before.model!.head.offset).toBe(text!.length + 6);
    expect(before.dom.paragraphId).toBe(before.model!.head.paragraphId);
    await page.evaluate(() => window.nativeSlotProof.refresh());
    expect(await page.evaluate(() => window.nativeSlotProof.selection())).toEqual(before);
    await page.keyboard.type('!');
    await expect(paragraph).toHaveText(text + ' Typed!');
    await expect(alice.locator('.docx-paragraph-fragment')).toHaveCount(12);
    await expect(alice.getByRole('button', { name: 'Count 1' })).toBeVisible();
  }
  await page.evaluate(() => window.nativeSlotProof.unregister());
  await expect(alice.locator('[data-sciobot-native-slot]')).toHaveCount(0);
  await expect(alice.locator('.docx-paragraph-fragment')).toHaveCount(12);
  await expect(alice.locator('.docx-page-content > .docx-table-fragment')).toHaveCount(2);
  await page.evaluate(() => window.nativeSlotProof.register());
  await expect(alice.locator('[data-sciobot-native-slot]')).toHaveCount(7);
  await page.evaluate(() => window.nativeSlotProof.dispose());
  await expect(alice.locator('[data-sciobot-native-slot]')).toHaveCount(0);
  await expect(alice.locator('.docx-page-content > .docx-paragraph-fragment')).toHaveCount(7);
  expect(errors).toEqual([]);
});

test('two real sessions keep native nested slots through concurrent typing, history, remote selection, reconnect, deletion, and saved reopen', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/native-slots.html');
  const alice = page.locator('[data-pane="0"]'),
    bob = page.locator('[data-pane="1"]');
  await expect(alice.locator('[data-sciobot-native-slot]')).toHaveCount(7);
  await page.evaluate(() => window.nativeSlotProof.pause());
  const inner = alice.locator('[data-container="inner"] .docx-paragraph-fragment');
  await inner.click({ position: { x: 15, y: 8 } });
  await page.keyboard.press('End');
  await page.keyboard.type(' Alice');
  await page.evaluate(async () => {
    await window.nativeSlotProof.append(' Bob');
    await window.nativeSlotProof.reconnect(1);
  });
  for (const pane of [alice, bob]) {
    await expect(pane.locator('[data-container="inner"] .docx-paragraph-fragment')).toHaveText(
      'Inner text Alice'
    );
    await expect(
      pane.locator('.docx-paragraph-fragment').filter({ hasText: 'Other cell Bob' })
    ).toHaveCount(1);
    await expect(pane.locator('[data-sciobot-native-slot]')).toHaveCount(7);
  }
  const remote = await page.evaluate(() => window.nativeSlotProof.remoteGeometry());
  expect(remote?.[0]).toBeTruthy();
  expect(remote?.[0]).toEqual(remote?.[1]);
  expect(await page.evaluate(() => window.nativeSlotProof.undo())).toBe(true);
  await expect(inner).toHaveText('Inner text');
  await expect(
    alice.locator('.docx-paragraph-fragment').filter({ hasText: 'Other cell Bob' })
  ).toHaveCount(1);
  expect(await page.evaluate(() => window.nativeSlotProof.redo())).toBe(true);
  await expect(inner).toHaveText('Inner text Alice');
  const digests = await page.evaluate(() => window.nativeSlotProof.savedDigests());
  expect(digests[0]).toEqual(digests[1]);
  await page.evaluate(() => window.nativeSlotProof.reopen());
  await page.locator('[data-reopened]').scrollIntoViewIfNeeded();
  await expect(
    page.locator('[data-reopened] [data-container="inner"] .docx-paragraph-fragment')
  ).toHaveText('Inner text Alice');
  await expect(page.locator('[data-reopened] [data-sciobot-native-slot]')).toHaveCount(7);
  await page.evaluate(async () => {
    window.nativeSlotProof.pause();
    await window.nativeSlotProof.append(' Editing', 1, 'Inner text');
    if (!(await window.nativeSlotProof.remove())) throw new Error('Removal refused');
    await window.nativeSlotProof.reconnect(0);
  });
  for (const pane of [alice, bob]) {
    await expect(pane.locator('[data-container="outer"]')).toHaveCount(0);
    await expect(pane.locator('[data-sciobot-native-slot]')).toHaveCount(2);
    await expect(pane.locator('.docx-paragraph-fragment')).toHaveCount(4);
  }
  const deleted = await page.evaluate(() => window.nativeSlotProof.savedDigests());
  expect(deleted[0]).toEqual(deleted[1]);
  expect(errors).toEqual([]);
});
