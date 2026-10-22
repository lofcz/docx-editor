import { expect, test, type Page } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/restore';

async function ready(page: Page) {
  await page.waitForFunction(() => window.reviewRestore?.ready === true);
  await expect(page.locator('[data-review-panel]')).toHaveCount(2);
}
async function recovered(page: Page, peer = 0) {
  const result = await page.evaluate((index) => window.reviewRestore.state(index), peer);
  if (!result.success) throw new Error(result.error);
  return result.state;
}
async function typeTail(page: Page, text: string) {
  const tail = page
    .locator('[data-pane="0"] .docx-paragraph-fragment')
    .filter({ hasText: 'Keep tail' });
  await tail.click({ position: { x: 20, y: 8 } });
  await page.keyboard.press('End');
  await page.keyboard.type(text);
  const selection = await page.evaluate(() => window.reviewRestore.selection());
  expect(selection.model!.head.paragraphId).toBe(selection.domParagraphId);
  await expect(tail).toContainText(text);
}
async function save(page: Page) {
  await page.getByRole('button', { name: 'Save DOCX', exact: true }).click();
  await expect(page.locator('[data-saved]')).toHaveText('Saved');
}

for (const storage of ['docx', 'room'] as const)
  test(
    'hard reload restores nested review without chat state, native typing, decisions and fresh history; ' +
      storage,
    async ({ page }) => {
      const errors: string[] = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto('/restore.html?storage=' + storage);
      await ready(page);
      await expect(page.locator('[data-review-controls]')).toHaveCount(0);
      await page.getByRole('button', { name: 'Create review', exact: true }).click();
      const panel = page.locator('[data-review-panel="0"]');
      await expect(panel.locator('[data-group-count]')).toHaveText('6');
      await expect(panel.locator('[data-property-count]')).toHaveText('2');
      await expect(page.locator('[data-pane="0"] [data-block="activity"]')).toHaveCount(1);
      await typeTail(page, ' Manual.');
      await save(page);
      const before = await recovered(page),
        beforeWord = await page.evaluate(() => window.reviewRestore.word()),
        epoch = await page.evaluate(() => window.reviewRestore.epoch);
      expect(await page.evaluate(() => window.reviewRestore.storageKeys())).toEqual(
        storage === 'room' ? ['document', 'snapshot', 'version'] : ['document', 'version']
      );
      await page.reload();
      await ready(page);
      const after = await recovered(page);
      expect(after.groups.map((group) => group.id)).toEqual(before.groups.map((group) => group.id));
      expect(after.pendingProperties.map((proposal) => proposal.id)).toEqual(
        before.pendingProperties.map((proposal) => proposal.id)
      );
      expect(after.reviewableGroupIds).toEqual(before.reviewableGroupIds);
      expect(after.revision.epoch).not.toBe(epoch);
      expect(await page.evaluate(() => window.reviewRestore.word())).toBe(beforeWord);
      expect(await page.evaluate(() => window.reviewRestore.canUndo())).toBe(false);
      await expect(panel.locator('[data-group-count]')).toHaveText('6');
      await expect(panel.locator('[data-property-count]')).toHaveText('2');
      await typeTail(page, ' Reloaded.');
      const activity = after.groups.findIndex((group) =>
        group.properties.some((member) => member.blockId === 'activity')
      );
      await panel
        .locator('[data-review-group]')
        .nth(activity)
        .getByRole('button', { name: 'Accept', exact: true })
        .click();
      await expect(panel.locator('[data-group-count]')).toHaveText('5');
      await expect(panel.locator('[data-property-count]')).toHaveText('1');
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await expect(panel.locator('[data-group-count]')).toHaveText('6');
      expect((await recovered(page)).groups.map((group) => group.id)).toEqual(
        before.groups.map((group) => group.id)
      );
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await expect(panel.locator('[data-group-count]')).toHaveText('5');
      await panel.getByRole('button', { name: 'Reject all', exact: true }).click();
      await expect(page.locator('[data-review-controls]')).toHaveCount(0);
      await expect(panel.locator('[data-property-count]')).toHaveText('0');
      expect(await page.evaluate(() => window.reviewRestore.texts())).toContain(
        'Keep tail Manual. Reloaded.'
      );
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await expect(panel.locator('[data-group-count]')).toHaveText('5');
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await expect(page.locator('[data-review-controls]')).toHaveCount(0);
      await save(page);
      await page.reload();
      await ready(page);
      await expect(page.locator('[data-review-group]')).toHaveCount(0);
      await expect(page.locator('[data-review-controls]')).toHaveCount(0);
      expect((await recovered(page)).pendingProperties).toEqual([]);
      expect(await page.evaluate(() => window.reviewRestore.texts())).toContain(
        'Keep tail Manual. Reloaded.'
      );
      expect(errors).toEqual([]);
    }
  );

for (const reverse of [false, true])
  for (const receiver of [0, 1] as const)
    test(
      'restored two-session review retains mixed decisions, offline input, history, geometry, and deletion; reverse=' +
        reverse +
        '/' +
        receiver,
      async ({ page }) => {
        const errors: string[] = [];
        page.on('pageerror', (error) => errors.push(error.message));
        await page.goto('/restore.html?storage=room&reverse=' + Number(reverse));
        await ready(page);
        await page.getByRole('button', { name: 'Create review', exact: true }).click();
        await expect(page.locator('[data-review-panel="0"] [data-group-count]')).toHaveText('6');
        const initial = await recovered(page),
          visual = initial.groups.find((group) =>
            group.properties.some((member) => member.blockId === 'visual')
          )!.id;
        await page.evaluate(
          async ({ first, visual }) => {
            await window.reviewRestore.decide('accept', first);
            await window.reviewRestore.decide('reject', visual);
          },
          { first: initial.groups[0]!.id, visual }
        );
        const pending = await recovered(page),
          ids = pending.groups.map((group) => group.id);
        expect(ids).toHaveLength(4);
        await save(page);
        await page.reload();
        await ready(page);
        for (const peer of [0, 1]) {
          expect((await recovered(page, peer)).groups.map((group) => group.id)).toEqual(ids);
          expect((await recovered(page, peer)).pendingProperties).toHaveLength(2);
        }
        await page.evaluate(async () => {
          window.reviewRestore.pause();
          await window.reviewRestore.peerAppend(' Peer note.');
        });
        await typeTail(page, ' Local note.');
        const action = receiver === 0 ? 'accept' : 'reject';
        await page
          .locator('[data-review-panel="0"]')
          .getByRole('button', {
            name: action === 'accept' ? 'Accept all' : 'Reject all',
            exact: true,
          })
          .click();
        await page.evaluate((first) => window.reviewRestore.reconnect(first), receiver);
        await expect(page.locator('[data-review-controls]')).toHaveCount(0);
        const texts = await page.evaluate(() => [
          window.reviewRestore.texts(),
          window.reviewRestore.texts(1),
        ]);
        expect(texts[0]).toEqual(texts[1]);
        expect(texts[0]).toContain('The car sat.');
        expect(texts[0]).toContain(
          action === 'accept' ? 'New activity — 15 minutes' : 'Old activity — 10 minutes'
        );
        expect(texts[0]).toContain(
          action === 'accept' ? 'New audio fallback' : 'Old audio fallback'
        );
        const tail = texts[0]!.find((text) => text?.includes('Keep tail'))!;
        expect(tail).toContain('Local note.');
        expect(tail).toContain('Peer note.');
        const geometry = await page.evaluate(() => window.reviewRestore.remoteGeometry());
        expect(geometry?.[0]).toBeTruthy();
        expect(geometry?.[0]).toEqual(geometry?.[1]);
        expect(await page.evaluate(() => window.reviewRestore.undo())).toBe(true);
        for (const peer of [0, 1])
          expect((await recovered(page, peer)).groups.map((group) => group.id)).toEqual(ids);
        expect(await page.evaluate(() => window.reviewRestore.undo(1))).toBe(true);
        expect(
          (await page.evaluate(() => window.reviewRestore.texts())).find((text) =>
            text?.includes('Keep tail')
          )
        ).toBe('Keep tail Local note.');
        expect(await page.evaluate(() => window.reviewRestore.redo(1))).toBe(true);
        expect(await page.evaluate(() => window.reviewRestore.redo())).toBe(true);
        await expect(page.locator('[data-review-controls]')).toHaveCount(0);
        const digests = await page.evaluate(() => window.reviewRestore.savedDigests());
        expect(digests[0]).toEqual(digests[1]);
        expect(await page.evaluate(() => window.reviewRestore.removeActivity())).toBe(true);
        await expect(page.locator('[data-block="activity"]')).toHaveCount(0);
        await expect(page.locator('[data-review-controls]')).toHaveCount(0);
        await save(page);
        await page.reload();
        await ready(page);
        await expect(page.locator('[data-block="activity"]')).toHaveCount(0);
        await expect(page.locator('[data-review-group]')).toHaveCount(0);
        await expect(page.locator('[data-review-controls]')).toHaveCount(0);
        expect((await recovered(page)).pendingProperties).toEqual([]);
        expect(errors).toEqual([]);
      }
    );
