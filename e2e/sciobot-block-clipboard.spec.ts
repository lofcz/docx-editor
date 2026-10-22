import { expect, test, type Page } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/clipboard';

const ready = async (page: Page) => {
  await page.goto('/clipboard.html');
  await expect(page.locator('[data-pane="0"] [data-block="note-1"]')).toHaveCount(1);
  await expect(page.locator('[data-pane="1"] [data-block="note-1"]')).toHaveCount(1);
};
const blocks = (page: Page, peer = 0) =>
  page.evaluate((peer) => window.blockClipboard.blocks(peer), peer);
const paragraph = (page: Page, pane: number, text: string) =>
  page
    .locator(`[data-pane="${pane}"] .docx-paragraph-fragment`)
    .filter({ hasText: new RegExp(`^${text}$`) })
    .first();

/** Real keyboard selection from the end of `Phase one` to the start of `Phase end`. */
async function selectActivity(page: Page, pane = 0) {
  await paragraph(page, pane, 'Phase one').click({ position: { x: 10, y: 8 } });
  await page.keyboard.press('End');
  for (let index = 0; index < 3; index++) await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+Home');
}
async function pasteAtTail(page: Page, pane = 0) {
  const tail = page.locator(`[data-pane="${pane}"] .docx-paragraph-fragment`).filter({
    hasText: /^Tail/,
  });
  await tail.last().click({ position: { x: 4, y: 8 } });
  await page.keyboard.press('Control+End');
  await page.keyboard.press('Control+v');
}
const pastedActivities = (result: Awaited<ReturnType<typeof blocks>>) =>
  result.blocks.filter((block) => block.type === 'activity' && block.id !== 'activity-1');

test('copying a nested activity carries its payloads, and two pastes create distinct remapped IDs', async ({
  page,
  context,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await ready(page);
  await selectActivity(page);
  await page.keyboard.press('Control+c');
  const transfer = await page.evaluate(() => window.blockClipboard.transfers.at(-1)!);
  expect(transfer['text/plain']).toBe('\nActivity text\nNote text\n');
  expect(transfer['text/html']).toContain('data-docx-fragment');
  const data = JSON.parse(transfer['application/x-sciobot-blocks+json']!);
  expect(data.blocks.map((entry: { id: string }) => entry.id)).toEqual([
    'phase-1',
    'activity-1',
    'note-1',
  ]);

  await pasteAtTail(page);
  const first = await blocks(page);
  expect(first.diagnostics).toEqual([]);
  const [pasted] = pastedActivities(first);
  expect(pasted).toMatchObject({ parentId: null, attributes: { name: 'Warmup', duration: '10' } });
  const pastedNote = first.blocks.find((block) => block.parentId === pasted!.id)!;
  expect(pastedNote).toMatchObject({
    type: 'note',
    attributes: { activityRef: pasted!.id, tone: 'info' },
  });
  expect(pastedNote.id).not.toBe('note-1');
  // The source keeps its own identities and references.
  expect(first.blocks.slice(0, 4).map((block) => [block.id, block.parentId])).toEqual([
    ['phase-1', null],
    ['activity-1', 'phase-1'],
    ['note-1', 'activity-1'],
    ['phase-2', null],
  ]);
  await expect(page.locator(`[data-pane="0"] [data-block="${pasted!.id}"]`)).toHaveCount(1);
  await expect(page.locator(`[data-pane="0"] [data-block="${pastedNote.id}"]`)).toHaveCount(1);

  // A later paste is its own undo unit; one undo removes its content and payload records.
  await page.waitForTimeout(1200);
  await pasteAtTail(page);
  const second = await blocks(page);
  expect(second.diagnostics).toEqual([]);
  const ids = second.blocks.map((block) => block.id);
  expect(new Set(ids).size).toBe(ids.length);
  expect(pastedActivities(second)).toHaveLength(2);
  expect(
    await page.evaluate(() =>
      window.blockClipboard.texts().filter((text) => text === 'Activity text')
    )
  ).toHaveLength(3);
  await page.keyboard.press('Control+z');
  expect((await blocks(page)).blocks).toEqual(first.blocks);
  await page.keyboard.press('Control+Shift+z');
  expect((await blocks(page)).blocks).toEqual(second.blocks);
  const reopened = await page.evaluate(() => window.blockClipboard.reopenedBlocks());
  expect(reopened.diagnostics).toEqual([]);
  expect(reopened.ids.sort()).toEqual([...ids].sort());

  // External plain text stays a native paste with no block repair.
  await page.evaluate(() => navigator.clipboard.writeText('Plain words'));
  const before = await page.evaluate(() => window.blockClipboard.revision());
  await pasteAtTail(page);
  const plain = await blocks(page);
  expect(plain.blocks).toEqual(second.blocks);
  expect(await page.evaluate(() => window.blockClipboard.texts().at(-1))).toBe('Plain words');
  expect(await page.evaluate(() => window.blockClipboard.revision())).toBe(before + 1);

  // Cut carries the payloads before the native delete removes the source.
  await selectActivity(page);
  await page.keyboard.press('Control+x');
  const cut = await page.evaluate(() => window.blockClipboard.transfers.at(-1)!);
  expect(
    JSON.parse(cut['application/x-sciobot-blocks+json']!).blocks.map(
      (entry: { id: string }) => entry.id
    )
  ).toEqual(['phase-1', 'activity-1', 'note-1']);
  expect((await blocks(page)).blocks.some((block) => block.id === 'activity-1')).toBe(false);
  await pasteAtTail(page);
  const moved = await blocks(page);
  expect(moved.diagnostics.filter((entry) => !entry.startsWith('orphan-payload:'))).toEqual([]);
  expect(pastedActivities(moved)).toHaveLength(3);
  expect(errors).toEqual([]);
});

test('two real sessions converge on pasted blocks through concurrent pastes, peer edits, history, reconnect, and deletion', async ({
  page,
  context,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await ready(page);
  await selectActivity(page);
  await page.keyboard.press('Control+c');
  await page.evaluate(() => window.blockClipboard.pause());
  await pasteAtTail(page, 0);
  await pasteAtTail(page, 1);
  expect(await page.evaluate(() => window.blockClipboard.insert(' Bob', 'Activity text', 1))).toBe(
    true
  );
  await page.evaluate(() => window.blockClipboard.reconnect(1));
  const [alice, bob] = [await blocks(page, 0), await blocks(page, 1)];
  expect(alice).toEqual(bob);
  expect(alice.diagnostics).toEqual([]);
  expect(pastedActivities(alice)).toHaveLength(2);
  const ids = alice.blocks.map((block) => block.id);
  expect(new Set(ids).size).toBe(ids.length);
  for (const activity of pastedActivities(alice))
    expect(
      alice.blocks.find((block) => block.parentId === activity.id)?.attributes.activityRef
    ).toBe(activity.id);
  for (const pane of [0, 1])
    await expect(page.locator(`[data-pane="${pane}"] [data-type="activity"]`)).toHaveCount(3);

  // Alice's paste undo keeps Bob's paste and Bob's concurrent text.
  expect(await page.evaluate(() => window.blockClipboard.exec('undo', 0))).toBe(true);
  for (const peer of [0, 1]) {
    const result = await blocks(page, peer);
    expect(result.diagnostics).toEqual([]);
    expect(pastedActivities(result)).toHaveLength(1);
  }
  expect(await page.evaluate(() => window.blockClipboard.texts(0))).toContain('Activity text Bob');
  expect(await page.evaluate(() => window.blockClipboard.exec('redo', 0))).toBe(true);
  expect(await blocks(page, 0)).toEqual(alice);
  expect(await blocks(page, 1)).toEqual(alice);

  // Bob deletes the source activity while Alice pastes another copy offline.
  await page.evaluate(() => window.blockClipboard.pause());
  await pasteAtTail(page, 0);
  expect(await page.evaluate(() => window.blockClipboard.removeBlock('activity-1', 1))).toBe(true);
  await page.evaluate(() => window.blockClipboard.reconnect(0));
  const final = await blocks(page, 0);
  expect(await blocks(page, 1)).toEqual(final);
  // Block deletion owns payload removal; the index reports the deleted source as orphaned.
  expect(final.diagnostics.filter((entry) => !entry.startsWith('orphan-payload:'))).toEqual([]);
  expect(final.blocks.some((block) => block.id === 'activity-1')).toBe(false);
  expect(pastedActivities(final)).toHaveLength(3);
  const digests = await page.evaluate(() => window.blockClipboard.savedDigests());
  expect(digests[0]).toEqual(digests[1]);
  expect(errors).toEqual([]);
});
