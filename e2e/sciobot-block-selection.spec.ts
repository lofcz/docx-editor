import { expect, test, type Page } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/selection';

const ready = async (page: Page) => {
  await page.goto('/selection.html');
  await expect(page.locator('[data-pane="0"] [data-sciobot-native-slot]')).toHaveCount(4);
  await expect(page.locator('[data-pane="0"] [data-container="inner"]')).toHaveCount(1);
};
const lastCopy = (page: Page) =>
  page.evaluate(() => window.blockSelection.copies[window.blockSelection.copies.length - 1]);

test('keyboard extension crosses container and page boundaries and copy returns the document range', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await ready(page);
  const alice = page.locator('[data-pane="0"]');
  const inner = alice.locator('[data-container="inner"] .docx-paragraph-fragment');
  await inner.click({ position: { x: 15, y: 8 } });
  await page.keyboard.press('End');
  await page.keyboard.press('Shift+ArrowUp');
  await page.keyboard.press('Shift+ArrowUp');
  const upward = await page.evaluate(() => window.blockSelection.blocks());
  expect(
    await page.evaluate((head) => window.blockSelection.textAt(head), upward!.selection.head)
  ).toBe('Intro');
  expect(upward!.selection.anchor).toEqual({
    paragraphId: await page.evaluate(() => window.blockSelection.idOf('Inner text')),
    offset: 10,
  });
  expect(upward!.blocks.map(({ id, relation, ordinals }) => [id, relation, ordinals])).toEqual([
    ['outer', 'crosses', [0]],
    ['inner', 'crosses', [0]],
  ]);
  await page.keyboard.press('Control+c');
  const upwardText = await page.evaluate(() => window.blockSelection.rangeText());
  expect(upwardText).toContain('Outer start\nInner text');
  expect(await lastCopy(page)).toBe(upwardText);

  await inner.click({ position: { x: 15, y: 8 } });
  await page.keyboard.press('End');
  for (let index = 0; index < 57; index++) await page.keyboard.press('Shift+ArrowDown');
  const downward = await page.evaluate(() => window.blockSelection.blocks());
  const head = downward!.selection.head,
    anchor = downward!.selection.anchor;
  expect(await page.evaluate((head) => window.blockSelection.textAt(head), head)).toBe('Outer end');
  expect(
    await page.evaluate(
      ([a, h]) => [
        window.blockSelection.pageOf(a!.paragraphId, a!.offset),
        window.blockSelection.pageOf(h!.paragraphId, h!.offset),
      ],
      [anchor, head]
    )
  ).toEqual([0, 2]);
  expect(downward!.blocks.map(({ id, relation, ordinals }) => [id, relation, ordinals])).toEqual([
    ['outer', 'within', [0, 1, 2]],
    ['inner', 'crosses', [0]],
  ]);
  await page.keyboard.press('Control+c');
  const downwardText = await page.evaluate(() => window.blockSelection.rangeText());
  expect(downwardText.startsWith('\nOuter line 1\nOuter line 2\n')).toBe(true);
  expect(downwardText).toMatch(/\nOuter line 56\nOuter[ a-z]*$/);
  expect(await lastCopy(page)).toBe(downwardText);
  await page.keyboard.press('Shift+ArrowDown');
  const beyond = await page.evaluate(() => window.blockSelection.blocks());
  expect(beyond!.blocks.map(({ id, relation }) => [id, relation])).toEqual([
    ['outer', 'crosses'],
    ['inner', 'crosses'],
    ['form', 'crosses'],
  ]);
  await page.keyboard.press('Control+c');
  expect(await lastCopy(page)).toBe(await page.evaluate(() => window.blockSelection.rangeText()));
  expect(errors).toEqual([]);
});

test('widget form selections stay separate from the document selection', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await ready(page);
  const alice = page.locator('[data-pane="0"]');
  const inner = alice.locator('[data-container="inner"] .docx-paragraph-fragment');
  await inner.click({ position: { x: 15, y: 8 } });
  await page.keyboard.press('Home');
  for (let index = 0; index < 60; index++) await page.keyboard.press('Shift+ArrowDown');
  const before = await page.evaluate(() => window.blockSelection.model());
  const revision = await page.evaluate(() => window.blockSelection.revision());
  const notes = alice.getByLabel('Widget notes');
  await notes.click();
  await expect(notes).toBeFocused();
  await page.keyboard.press('Control+a');
  await page.keyboard.press('Shift+ArrowLeft');
  expect(
    await notes.evaluate((node: HTMLTextAreaElement) => [node.selectionStart, node.selectionEnd])
  ).toEqual([0, 'Widget words'.length - 1]);
  await page.keyboard.type('Form ');
  await expect(notes).toHaveValue('Form s');
  const box = (await notes.boundingBox())!;
  await page.mouse.move(box.x + 6, box.y + 9);
  await page.mouse.down();
  await page.mouse.move(box.x + 40, box.y + 9, { steps: 4 });
  await page.mouse.up();
  await expect(notes).toBeFocused();
  expect(await page.evaluate(() => window.blockSelection.model())).toEqual(before);
  expect(await page.evaluate(() => window.blockSelection.revision())).toBe(revision);
  expect(await page.evaluate(() => window.blockSelection.text('Tail'))).toBe('Tail');
  const tail = alice.locator('.docx-paragraph-fragment').filter({ hasText: /^Tail$/ });
  await tail.click({ position: { x: 4, y: 8 } });
  await page.keyboard.press('End');
  await page.keyboard.press('Shift+Home');
  const tailId = await page.evaluate(() => window.blockSelection.idOf('Tail'));
  expect(await page.evaluate(() => window.blockSelection.model())).toEqual({
    anchor: { paragraphId: tailId, offset: 4 },
    head: { paragraphId: tailId, offset: 0 },
  });
  await expect(notes).toHaveValue('Form s');
  expect(errors).toEqual([]);
});

test('two real sessions keep block selections through concurrent edits, history, remote geometry, reconnect, deletion, and saved convergence', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await ready(page);
  await page.evaluate(() =>
    window.blockSelection.select(['Inner text', 2], ['Outer line 50', 4], 1)
  );
  const remote = await page.evaluate(() => window.blockSelection.remoteGeometry());
  expect(remote![0]).toEqual(remote![1]);
  expect(remote![0]!.map((caret) => caret!.page)).toEqual([0, 2]);
  expect(
    (await page.evaluate(() => window.blockSelection.blocks(1)))!.blocks.map(({ id, relation }) => [
      id,
      relation,
    ])
  ).toEqual([
    ['outer', 'within'],
    ['inner', 'crosses'],
  ]);

  await page.evaluate(() => window.blockSelection.pause());
  const alice = page.locator('[data-pane="0"]');
  const inner = alice.locator('[data-container="inner"] .docx-paragraph-fragment');
  await inner.click({ position: { x: 15, y: 8 } });
  await page.keyboard.press('End');
  await page.keyboard.type(' Alice');
  await page.evaluate(async () => {
    await window.blockSelection.append(' Bob', 'Outer line 20', 1);
    await window.blockSelection.select(['Inner text', 2], ['Outer line 50', 4], 1);
    await window.blockSelection.reconnect(1);
  });
  for (const peer of [0, 1])
    expect(
      await page.evaluate(
        (peer) => [
          window.blockSelection.text('Inner text Alice', peer),
          window.blockSelection.text('Outer line 20 Bob', peer),
        ],
        peer
      )
    ).toEqual(['Inner text Alice', 'Outer line 20 Bob']);
  const bob = await page.evaluate(() => window.blockSelection.blocks(1));
  expect(
    await page.evaluate((head) => window.blockSelection.textAt(head, 1), bob!.selection.head)
  ).toBe('Outer line 50');
  expect(bob!.blocks.map(({ id, relation }) => [id, relation])).toEqual([
    ['outer', 'within'],
    ['inner', 'crosses'],
  ]);
  const geometry = await page.evaluate(() => window.blockSelection.remoteGeometry());
  expect(geometry![0]).toEqual(geometry![1]);

  expect(await page.evaluate(() => window.blockSelection.exec('undo'))).toBe(true);
  expect(await page.evaluate(() => window.blockSelection.text('Inner text', 1))).toBe('Inner text');
  expect(await page.evaluate(() => window.blockSelection.text('Outer line 20 Bob', 0))).toBe(
    'Outer line 20 Bob'
  );
  expect(await page.evaluate(() => window.blockSelection.exec('redo'))).toBe(true);
  expect(await page.evaluate(() => window.blockSelection.text('Inner text Alice', 1))).toBe(
    'Inner text Alice'
  );

  await page.evaluate(async () => {
    window.blockSelection.pause();
    await window.blockSelection.append(' editing', 'Inner text Alice', 1);
    await window.blockSelection.select(['Inner text Alice editing', 3], ['Outer line 50', 4], 1);
    if (!(await window.blockSelection.removeBlock('inner', 0))) throw new Error('Removal refused');
    await window.blockSelection.reconnect(0);
  });
  for (const peer of [0, 1]) {
    const result = await page.evaluate((peer) => window.blockSelection.blocks(peer), peer);
    expect(result!.blocks.some(({ id }) => id === 'inner')).toBe(false);
  }
  for (const pane of [0, 1])
    await expect(page.locator(`[data-pane="${pane}"] [data-container="inner"]`)).toHaveCount(0);
  const after = await page.evaluate(() => window.blockSelection.remoteGeometry());
  expect(after![0]).toEqual(after![1]);
  const digests = await page.evaluate(() => window.blockSelection.savedDigests());
  expect(digests[0]).toEqual(digests[1]);
  expect(errors).toEqual([]);
});
