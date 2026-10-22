import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/block-controls';

test('toolbar heading change keeps the selected range and one undo preserves browser text', async ({
  page,
}) => {
  await page.goto('/block-controls.html');
  await page.waitForFunction(() => !!window.blockControls);
  await page.evaluate(() => window.blockControls.select(0, 0, 5));
  const before = await page.evaluate(() => window.blockControls.state());
  await page
    .locator('[data-pane="0"]')
    .getByRole('button', { name: 'Heading 2', exact: true })
    .click();
  await page.evaluate(() => window.blockControls.settle());
  const changed = await page.evaluate(() => window.blockControls.state());
  expect(changed.selection).toEqual(before.selection);
  expect(changed.paragraphs[0]).toEqual({ ...before.paragraphs[0], style: 'Heading2' });
  expect(await page.evaluate(() => window.blockControls.state(1))).toMatchObject({
    paragraphs: changed.paragraphs,
  });
  expect(await page.evaluate(() => window.blockControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.blockControls.state())).paragraphs).toEqual(
    before.paragraphs
  );
  await page.keyboard.type('Hello');
  expect((await page.evaluate(() => window.blockControls.state())).paragraphs[0]!.text).toBe(
    'Hello beta'
  );
  const saved = await page.evaluate(() => window.blockControls.saved());
  expect(saved[0]).toEqual(saved[1]);
});

test('real heading input, empty Enter reset, shortcut, and horizontal-rule input reach native paragraphs', async ({
  page,
}) => {
  await page.goto('/block-controls.html');
  await page.waitForFunction(() => !!window.blockControls);
  await page.evaluate(() => window.blockControls.select(1, 0));
  await page.keyboard.type('## ');
  expect((await page.evaluate(() => window.blockControls.state())).paragraphs[1]).toMatchObject({
    text: '',
    style: 'Heading2',
  });
  await page.keyboard.press('Enter');
  expect((await page.evaluate(() => window.blockControls.state())).paragraphs[1]).toMatchObject({
    text: '',
    style: 'Normal',
  });
  await page.keyboard.press('Control+Alt+1');
  expect((await page.evaluate(() => window.blockControls.state())).paragraphs[1]!.style).toBe(
    'Heading1'
  );
  await page.keyboard.press('Control+Alt+1');
  expect((await page.evaluate(() => window.blockControls.state())).paragraphs[1]!.style).toBe(
    'Normal'
  );
  await page.keyboard.type('--- ');
  const changed = await page.evaluate(() => window.blockControls.state());
  expect(changed.paragraphs[1]).toMatchObject({ text: '', style: 'SciobotHorizontalRule' });
  expect(changed.paragraphs[2]!.text).toBe(' ');
  expect(changed.selection.head).toEqual({ paragraphId: changed.paragraphs[2]!.id, offset: 1 });
  await page.evaluate(() => window.blockControls.settle());
  const saved = await page.evaluate(() => window.blockControls.saved());
  expect(saved[0]).toEqual(saved[1]);
});
