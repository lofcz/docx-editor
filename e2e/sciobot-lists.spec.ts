import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/lists';
for (const [style, glyph] of [
  ['circle', '○'],
  ['upper-roman', 'I.'],
] as const)
  test(`real ${style} controls preserve imported numbering, native selection, nesting, history and Word output`, async ({
    page,
  }) => {
    await page.goto('/lists.html');
    await page.waitForFunction(() => !!window.listControls);
    await page.evaluate(() => window.listControls.select(0, 2));
    const before = await page.evaluate(() => window.listControls.state());
    const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Lists' });
    await toolbar.getByRole('combobox', { name: 'List style' }).selectOption(style);
    await toolbar.getByRole('button', { name: 'Apply list', exact: true }).click();
    await page.evaluate(() => window.listControls.settle());
    const applied = await page.evaluate(() => window.listControls.state());
    expect(applied.paragraphs[0]!.marker!.text).toBe(glyph);
    expect(applied.selection).toEqual(before.selection);
    expect(applied.paragraphs[2]!.marker).toEqual(before.paragraphs[2]!.marker);
    await toolbar.getByRole('button', { name: 'Indent list', exact: true }).click();
    await page.evaluate(() => window.listControls.settle());
    const nested = await page.evaluate(() => window.listControls.state());
    expect(nested.paragraphs[0]!.marker).toMatchObject({ text: glyph, level: 1 });
    expect(await page.evaluate(() => window.listControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
      applied.fingerprint
    );
    expect(await page.evaluate(() => window.listControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
      before.fingerprint
    );
    expect(await page.evaluate(() => window.listControls.redo())).toBe(true);
    expect(await page.evaluate(() => window.listControls.redo())).toBe(true);
    await page.keyboard.type('Live');
    await page.evaluate(() => window.listControls.settle());
    const typed = await page.evaluate(() => window.listControls.state());
    expect(typed.paragraphs[0]!.text).toBe('AlLivepha');
    const saved = await page.evaluate(() => window.listControls.saved());
    expect(saved[0]).toEqual(saved[1]);
    expect(saved[0]!.fingerprint).toEqual(typed.fingerprint);
  });
test('real checklist prefix, typing, nested Enter, empty-item outdent and Backspace retain exact native history', async ({
  page,
}) => {
  await page.goto('/lists.html');
  await page.waitForFunction(() => !!window.listControls);
  await page.evaluate(() => window.listControls.select(0, 0, 5));
  await page.keyboard.type('[x] ');
  let state = await page.evaluate(() => window.listControls.state());
  expect(state.paragraphs[0]!.membership!.checkbox!.checked).toBe(true);
  expect(state.selection.head.offset).toBe(1);
  await page.keyboard.type('Task');
  await page.keyboard.press('Tab');
  state = await page.evaluate(() => window.listControls.state());
  expect(state.paragraphs[0]!.membership!.checkbox!.level).toBe(1);
  const before = state.fingerprint;
  await page.keyboard.press('Enter');
  await page.evaluate(() => window.listControls.settle());
  const after = await page.evaluate(() => window.listControls.state());
  expect(after.paragraphs[1]!.membership!.checkbox).toMatchObject({ checked: false, level: 1 });
  expect(after.selection.head.offset).toBe(1);
  expect(await page.evaluate(() => window.listControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(before);
  expect(await page.evaluate(() => window.listControls.redo())).toBe(true);
  expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
    after.fingerprint
  );
  expect((await page.evaluate(() => window.listControls.state())).selection).toEqual(
    after.selection
  );
  await page.keyboard.press('Enter');
  state = await page.evaluate(() => window.listControls.state());
  expect(state.paragraphs[1]!.membership!.checkbox!.level).toBe(0);
  await page.keyboard.press('Backspace');
  await page.evaluate(() => window.listControls.settle());
  state = await page.evaluate(() => window.listControls.state());
  expect(state.paragraphs[1]!.membership!.checkbox).toBeNull();
  expect(state.paragraphs[1]!.text).toBe('');
  const saved = await page.evaluate(() => window.listControls.saved());
  expect(saved[0]).toEqual(saved[1]);
  expect(saved[0]!.fingerprint).toEqual(state.fingerprint);
});
