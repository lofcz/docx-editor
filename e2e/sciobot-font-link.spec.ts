import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/font-link';
test('font controls load packaged faces and paint exact odd-pixel sizes with native typing and Word history', async ({
  page,
}) => {
  const fetched: string[] = [];
  page.on('requestfinished', (r) => {
    if (/\.ttf(?:\?|$)/i.test(r.url())) fetched.push(r.url());
  });
  await page.goto('/font-link.html');
  await page.waitForFunction(() => !!window.fontLinkControls);
  await expect
    .poll(() => page.evaluate(() => window.fontLinkControls.state().measurement))
    .toMatchObject({ measurer: 'shaped', resolving: false });
  await page.evaluate(() => window.fontLinkControls.select(0, 5));
  const before = await page.evaluate(() => window.fontLinkControls.state());
  const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Fonts' });
  await toolbar.getByRole('combobox', { name: 'Font family' }).selectOption('Arial');
  await expect
    .poll(() => page.evaluate(() => window.fontLinkControls.state().measurement))
    .toMatchObject({ measurer: 'shaped', resolving: false });
  await expect.poll(() => fetched.some((f) => /LiberationSans/i.test(f))).toBe(true);
  const family = await page.evaluate(() => window.fontLinkControls.state());
  expect(family.font.family).toBe('Arial');
  await toolbar.getByRole('spinbutton', { name: 'Font size (px)' }).fill('9');
  await toolbar.getByRole('button', { name: 'Apply font size', exact: true }).click();
  await page.evaluate(() => window.fontLinkControls.settle());
  const sized = await page.evaluate(() => window.fontLinkControls.state());
  expect(sized.font.sizePixels).toBe(9);
  expect(sized.spans.find((s) => s.start === 0)!.points).toBe(6.75);
  expect(sized.spans.find((s) => s.start === 0)!.family).toBe('Arial');
  expect(await page.evaluate(() => window.fontLinkControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.fontLinkControls.state())).fingerprint).toEqual(
    family.fingerprint
  );
  expect(await page.evaluate(() => window.fontLinkControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.fontLinkControls.state())).fingerprint).toEqual(
    before.fingerprint
  );
  expect(await page.evaluate(() => window.fontLinkControls.redo())).toBe(true);
  expect(await page.evaluate(() => window.fontLinkControls.redo())).toBe(true);
  await page.evaluate(() => window.fontLinkControls.select(2));
  await page.keyboard.type('X');
  const typed = await page.evaluate(() => window.fontLinkControls.state());
  expect(typed.text).toBe('AlXpha beta');
  expect(typed.font.sizePixels).toBe(9);
  await page.evaluate(() => window.fontLinkControls.settle());
  const saved = await page.evaluate(() => window.fontLinkControls.saved());
  expect(saved[0]).toEqual(saved[1]);
  expect(saved[0]).toEqual(typed.fingerprint);
});
test('mixed values, family clear and size clear keep inherited fonts and the native selection', async ({
  page,
}) => {
  await page.goto('/font-link.html');
  await page.waitForFunction(() => !!window.fontLinkControls);
  await expect
    .poll(() => page.evaluate(() => window.fontLinkControls.state().measurement.resolving))
    .toBe(false);
  const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Fonts' });
  await page.evaluate(() => window.fontLinkControls.select(0, 10));
  expect((await page.evaluate(() => window.fontLinkControls.state())).font.family).toBeNull();
  expect(await toolbar.getByRole('combobox', { name: 'Font family' }).inputValue()).toBe('');
  await page.evaluate(() => window.fontLinkControls.select(0, 5));
  await toolbar.getByRole('spinbutton', { name: 'Font size (px)' }).fill('17');
  await toolbar.getByRole('button', { name: 'Apply font size', exact: true }).click();
  await page.evaluate(() => window.fontLinkControls.select(0, 10));
  expect((await page.evaluate(() => window.fontLinkControls.state())).font.sizePixels).toBeNull();
  expect(await toolbar.getByRole('spinbutton', { name: 'Font size (px)' }).inputValue()).toBe('');
  await page.evaluate(() => window.fontLinkControls.select(0, 5));
  await toolbar.getByRole('button', { name: 'Clear font size', exact: true }).click();
  await toolbar.getByRole('button', { name: 'Clear font family', exact: true }).click();
  expect((await page.evaluate(() => window.fontLinkControls.state())).font).toEqual({
    family: 'Calibri',
    sizePixels: 22 / 1.5,
  });
});
test('link dialog retains native selection, edits and undoes targets without changing labels, rejects unsafe URLs and activates safely', async ({
  page,
  context,
}) => {
  await context.route('https://example.com/**', (route) =>
    route.fulfill({ contentType: 'text/html', body: '<p>Link target</p>' })
  );
  await page.goto('/font-link.html');
  await page.waitForFunction(() => !!window.fontLinkControls);
  await page.evaluate(() => window.fontLinkControls.select(0, 5));
  const before = await page.evaluate(() => window.fontLinkControls.state());
  const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Links' });
  await toolbar.getByRole('button', { name: 'Edit link', exact: true }).click();
  const dialog = toolbar.getByRole('dialog', { name: 'Edit link' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('textbox', { name: 'Link target' })).toHaveValue('');
  expect((await page.evaluate(() => window.fontLinkControls.state())).selection).toEqual(
    before.selection
  );
  await dialog.getByRole('textbox', { name: 'Link target' }).fill('https://example.com/old');
  await dialog.getByRole('button', { name: 'Apply link', exact: true }).click();
  await page.evaluate(() => window.fontLinkControls.settle());
  const inserted = await page.evaluate(() => window.fontLinkControls.state());
  expect(inserted.links[0]!.text).toBe('Alpha');
  await page.evaluate(() => window.fontLinkControls.select(2));
  await page.keyboard.press('Control+k');
  await expect(dialog).toBeVisible();
  await dialog.getByRole('textbox', { name: 'Link target' }).fill('javascript:alert(1)');
  await dialog.getByRole('button', { name: 'Apply link', exact: true }).click();
  expect((await page.evaluate(() => window.fontLinkControls.state())).fingerprint).toEqual(
    inserted.fingerprint
  );
  await expect(dialog).toBeVisible();
  await dialog.getByRole('textbox', { name: 'Link target' }).fill('https://example.com/new');
  await dialog.getByRole('button', { name: 'Apply link', exact: true }).click();
  await page.evaluate(() => window.fontLinkControls.settle());
  const updated = await page.evaluate(() => window.fontLinkControls.state());
  expect(updated.links[0]!.id).toBe(inserted.links[0]!.id);
  expect(updated.links[0]!.text).toBe('Alpha');
  expect(await page.evaluate(() => window.fontLinkControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.fontLinkControls.state())).fingerprint).toEqual(
    inserted.fingerprint
  );
  expect(await page.evaluate(() => window.fontLinkControls.redo())).toBe(true);
  expect((await page.evaluate(() => window.fontLinkControls.state())).fingerprint).toEqual(
    updated.fingerprint
  );
  const popupPromise = context.waitForEvent('page');
  await toolbar.getByRole('button', { name: 'Open link', exact: true }).click();
  const popup = await popupPromise;
  await popup.waitForLoadState();
  expect(popup.url()).toBe('https://example.com/new');
  expect(await popup.evaluate(() => window.opener === null)).toBe(true);
  await popup.close();
  expect(page.url()).toContain('/font-link.html');
  await page.locator('[data-pane="0"] a[href="https://example.com/new"]').first().click();
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Remove link', exact: true }).click();
  await page.evaluate(() => window.fontLinkControls.settle());
  const unlinked = await page.evaluate(() => window.fontLinkControls.state());
  expect(unlinked.links).toHaveLength(0);
  expect(unlinked.text).toBe('Alpha beta');
  expect(await page.evaluate(() => window.fontLinkControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.fontLinkControls.state())).fingerprint).toEqual(
    updated.fingerprint
  );
  const saved = await page.evaluate(() => window.fontLinkControls.saved());
  expect(saved[0]).toEqual(saved[1]);
  expect(saved[0]).toEqual(updated.fingerprint);
});
