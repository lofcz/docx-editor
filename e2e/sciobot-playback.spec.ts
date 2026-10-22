import { expect, test, type Page } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/playback';

let TARGET = '';

async function open(page: Page) {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/playback.html');
  await page.waitForFunction(() => !!window.playback);
  // Edit the native projection itself: change two paragraphs and delete one.
  const source = await page.evaluate(() => window.playback.source());
  const blocks = source.split('\n\n');
  expect(blocks[1]).toContain('Beta original');
  TARGET = blocks
    .filter((_, index) => index !== 1)
    .join('\n\n')
    .replace(/Alpha original/, 'Alpha changed')
    .replace(/Omega original/, 'Omega changed');
  return errors;
}
const pane = (page: Page) => page.locator('[data-pane="0"]');

test('system reduced motion reveals the committed edit immediately and allows native typing', async ({
  page,
}) => {
  await page.emulateMedia({ reducedMotion: 'reduce' });
  const errors = await open(page);
  expect(
    (await page.evaluate((target) => window.playback.apply(target, 4000), TARGET)).motion
  ).toBe(false);
  expect(await page.evaluate(() => window.playback.state())).toMatchObject({
    sweeps: 0,
    carets: 0,
    pending: 0,
  });
  const alpha = pane(page)
    .locator('.docx-paragraph-fragment')
    .filter({ hasText: 'Alpha changed.' });
  await alpha.click({ position: { x: 20, y: 8 } });
  await page.keyboard.press('End');
  await page.keyboard.type(' Typed');
  expect((await page.evaluate(() => window.playback.state())).texts).toContain(
    'Alpha changed. Typed'
  );
  expect(errors).toEqual([]);
});

test('abort and detach controls stop playback without rolling back shared content', async ({
  page,
}) => {
  const errors = await open(page);
  await page.evaluate((target) => window.playback.apply(target, 4000), TARGET);
  await page.getByRole('button', { name: 'Abort playback', exact: true }).click();
  await page.evaluate(() => window.playback.finished());
  expect(await page.evaluate(() => window.playback.state())).toMatchObject({
    sweeps: 0,
    carets: 0,
    pending: 0,
  });
  const next = TARGET.replace('Alpha changed', 'Alpha again');
  await page.evaluate((target) => window.playback.apply(target, 4000), next);
  await page.getByRole('button', { name: 'Detach editor', exact: true }).click();
  await page.evaluate(() => window.playback.finished());
  expect(await page.evaluate(() => window.playback.state())).toMatchObject({
    sweeps: 0,
    carets: 0,
    pending: 0,
  });
  await page.getByRole('button', { name: 'Attach editor', exact: true }).click();
  expect((await page.evaluate(() => window.playback.state())).texts).toContain('Alpha again.');
  expect(await page.evaluate(() => window.playback.peerTexts())).toEqual(
    (await page.evaluate(() => window.playback.state())).texts
  );
  const saved = await page.evaluate(() => window.playback.savedDigests());
  expect(saved[0]).toEqual(saved[1]);
  expect(errors).toEqual([]);
});

test('changing the system motion preference stops an active playback', async ({ page }) => {
  await page.emulateMedia({ reducedMotion: 'no-preference' });
  const errors = await open(page);
  await page.evaluate((target) => window.playback.apply(target, 4000), TARGET);
  await page.emulateMedia({ reducedMotion: 'reduce' });
  await page.evaluate(() => window.playback.finished());
  expect(await page.evaluate(() => window.playback.state())).toMatchObject({
    sweeps: 0,
    carets: 0,
    pending: 0,
  });
  expect((await page.evaluate(() => window.playback.state())).texts).toContain('Alpha changed.');
  expect(errors).toEqual([]);
});

test('playback keeps every committed target readable and ends within the 4000 ms cap', async ({
  page,
}) => {
  const errors = await open(page);
  const started = await page.evaluate((target) => window.playback.apply(target, 10_000), TARGET);
  expect(started.motion).toBe(true);
  // Text intervals for the two edited paragraphs, then one sibling removal.
  expect(started.kinds.at(-1)).toBe('siblings');
  expect(new Set(started.kinds.slice(0, -1))).toEqual(new Set(['text']));
  // The edit is canonical and painted before any playback step completes.
  const early = await page.evaluate(() => window.playback.state());
  expect(early.texts).toContain('Alpha changed.');
  expect(early.texts).toContain('Omega changed.');
  expect(early.texts).not.toContain('Beta original.');
  expect(early.pending).toBeGreaterThan(0);
  await expect(
    pane(page).locator('.docx-paragraph-fragment').filter({ hasText: 'Alpha changed.' })
  ).toBeVisible();
  await expect(page.locator('[data-sciobot-ai-caret]')).toContainText('ScioBot');
  expect(early.sweeps).toBe(1);
  const elapsed = await page.evaluate(() => window.playback.finished());
  // A requested 10 s budget clamps to the 4000 ms cap.
  expect(elapsed).toBeGreaterThan(1000);
  expect(elapsed).toBeLessThan(4400);
  const done = await page.evaluate(() => window.playback.state());
  expect(done).toMatchObject({ sweeps: 0, carets: 0, pending: 0 });
  // The deleted paragraph leaves no highlight; both edited paragraphs stay highlighted.
  expect(done.highlights).toEqual(['Alpha changed.', 'Omega changed.']);
  // Playback revealed the far target through the native reveal API.
  expect(done.scroll).toBeGreaterThan(0);
  await expect(
    pane(page).locator('.docx-paragraph-fragment').filter({ hasText: 'Omega changed.' })
  ).toBeInViewport();
  expect(errors).toEqual([]);
});

test('cancelling playback retains the edit and real typing settles its change', async ({
  page,
}) => {
  const errors = await open(page);
  await page.evaluate((target) => window.playback.apply(target, 4000), TARGET);
  // Real input into the first changed paragraph during playback.
  const alpha = pane(page)
    .locator('.docx-paragraph-fragment')
    .filter({ hasText: 'Alpha changed.' });
  await alpha.click({ position: { x: 20, y: 8 } });
  await page.keyboard.press('End');
  await page.keyboard.type(' Typed');
  await expect(alpha).toContainText('Alpha changed. Typed');
  await page.evaluate(() => window.playback.cancel());
  const state = await page.evaluate(() => window.playback.state());
  expect(state.sweeps).toBe(0);
  expect(state.carets).toBe(0);
  expect(state.pending).toBe(0);
  expect(state.texts).toContain('Alpha changed. Typed');
  expect(state.texts).toContain('Omega changed.');
  // Typing settled Alpha's change; Omega's committed change remains highlighted.
  expect(state.highlights).toEqual(['Omega changed.']);
  expect(await page.evaluate(() => window.playback.finished())).toBeLessThan(4400);
  expect(errors).toEqual([]);
});

test('a peer deleting a played paragraph leaves no stale overlay or highlight', async ({
  page,
}) => {
  const errors = await open(page);
  await page.evaluate((target) => window.playback.apply(target, 4000), TARGET);
  expect(await page.evaluate(() => window.playback.peerDelete('Omega changed.'))).toBe(true);
  await page.evaluate(() => window.playback.finished());
  const state = await page.evaluate(() => window.playback.state());
  expect(state).toMatchObject({ sweeps: 0, carets: 0, pending: 0 });
  expect(state.texts).not.toContain('Omega changed.');
  expect(state.highlights).toEqual(['Alpha changed.']);
  expect(await page.evaluate(() => window.playback.peerTexts())).toEqual(state.texts);
  const saved = await page.evaluate(() => window.playback.savedDigests());
  expect(saved[0]).toEqual(saved[1]);
  expect(errors).toEqual([]);
});
