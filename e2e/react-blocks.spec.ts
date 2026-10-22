import { expect, test } from '@playwright/test';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { fixture } from '../examples/react-blocks/src/document';

test('React blocks wrap native paragraphs and keep local React state through text edits', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/');
  await expect(page.getByRole('button', { name: 'Local React state: 0' })).toBeVisible();
  await page.getByRole('button', { name: 'Local React state: 0' }).click();
  const paragraph = page
    .locator(
      '[data-react-block="phase"] [data-react-block="activity"] [data-native-slot] .docx-paragraph-fragment'
    )
    .filter({ hasText: 'Plan an experiment.' });
  await expect(paragraph).toHaveCount(1);
  await paragraph.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Typed inside React.');
  await expect(paragraph).toContainText('Typed inside React.');
  await expect(page.getByRole('button', { name: 'Local React state: 1' })).toBeVisible();
  expect(errors).toEqual([]);
  await page.screenshot({ path: 'screenshots/react-blocks-proof.png', fullPage: true });
});

test('native suggestions and review decisions work inside a React container', async ({ page }) => {
  await page.goto('/');
  await expect(page.getByRole('button', { name: 'Local React state: 0' })).toBeVisible();
  await page.getByRole('button', { name: 'Propose AI text edit' }).click();
  await expect(page.getByTestId('review-count')).not.toHaveText('0 pending changes');
  await expect(page.locator('[data-react-block="activity"]')).toContainText(
    'controlled experiment'
  );
  await page.getByRole('button', { name: 'Reject changes', exact: true }).click();
  await expect(page.getByTestId('review-count')).toHaveText('0 pending changes');
  await expect(page.locator('[data-react-block="activity"]')).not.toContainText(
    'controlled experiment'
  );
  await page.getByRole('button', { name: 'Propose AI text edit' }).click();
  await page.getByRole('button', { name: 'Accept changes', exact: true }).click();
  await expect(page.getByTestId('review-count')).toHaveText('0 pending changes');
  await expect(page.locator('[data-react-block="activity"]')).toContainText(
    'controlled experiment'
  );
});

test('React data and native content survive saving and reopening DOCX', async ({ page }) => {
  await page.goto('/');
  await page.getByLabel('Activity duration').fill('15');
  await expect(page.getByLabel('Activity duration')).toHaveValue('15');
  await page.getByLabel('Prediction', { exact: true }).selectOption('warmer');
  await expect(page.getByLabel('Prediction', { exact: true })).toHaveValue('warmer');
  await page.getByRole('button', { name: 'Save and reopen DOCX' }).click();
  await expect(page.getByLabel('Activity duration')).toHaveValue('15');
  await expect(page.getByLabel('Prediction', { exact: true })).toHaveValue('warmer');
  await expect(page.locator('[data-react-block="activity"] [data-native-slot]')).toContainText(
    'Record the result'
  );
});

test('local React state survives moving an activity to another page', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Local React state: 0' }).click();
  const state = page.getByRole('button', { name: 'Local React state: 1' });
  const previous = await state.evaluate((element) =>
    Number(element.closest('.docx-page')?.getAttribute('data-page-index'))
  );
  await page.getByRole('button', { name: 'Push blocks to next page' }).click();
  await expect(state).toHaveCount(1);
  await expect
    .poll(() =>
      state.evaluate((element) =>
        Number(element.closest('.docx-page')?.getAttribute('data-page-index'))
      )
    )
    .toBeGreaterThan(previous);
});

test('arbitrary React height updates native pagination and moves following content', async ({
  page,
}) => {
  await page.goto('/');
  const toggle = page.getByRole('button', { name: 'Show explanation' });
  await toggle.scrollIntoViewIfNeeded();
  const after = page
    .locator('.docx-paragraph-fragment')
    .filter({ hasText: 'Reflection inside the phase.' });
  const before = await after.boundingBox();
  const pagesBefore = await page.getByTestId('page-count').textContent();
  await toggle.click();
  await expect(page.getByLabel('Scratch notes')).toBeVisible();
  await expect
    .poll(async () => {
      const next = await after.boundingBox();
      return next && before ? next.y - before.y : 0;
    })
    .toBeGreaterThan(70);
  const explanation = await page.locator('.quiz-explanation').boundingBox();
  const following = await after.boundingBox();
  expect(following!.y).toBeGreaterThan(explanation!.y + explanation!.height);
  console.log({ pagesBefore, pagesAfter: await page.getByTestId('page-count').textContent() });
});

test('an unregistered React block tag retains its native Word fallback', async ({ page }) => {
  await page.goto('/');
  await expect(page.getByRole('button', { name: 'Local React state: 0' })).toBeVisible();
  const parts = unzipSync(fixture());
  parts['word/document.xml'] = strToU8(
    strFromU8(parts['word/document.xml']!).replace('scio:activity:one', 'scio:unregistered:one')
  );
  const bytes = Array.from(zipSync(parts));
  await page.evaluate((bytes) => {
    const proof = (window as unknown as { proof: { load(bytes: ArrayBuffer): void } }).proof;
    proof.load(Uint8Array.from(bytes).buffer);
  }, bytes);
  await expect(page.getByRole('button', { name: 'Local React state: 0' })).toHaveCount(0);
  await expect(
    page.locator('.docx-paragraph-fragment').filter({ hasText: 'Plan an experiment.' })
  ).toHaveCount(1);
  await expect(
    page.locator('.docx-paragraph-fragment').filter({ hasText: 'Record the result' })
  ).toHaveCount(1);
});
