import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/manual-tracking';
for (const decision of ['accept', 'reject'] as const)
  test(`manual suggestion toggle tracks real typing, deletion and bold independently of two AI waves and ${decision}`, async ({
    page,
  }) => {
    await page.goto('/manual-tracking.html');
    await page.waitForFunction(() => !!window.manualTrackingControls);
    await page.getByRole('button', { name: 'Next AI wave', exact: true }).click();
    await page.waitForFunction(
      () =>
        window.manualTrackingControls.state().waveCount === 1 ||
        !!window.manualTrackingControls.state().error
    );
    expect((await page.evaluate(() => window.manualTrackingControls.state())).error).toBe('');
    await page.evaluate(() => window.manualTrackingControls.select(1, 7));
    const selection = (await page.evaluate(() => window.manualTrackingControls.state())).selection;
    const pane = page.locator('[data-pane="0"]'),
      toggle = pane.getByRole('button', { name: 'Suggestion edits', exact: true });
    await toggle.click();
    await expect(toggle).toHaveAttribute('aria-pressed', 'true');
    expect((await page.evaluate(() => window.manualTrackingControls.state())).selection).toEqual(
      selection
    );
    await page.keyboard.type('new ');
    await page.evaluate(() => window.manualTrackingControls.select(2, 0, 7));
    await page.keyboard.press('Backspace');
    await page.evaluate(() => window.manualTrackingControls.select(3, 0, 4));
    await pane.getByRole('button', { name: 'Bold', exact: true }).click();
    await page.evaluate(() => window.manualTrackingControls.settle());
    const manual = (await page.evaluate(() => window.manualTrackingControls.state())).manual;
    expect(manual).toHaveLength(3);
    const before = await page.evaluate(() => window.manualTrackingControls.state());
    await page.getByRole('button', { name: 'Next AI wave', exact: true }).click();
    await page.waitForFunction(
      () =>
        window.manualTrackingControls.state().waveCount === 2 ||
        !!window.manualTrackingControls.state().error
    );
    const recomposed = await page.evaluate(() => window.manualTrackingControls.state());
    expect(recomposed.error).toBe('');
    expect(recomposed.manual).toEqual(manual);
    expect(recomposed.manualParagraphs).toEqual(before.manualParagraphs);
    expect(recomposed.author).toBe('Alice');
    expect(recomposed.mode).toBe(true);
    expect(await page.evaluate(() => window.manualTrackingControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.manualTrackingControls.state())).fingerprint).toEqual(
      before.fingerprint
    );
    expect(await page.evaluate(() => window.manualTrackingControls.redo())).toBe(true);
    expect((await page.evaluate(() => window.manualTrackingControls.state())).fingerprint).toEqual(
      recomposed.fingerprint
    );
    await page
      .getByRole('button', {
        name: decision === 'accept' ? 'Accept AI proposal' : 'Reject AI proposal',
        exact: true,
      })
      .click();
    await page.waitForFunction(
      (previous) => window.manualTrackingControls.state().fingerprint !== previous,
      recomposed.fingerprint
    );
    await page.evaluate(() => window.manualTrackingControls.settle());
    const resolved = await page.evaluate(() => window.manualTrackingControls.state());
    expect(resolved.error).toBe('');
    expect(resolved.manual).toEqual(manual);
    expect(resolved.manualParagraphs).toEqual(before.manualParagraphs);
    expect(await page.evaluate(() => window.manualTrackingControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.manualTrackingControls.state())).fingerprint).toEqual(
      recomposed.fingerprint
    );
    expect(await page.evaluate(() => window.manualTrackingControls.redo())).toBe(true);
    expect((await page.evaluate(() => window.manualTrackingControls.state())).fingerprint).toEqual(
      resolved.fingerprint
    );
    await toggle.click();
    await expect(toggle).toHaveAttribute('aria-pressed', 'false');
    await page.evaluate(() => window.manualTrackingControls.select(4, 4));
    await page.keyboard.type('!');
    await page.evaluate(() => window.manualTrackingControls.settle());
    const final = await page.evaluate(() => window.manualTrackingControls.state());
    expect(final.manual).toEqual(manual);
    expect(final.manualParagraphs).toEqual(before.manualParagraphs);
    expect(final.paragraphs.at(-1)!.text).toBe('Tail!');
    const remote = await page.evaluate(() => window.manualTrackingControls.state(1));
    expect(remote.fingerprint).toEqual(final.fingerprint);
    expect(remote.author).toBe('Bob');
    const saved = await page.evaluate(() => window.manualTrackingControls.saved());
    expect(saved[0]).toEqual(saved[1]);
    expect(saved[0]!.fingerprint).toEqual(final.fingerprint);
  });
