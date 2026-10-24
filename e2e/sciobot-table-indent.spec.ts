import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/table-indent';
for (const merged of [false, true])
  test(`native table margin controls move ${merged ? 'merged' : 'plain'} cells by 24px, retain input selection, undo, reset and reopen`, async ({
    page,
  }) => {
    await page.goto('/table-indent.html' + (merged ? '?merged=1' : ''));
    await page.waitForFunction(() => !!window.tableIndentControls);
    await page.evaluate(() => window.tableIndentControls.select());
    const before = await page.evaluate(() => window.tableIndentControls.state());
    const pane = page.locator('[data-pane="0"]'),
      toolbar = pane.getByRole('toolbar', { name: 'Table indentation' });
    const paragraph = pane
      .locator(`[data-paragraph-id="${before.selection.head.paragraphId}"]`)
      .first();
    const box = await paragraph.boundingBox();
    expect(box).not.toBeNull();
    await toolbar.getByRole('spinbutton', { name: 'Table left margin (px)' }).fill('24');
    await toolbar.getByRole('button', { name: 'Apply table margin', exact: true }).click();
    await page.evaluate(() => window.tableIndentControls.settle());
    const moved = await page.evaluate(() => window.tableIndentControls.state());
    expect(moved.indent).toBe(24);
    expect(moved.selection).toEqual(before.selection);
    expect(moved.paragraphs).toEqual(before.paragraphs);
    expect(moved.table.width).toBe(before.table.width);
    expect(moved.table.x - before.table.x).toBeCloseTo(18, 6);
    expect((await paragraph.boundingBox())!.x - box!.x).toBeCloseTo(24, 1);
    expect(moved.fingerprint).toEqual(
      await page.evaluate(() => window.tableIndentControls.expected(360))
    );
    expect(await page.evaluate(() => window.tableIndentControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.tableIndentControls.state())).fingerprint).toEqual(
      before.fingerprint
    );
    expect(await page.evaluate(() => window.tableIndentControls.redo())).toBe(true);
    expect((await page.evaluate(() => window.tableIndentControls.state())).fingerprint).toEqual(
      moved.fingerprint
    );
    await toolbar.getByRole('button', { name: 'Reset table margin', exact: true }).click();
    await page.evaluate(() => window.tableIndentControls.settle());
    const reset = await page.evaluate(() => window.tableIndentControls.state());
    expect(reset.fingerprint).toEqual(before.fingerprint);
    expect(reset.table).toEqual(before.table);
    expect(await page.evaluate(() => window.tableIndentControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.tableIndentControls.state())).fingerprint).toEqual(
      moved.fingerprint
    );
    await page.keyboard.type('Live');
    await page.evaluate(() => window.tableIndentControls.settle());
    const typed = await page.evaluate(() => window.tableIndentControls.state());
    expect(
      typed.paragraphs.find((p) => p.id === before.selection.head.paragraphId)!.text
    ).toContain('Live');
    expect((await page.evaluate(() => window.tableIndentControls.state(1))).fingerprint).toEqual(
      typed.fingerprint
    );
    expect(await page.evaluate(() => window.tableIndentControls.saved())).toEqual([
      typed.fingerprint,
      typed.fingerprint,
    ]);
  });
