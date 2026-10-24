import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/tables';

for (const merge of ['horizontal', 'vertical', 'none'] as const)
  for (const action of merge === 'none'
    ? (['left', 'right'] as const)
    : (['left', 'right', 'delete'] as const))
    test(`column ${action} beside ${merge} merges uses pointer controls, native typing, exact history and Word export`, async ({
      page,
    }) => {
      await page.goto(`/tables.html?columns=1&merge=${merge}${merge === 'none' ? '&grid=0' : ''}`);
      await page.waitForFunction(() => !!window.tableControls);
      const original = await page.evaluate(() => window.tableControls.columns());
      const target = original.rows[1]!.cells[1]!.paragraphs[0]!;
      const index = await page.evaluate(
        (id) => window.tableControls.state().paragraphs.findIndex((p) => p.id === id),
        target.id
      );
      await page.evaluate((index) => window.tableControls.select(index, 3), index);
      const before = await page.evaluate(() => window.tableControls.state());
      const label = action === 'delete' ? 'Delete column' : `Insert column ${action}`;
      await page
        .locator('[data-pane="0"]')
        .getByRole('toolbar', { name: 'Table columns' })
        .getByRole('button', { name: label, exact: true })
        .click();
      await page.evaluate(() => window.tableControls.settle());
      const after = await page.evaluate(() => window.tableControls.state());
      const columns = await page.evaluate(() => window.tableControls.columns());
      expect(columns.widths).toEqual(action === 'delete' ? [1200, 2400] : [1200, 1800, 1800, 2400]);
      for (const row of columns.rows)
        expect(row.cells.reduce((n, c) => n + c.span, 0)).toBe(columns.widths.length);
      if (merge === 'horizontal') {
        expect(columns.rows[0]!.cells[0]!.span).toBe(
          action === 'left' ? 3 : action === 'right' ? 2 : 1
        );
        expect(columns.rows[0]!.cells[0]!.width).toBe(
          action === 'left' ? 4800 : action === 'right' ? 3000 : 1200
        );
      } else if (merge === 'vertical')
        expect(columns.rows.map((r) => r.cells[0]!.vMerge)).toEqual([
          'restart',
          'continue',
          'continue',
        ]);
      expect(after.anchor).not.toBeNull();
      if (action !== 'delete')
        expect(after.paragraphs.find((p) => p.id === after.selection.head.paragraphId)!.text).toBe(
          ''
        );
      const saved = await page.evaluate(() => window.tableControls.saved());
      expect(saved[0]).toEqual(saved[1]);
      expect(saved[0]).toEqual(after.fingerprint);
      expect(await page.evaluate(() => window.tableControls.undo())).toBe(true);
      expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
        before.fingerprint
      );
      expect(await page.evaluate(() => window.tableControls.redo())).toBe(true);
      expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
        after.fingerprint
      );
      await page.keyboard.type('Live');
      const typed = await page.evaluate(() => window.tableControls.state());
      expect(
        typed.paragraphs.find((p) => p.id === typed.selection.head.paragraphId)!.text
      ).toContain('Live');
      expect(await page.evaluate(() => window.tableControls.undo())).toBe(true);
      expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
        after.fingerprint
      );
      expect(await page.evaluate(() => window.tableControls.redo())).toBe(true);
      expect((await page.evaluate(() => window.tableControls.state())).fingerprint).toEqual(
        typed.fingerprint
      );
      const final = await page.evaluate(() => window.tableControls.saved());
      expect(final[0]).toEqual(final[1]);
      expect(final[0]).toEqual(typed.fingerprint);
    });
