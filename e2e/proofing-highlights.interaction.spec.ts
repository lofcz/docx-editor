import { expect, test } from '@playwright/test';
import { resolve } from 'node:path';

test('proofing underlines follow marked text and leave other spans unstyled', async ({ page }) => {
  await page.setContent(`
    <section class="docx-editor">
      <span aria-invalid="spelling"><b>Overlapping</b></span> <span>Unmarked</span>
    </section>
    <section class="docx-editor">
      <span aria-invalid="grammar"><b>Overlapping</b></span> <span>Unmarked</span>
    </section>
  `);
  await page.addStyleTag({
    path: resolve(import.meta.dirname, '../packages/core/src/styles/editor.css'),
  });
  await page.addStyleTag({ content: '.docx-editor { font: 32px/1.6 sans-serif; }' });
  const result = await page.evaluate(() => {
    const spelling = new Highlight();
    const grammar = new Highlight();
    spelling.type = 'spelling-error';
    grammar.type = 'grammar-error';
    CSS.highlights.set('docx-spelling-error', spelling);
    CSS.highlights.set('docx-grammar-error', grammar);
    const decoration = (el: Element, kind: string) => {
      const style = getComputedStyle(el, `::highlight(docx-${kind}-error)`);
      return [style.textDecorationLine, style.textDecorationStyle, style.textDecorationColor];
    };
    return [...document.querySelectorAll('section')].map((section) => {
      const nested = section.querySelector('b')!;
      const unmarked = section.lastElementChild!;
      const text = nested.firstChild!;
      const spellingRange = new Range();
      spellingRange.setStart(text, 0);
      spellingRange.setEnd(text, 7);
      spelling.add(spellingRange);
      const grammarRange = new Range();
      grammarRange.setStart(text, 4);
      grammarRange.setEnd(text, text.textContent!.length);
      grammar.add(grammarRange);
      const selection = getSelection()!;
      selection.removeAllRanges();
      selection.addRange(grammarRange);
      const selected = selection.toString();
      selection.removeAllRanges();
      return {
        spelling: decoration(nested, 'spelling'),
        grammar: decoration(nested, 'grammar'),
        unmarked: [decoration(unmarked, 'spelling')[0], decoration(unmarked, 'grammar')[0]],
        selected,
      };
    });
  });
  for (const row of result) {
    expect(row.selected).toBe('lapping');
    expect(row.spelling.slice(0, 2)).toEqual(['underline', 'wavy']);
    expect(row.grammar.slice(0, 2)).toEqual(['underline', 'wavy']);
    expect(row.spelling[2]).not.toBe(row.grammar[2]);
    expect(row.unmarked).toEqual(['none', 'none']);
  }
  const text = page.locator('section').first();
  const ink = await text.screenshot();
  await page.evaluate(() => CSS.highlights.clear());
  expect(await text.screenshot()).not.toEqual(ink);
});
