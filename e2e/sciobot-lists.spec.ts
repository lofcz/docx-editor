import { expect, test } from '@playwright/test';
import type {} from '../packages/sciobot/test/browser/lists';
for (const heading of [false, true])
  for (const decision of ['direct', 'accept', 'reject'] as const)
    test(`list sequence bridge respects native heading boundaries and coupled review: ${heading}/${decision}`, async ({
      page,
    }) => {
      await page.goto(heading ? '/lists.html?heading=1' : '/lists.html');
      await page.waitForFunction(() => !!window.listControls);
      const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Lists' });
      await toolbar.getByRole('combobox', { name: 'List style' }).selectOption('decimal');
      for (const index of heading ? [1, 0, 3] : [0, 3]) {
        await page.evaluate((index) => window.listControls.select(index, 1), index);
        await toolbar.getByRole('button', { name: 'Apply list', exact: true }).click();
      }
      await page.evaluate((index) => window.listControls.select(index, 1), heading ? 2 : 1);
      if (decision !== 'direct')
        expect(await page.evaluate(() => window.listControls.suggest())).toBe(true);
      const before = await page.evaluate(() => window.listControls.state());
      await toolbar.getByRole('button', { name: 'Apply list', exact: true }).click();
      await page.evaluate(() => window.listControls.settle());
      const after = await page.evaluate(() => window.listControls.state());
      expect(after.paragraphs.map((p) => p.marker?.text ?? null)).toEqual(
        heading ? ['1.', '1.', '1.', '2.'] : ['1.', '2.', 'i.', '3.']
      );
      expect(after.selection).toEqual(before.selection);
      expect(await page.evaluate(() => window.listControls.undo())).toBe(true);
      expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
        before.fingerprint
      );
      expect(await page.evaluate(() => window.listControls.redo())).toBe(true);
      expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
        after.fingerprint
      );
      if (decision !== 'direct') {
        const groups = await page.evaluate(() => window.listControls.groups());
        expect(groups).toHaveLength(1);
        expect(
          await page.evaluate(
            ({ id, action }) =>
              action === 'accept' ? window.listControls.accept(id) : window.listControls.reject(id),
            { id: groups[0]!.id, action: decision }
          )
        ).toBe(true);
        await page.evaluate(() => window.listControls.settle());
      }
      const saved = await page.evaluate(() => window.listControls.saved());
      expect(saved[0]).toEqual(saved[1]);
      if (decision === 'reject') expect(saved[0]!.fingerprint).toEqual(before.fingerprint);
      else {
        await page.keyboard.type('!');
        expect(
          (await page.evaluate(() => window.listControls.state())).paragraphs[heading ? 2 : 1]!.text
        ).toContain('!');
      }
    });
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

test('adjacent counters and mixed indentation keep one-step history through real controls', async ({
  page,
}) => {
  await page.goto('/lists.html');
  await page.waitForFunction(() => !!window.listControls);
  const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Lists' });
  await toolbar.getByRole('combobox', { name: 'List style' }).selectOption('decimal');
  for (const index of [0, 1]) {
    await page.evaluate((index) => window.listControls.select(index, 2), index);
    await toolbar.getByRole('button', { name: 'Apply list', exact: true }).click();
  }
  let state = await page.evaluate(() => window.listControls.state());
  expect(state.paragraphs[0]!.membership!.numbering!.numId).toBe(
    state.paragraphs[1]!.membership!.numbering!.numId
  );
  expect(state.paragraphs[1]!.marker!.text).toBe('2.');
  await page.evaluate(() => window.listControls.selectRange(1, 1, 3, 4));
  const before = await page.evaluate(() => window.listControls.state());
  await toolbar.getByRole('button', { name: 'Indent list', exact: true }).click();
  await page.evaluate(() => window.listControls.settle());
  state = await page.evaluate(() => window.listControls.state());
  expect(state.paragraphs[1]!.membership!.numbering!.level).toBe(1);
  expect(state.paragraphs[2]!.membership!.numbering!.level).toBe(1);
  expect(state.paragraphs[3]!.properties.find((p) => p.localName === 'ind')!.attributes!.left).toBe(
    '360'
  );
  expect(state.selection).toEqual(before.selection);
  expect(await page.evaluate(() => window.listControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
    before.fingerprint
  );
  expect(await page.evaluate(() => window.listControls.redo())).toBe(true);
  expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
    state.fingerprint
  );
  const saved = await page.evaluate(() => window.listControls.saved());
  expect(saved[0]).toEqual(saved[1]);
  expect(saved[0]!.fingerprint).toEqual(state.fingerprint);
});
test('tracked checklist Enter creates one review group and restores native history and caret', async ({
  page,
}) => {
  await page.goto('/lists.html');
  await page.waitForFunction(() => !!window.listControls);
  const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Lists' });
  await page.evaluate(() => window.listControls.select(0, 5));
  await toolbar.getByRole('combobox', { name: 'List style' }).selectOption('todo');
  await toolbar.getByRole('button', { name: 'Apply list', exact: true }).click();
  await toolbar.getByRole('button', { name: 'Check item', exact: true }).click();
  await toolbar.getByRole('button', { name: 'Indent list', exact: true }).click();
  expect(await page.evaluate(() => window.listControls.suggest())).toBe(true);
  const before = await page.evaluate(() => window.listControls.state());
  await page.keyboard.press('Enter');
  await page.evaluate(() => window.listControls.settle());
  const after = await page.evaluate(() => window.listControls.state());
  expect(after.paragraphs[1]!.membership!.checkbox).toMatchObject({ checked: false, level: 1 });
  expect(after.selection.head.offset).toBe(1);
  const groups = await page.evaluate(() => window.listControls.groups());
  expect(groups).toHaveLength(1);
  expect(await page.evaluate(() => window.listControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
    before.fingerprint
  );
  expect(await page.evaluate(() => window.listControls.redo())).toBe(true);
  expect((await page.evaluate(() => window.listControls.state())).selection).toEqual(
    after.selection
  );
  expect(await page.evaluate((id) => window.listControls.reject(id), groups[0]!.id)).toBe(true);
  await page.evaluate(() => window.listControls.settle());
  const saved = await page.evaluate(() => window.listControls.saved());
  expect(saved[0]).toEqual(saved[1]);
  expect(await page.evaluate(() => window.listControls.undo())).toBe(true);
  expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
    after.fingerprint
  );
});

for (const [style, label] of [
  ['disc', 'Insert bullets'],
  ['decimal', 'Insert numbers'],
  ['todo', 'Insert checklist'],
] as const)
  for (const decision of ['direct', 'accept', 'reject'] as const)
    test(`real ${style} insertion replaces the selection and retains native review/history: ${decision}`, async ({
      page,
    }) => {
      await page.goto('/lists.html');
      await page.waitForFunction(() => !!window.listControls);
      await page.evaluate(() => window.listControls.select(0, 0, 5));
      await page.keyboard.type('Alpha beta');
      await page.evaluate(() => window.listControls.select(0, 0, 5));
      if (decision !== 'direct')
        expect(await page.evaluate(() => window.listControls.suggest())).toBe(true);
      const before = await page.evaluate(() => window.listControls.state());
      const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Lists' });
      await toolbar.getByRole('button', { name: label, exact: true }).click();
      await page.evaluate(() => window.listControls.settle());
      const after = await page.evaluate(() => window.listControls.state());
      expect(after.paragraphs[1]!.text).toBe(' beta');
      expect(
        after.paragraphs[0]!.membership?.checkbox?.checked ?? after.paragraphs[0]!.marker?.text
      ).toBe(style === 'todo' ? false : style === 'disc' ? '•' : '1.');
      expect(await page.evaluate(() => window.listControls.undo())).toBe(true);
      expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
        before.fingerprint
      );
      expect(await page.evaluate(() => window.listControls.redo())).toBe(true);
      expect((await page.evaluate(() => window.listControls.state())).selection).toEqual(
        after.selection
      );
      if (decision !== 'direct') {
        const groups = await page.evaluate(() => window.listControls.groups());
        expect(groups).toHaveLength(1);
        expect(
          await page.evaluate(
            ({ id, action }) =>
              action === 'accept' ? window.listControls.accept(id) : window.listControls.reject(id),
            { id: groups[0]!.id, action: decision }
          )
        ).toBe(true);
        await page.evaluate(() => window.listControls.settle());
      }
      const saved = await page.evaluate(() => window.listControls.saved());
      expect(saved[0]).toEqual(saved[1]);
      if (decision === 'reject') expect(saved[0]!.paragraphs[0]).toBe('Alpha beta');
      else {
        const resolved = await page.evaluate(() => window.listControls.state());
        expect(resolved.paragraphs[0]!.text).toBe(style === 'todo' ? '☐' : '');
        await page.evaluate(
          (offset) => window.listControls.select(0, offset),
          style === 'todo' ? 1 : 0
        );
        await page.keyboard.type('Next');
        expect((await page.evaluate(() => window.listControls.state())).paragraphs[0]!.text).toBe(
          style === 'todo' ? '☐Next' : 'Next'
        );
      }
    });

for (const [style, label] of [
  ['decimal', 'Insert numbers'],
  ['todo', 'Insert checklist'],
] as const)
  for (const decision of ['direct', 'accept', 'reject'] as const)
    test(`empty list insertion retains the previous ${style} item and one-step history: ${decision}`, async ({
      page,
    }) => {
      await page.goto('/lists.html');
      await page.waitForFunction(() => !!window.listControls);
      const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Lists' });
      await page.evaluate(() => window.listControls.select(0, 0, 5));
      await page.keyboard.press('Backspace');
      await toolbar.getByRole('combobox', { name: 'List style' }).selectOption(style);
      await toolbar.getByRole('button', { name: 'Apply list', exact: true }).click();
      if (style === 'todo')
        await toolbar.getByRole('button', { name: 'Check item', exact: true }).click();
      await page.evaluate(
        (offset) => window.listControls.select(0, offset),
        style === 'todo' ? 1 : 0
      );
      if (decision !== 'direct')
        expect(await page.evaluate(() => window.listControls.suggest())).toBe(true);
      const before = await page.evaluate(() => window.listControls.state());
      await toolbar.getByRole('button', { name: label, exact: true }).click();
      await page.evaluate(() => window.listControls.settle());
      const after = await page.evaluate(() => window.listControls.state());
      expect(after.paragraphs).toHaveLength(before.paragraphs.length + 1);
      expect(after.paragraphs[0]!.membership).toEqual(before.paragraphs[0]!.membership);
      expect(
        style === 'todo'
          ? after.paragraphs[0]!.membership?.checkbox?.checked
          : after.paragraphs[0]!.marker?.text
      ).toBe(style === 'todo' ? true : '1.');
      expect(
        style === 'todo'
          ? after.paragraphs[1]!.membership?.checkbox?.checked
          : after.paragraphs[1]!.marker?.text
      ).toBe(style === 'todo' ? false : '2.');
      expect(await page.evaluate(() => window.listControls.undo())).toBe(true);
      expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
        before.fingerprint
      );
      expect(await page.evaluate(() => window.listControls.redo())).toBe(true);
      expect((await page.evaluate(() => window.listControls.state())).selection).toEqual(
        after.selection
      );
      if (decision !== 'direct') {
        const groups = await page.evaluate(() => window.listControls.groups());
        expect(groups).toHaveLength(1);
        expect(
          await page.evaluate(
            ({ id, action }) =>
              action === 'accept' ? window.listControls.accept(id) : window.listControls.reject(id),
            { id: groups[0]!.id, action: decision }
          )
        ).toBe(true);
        await page.evaluate(() => window.listControls.settle());
      }
      const saved = await page.evaluate(() => window.listControls.saved());
      expect(saved[0]).toEqual(saved[1]);
      if (decision === 'reject') expect(saved[0]!.fingerprint).toEqual(before.fingerprint);
      else {
        await page.evaluate(
          (offset) => window.listControls.select(1, offset),
          style === 'todo' ? 1 : 0
        );
        await page.keyboard.type('Next');
        expect((await page.evaluate(() => window.listControls.state())).paragraphs[1]!.text).toBe(
          style === 'todo' ? '☐Next' : 'Next'
        );
      }
    });

for (const prior of [false, true])
  test(`captured checklist indentation uses the real toolbar and preserves native typing: prior=${prior}`, async ({
    page,
  }) => {
    await page.goto('/lists.html');
    await page.waitForFunction(() => !!window.listControls);
    const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Lists' });
    if (prior) {
      await page.evaluate(() => window.listControls.select(0, 2));
      await toolbar.getByRole('combobox', { name: 'List style' }).selectOption('todo');
      await toolbar.getByRole('button', { name: 'Apply list', exact: true }).click();
      await page.evaluate(() => window.listControls.settle());
    }
    await page.evaluate((prior) => window.listControls.select(0, prior ? 3 : 2), prior);
    const before = await page.evaluate(() => window.listControls.state());
    await toolbar.getByRole('button', { name: 'Indent checklist', exact: true }).click();
    await page.evaluate(() => window.listControls.settle());
    const after = await page.evaluate(() => window.listControls.state());
    expect(after.paragraphs[0]!.membership).toMatchObject({
      checkbox: null,
      dormantChecked: false,
      numbering: { level: prior ? 1 : 0 },
    });
    expect(after.paragraphs[0]!.marker?.text).toBe('•');
    expect(after.selection).toEqual({
      anchor: { paragraphId: before.paragraphs[0]!.id, offset: 2 },
      head: { paragraphId: before.paragraphs[0]!.id, offset: 2 },
    });
    expect(await page.evaluate(() => window.listControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
      before.fingerprint
    );
    expect(await page.evaluate(() => window.listControls.redo())).toBe(true);
    expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
      after.fingerprint
    );
    await page.keyboard.type('!');
    await page.evaluate(() => window.listControls.settle());
    expect((await page.evaluate(() => window.listControls.state())).paragraphs[0]!.text).toBe(
      'Al!pha'
    );
    const saved = await page.evaluate(() => window.listControls.saved());
    expect(saved[0]).toEqual(saved[1]);
  });

for (const decision of ['direct', 'accept', 'reject'] as const)
  test(`source list toggles convert siblings and split the removed counter through the toolbar: ${decision}`, async ({
    page,
  }) => {
    await page.goto('/lists.html');
    await page.waitForFunction(() => !!window.listControls);
    const toolbar = page.locator('[data-pane="0"]').getByRole('toolbar', { name: 'Lists' });
    await page.evaluate(() => window.listControls.selectRange(0, 1, 1, 2));
    await toolbar.getByRole('combobox', { name: 'List style' }).selectOption('decimal');
    await toolbar.getByRole('button', { name: 'Apply list', exact: true }).click();
    await page.evaluate(() => window.listControls.settle());
    await page.evaluate(() => window.listControls.select(1, 2));
    await toolbar.getByRole('combobox', { name: 'List style' }).selectOption('lower-roman');
    await toolbar.getByRole('button', { name: 'Toggle list', exact: true }).click();
    await page.evaluate(() => window.listControls.settle());
    const converted = await page.evaluate(() => window.listControls.state());
    expect(converted.paragraphs.slice(0, 3).map((p) => p.marker?.text)).toEqual([
      'i.',
      'ii.',
      'iii.',
    ]);
    if (decision !== 'direct')
      expect(await page.evaluate(() => window.listControls.suggest())).toBe(true);
    await toolbar.getByRole('button', { name: 'Toggle list', exact: true }).click();
    await page.evaluate(() => window.listControls.settle());
    const pending = await page.evaluate(() => window.listControls.state());
    expect(pending.paragraphs.slice(0, 3).map((p) => p.marker?.text ?? null)).toEqual([
      'i.',
      null,
      'i.',
    ]);
    expect(pending.selection).toEqual(converted.selection);
    if (decision !== 'direct') {
      const groups = await page.evaluate(() => window.listControls.groups());
      expect(groups).toHaveLength(1);
      expect(
        await page.evaluate(
          ({ id, decision }) =>
            decision === 'accept' ? window.listControls.accept(id) : window.listControls.reject(id),
          { id: groups[0]!.id, decision }
        )
      ).toBe(true);
      await page.evaluate(() => window.listControls.settle());
    }
    const final = await page.evaluate(() => window.listControls.state());
    if (decision === 'reject') expect(final.fingerprint).toEqual(converted.fingerprint);
    expect(await page.evaluate(() => window.listControls.undo())).toBe(true);
    expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
      decision === 'direct' ? converted.fingerprint : pending.fingerprint
    );
    expect(await page.evaluate(() => window.listControls.redo())).toBe(true);
    expect((await page.evaluate(() => window.listControls.state())).fingerprint).toEqual(
      final.fingerprint
    );
    await page.keyboard.type('!');
    await page.evaluate(() => window.listControls.settle());
    expect((await page.evaluate(() => window.listControls.state())).paragraphs[1]!.text).toBe(
      'Be!ta'
    );
    const saved = await page.evaluate(() => window.listControls.saved());
    expect(saved[0]).toEqual(saved[1]);
  });
