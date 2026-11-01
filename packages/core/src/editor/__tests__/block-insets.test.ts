import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { cell, docx, p, row, table } from '../../automation/__tests__/support/protocol.ts';
import { createBlockReservationState } from '../../layout/block-reservations.ts';
import { caretAt } from '../../layout/semantic-interaction.ts';
import { paragraphTextOf } from '../../store/store/tree-op-apply.ts';

const LONG = 'Insets narrow the column the paragraph breaks in, '.repeat(6).trim();
const FULL_WIDTH_TABLE = (text: string) =>
  `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>` +
  table(row(cell(p(text)))).slice('<w:tbl>'.length);

function open(document: Uint8Array) {
  const container = globalThis.document.createElement('div');
  globalThis.document.body.append(container);
  const editor = createDocxEditor({ container, document });
  const surface = editor.surface!;
  const idOf = (prefix: string) =>
    surface.session
      .paragraphIds()
      .find((id) => paragraphTextOf(surface.session.part(), id)?.startsWith(prefix))!;
  const caretOf = (prefix: string, offset = 0) => {
    surface.flushPendingInput();
    return caretAt(surface.layout(), { paragraphId: idOf(prefix), offset })!;
  };
  const tableFragment = () => {
    surface.flushPendingInput();
    for (const page of surface.layout().pages)
      for (const fragment of page.fragments)
        if (fragment.kind === 'table')
          return fragment as { tableId: string; box: { x: number; width: number } };
    throw new Error('no table');
  };
  const lineCount = (prefix: string) => {
    surface.flushPendingInput();
    const id = idOf(prefix);
    return surface
      .layout()
      .pages.flatMap((page) => page.fragments)
      .filter((fragment) => fragment.kind === 'paragraph' && fragment.paragraphId === id)
      .reduce(
        (sum, fragment) => sum + (fragment.kind === 'paragraph' ? fragment.lines.length : 0),
        0
      );
  };
  return {
    editor,
    surface,
    idOf,
    caretOf,
    tableFragment,
    lineCount,
    close: () => {
      editor.destroy();
      container.remove();
    },
  };
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 20));

test('an inset indents only its paragraph and narrows the column it breaks in', async () => {
  const doc = open(docx(p('Alpha') + p(LONG) + p('Gamma')));
  try {
    const before = {
      alpha: doc.caretOf('Alpha').x,
      long: doc.caretOf(LONG).x,
      lines: doc.lineCount(LONG),
    };
    const saved = new Uint8Array(await doc.editor.save());
    const revision = doc.surface.session.packageRevision();
    doc.surface.setBlockInsets(new Map([[doc.idOf(LONG), { left: 40, right: 120 }]]));
    await settle();
    expect(doc.caretOf(LONG).x).toBe(before.long + 40);
    expect(doc.caretOf('Alpha').x).toBe(before.alpha);
    expect(doc.caretOf('Gamma').x).toBe(before.alpha);
    expect(doc.lineCount(LONG)).toBeGreaterThan(before.lines);
    // A view input: no revision, no history entry, and the saved file is unchanged.
    expect(doc.surface.session.packageRevision()).toBe(revision);
    expect(doc.surface.state().canUndo).toBe(false);
    expect(new Uint8Array(await doc.editor.save())).toEqual(saved);
    // Typing inside the inset paragraph keeps the inset.
    doc.surface.setSelection({
      anchor: { paragraphId: doc.idOf(LONG), offset: 0 },
      head: { paragraphId: doc.idOf(LONG), offset: 0 },
    });
    doc.surface.type('X');
    expect(doc.caretOf('X').x).toBe(before.long + 40);
    doc.surface.setBlockInsets(new Map());
    await settle();
    expect(doc.caretOf('X').x).toBe(before.long);
  } finally {
    doc.close();
  }
}, 60_000);

test('an inset moves and narrows an in-flow table', async () => {
  const doc = open(docx(p('Alpha') + FULL_WIDTH_TABLE('Cell') + p('Omega')));
  try {
    const before = doc.tableFragment();
    doc.surface.setBlockInsets(new Map([[before.tableId, { left: 30, right: 20 }]]));
    await settle();
    const after = doc.tableFragment();
    expect(after.box.x).toBeCloseTo(before.box.x + 30, 3);
    expect(after.box.width).toBeCloseTo(before.box.width - 50, 3);
  } finally {
    doc.close();
  }
});

test('an inset that leaves less than an inch of column is ignored', async () => {
  const doc = open(docx(p('Alpha') + p('Beta')));
  try {
    const before = doc.caretOf('Beta').x;
    doc.surface.setBlockInsets(new Map([[doc.idOf('Beta'), { left: 300, right: 300 }]]));
    await settle();
    expect(doc.caretOf('Beta').x).toBe(before);
  } finally {
    doc.close();
  }
});

test('inset values quantize, cap, and drop invalid entries', () => {
  const state = createBlockReservationState();
  expect(
    state.setInsets(
      new Map([
        ['a', { left: 10.1, right: 0 }],
        ['b', { left: Number.NaN, right: -4 }],
        ['c', { left: 9999, right: 1 }],
      ])
    )
  ).toBe(true);
  const option = state.option()!;
  expect(option.insetOf?.('a')).toEqual({ left: 10.25, right: 0 });
  expect(option.insetOf?.('b')).toBeUndefined();
  expect(option.insetOf?.('c')).toEqual({ left: 324, right: 1 });
  // Unchanged values do not publish a new epoch.
  expect(
    state.setInsets(
      new Map([
        ['a', { left: 10.2, right: 0 }],
        ['c', { left: 400, right: 1 }],
      ])
    )
  ).toBe(false);
  expect(state.option()).toBe(option);
  // Reservations and insets share one epoch.
  expect(state.set(new Map([['a', 5]]))).toBe(true);
  expect(state.option()!.epoch).not.toBe(option.epoch);
  expect(state.option()!.insetOf?.('a')).toEqual({ left: 10.25, right: 0 });
});
