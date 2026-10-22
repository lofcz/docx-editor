import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { docx, p } from '../../automation/__tests__/support/protocol.ts';
import {
  createBlockReservationState,
  quantizeBlockReservation,
} from '../../layout/block-reservations.ts';
import { caretAt } from '../../layout/semantic-interaction.ts';
import { paragraphTextOf } from '../../store/store/tree-op-apply.ts';

const contextual = (text: string) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="400"/><w:contextualSpacing/></w:pPr>` +
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

function open(document: Uint8Array) {
  const container = globalThis.document.createElement('div');
  globalThis.document.body.append(container);
  const editor = createDocxEditor({ container, document });
  const surface = editor.surface!;
  const idOf = (text: string) =>
    surface.session
      .paragraphIds()
      .find((id) => paragraphTextOf(surface.session.part(), id) === text)!;
  const caretOf = (text: string) => {
    surface.flushPendingInput();
    return caretAt(surface.layout(), { paragraphId: idOf(text), offset: 0 })!;
  };
  let frames = 0;
  const stop = surface.observePaint({ afterPaint: () => (frames += 1) });
  return {
    editor,
    surface,
    idOf,
    caretOf,
    yOf: (text: string) => caretOf(text).y,
    frames: () => frames,
    close: () => {
      stop();
      editor.destroy();
      container.remove();
    },
  };
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 20));
const bytesOf = async (doc: ReturnType<typeof open>) => new Uint8Array(await doc.editor.save());

test('a reservation moves only the following content, by exactly its points', async () => {
  const doc = open(docx(p('Alpha') + p('Beta') + p('Gamma')));
  try {
    const before = [doc.yOf('Alpha'), doc.yOf('Beta'), doc.yOf('Gamma')];
    const revision = doc.surface.session.packageRevision();
    doc.surface.setBlockReservations(new Map([[doc.idOf('Alpha'), 30]]));
    await settle();
    expect([doc.yOf('Alpha'), doc.yOf('Beta'), doc.yOf('Gamma')]).toEqual([
      before[0]!,
      before[1]! + 30,
      before[2]! + 30,
    ]);
    // Painted, not merely laid out.
    const painted = doc.surface
      .hostElement()
      .querySelector<HTMLElement>(`[data-paragraph-id="${doc.idOf('Beta')}"]`);
    expect(painted).not.toBeNull();
    // A view input: no revision, no history entry.
    expect(doc.surface.session.packageRevision()).toBe(revision);
    expect(doc.surface.state().canUndo).toBe(false);
    doc.surface.setBlockReservations(new Map([[doc.idOf('Beta'), 12]]));
    await settle();
    expect([doc.yOf('Beta'), doc.yOf('Gamma')]).toEqual([before[1]!, before[2]! + 12]);
    doc.surface.setBlockReservations(new Map());
    await settle();
    expect([doc.yOf('Alpha'), doc.yOf('Beta'), doc.yOf('Gamma')]).toEqual(before);
    expect(doc.surface.state().canUndo).toBe(false);
  } finally {
    doc.close();
  }
});

test('contextual spacing never absorbs a reservation', async () => {
  const doc = open(docx(contextual('One') + contextual('Two') + contextual('Three')));
  try {
    const gap = doc.yOf('Two') - doc.yOf('One');
    doc.surface.setBlockReservations(new Map([[doc.idOf('One'), 18]]));
    await settle();
    // Contextual spacing still drops the authored 20 pt gap; only the reservation remains.
    expect(gap).toBeLessThan(20);
    expect(doc.yOf('Two') - doc.yOf('One')).toBeCloseTo(gap + 18, 6);
    expect(doc.yOf('Three') - doc.yOf('Two')).toBeCloseTo(gap, 6);
  } finally {
    doc.close();
  }
});

test('equal or noisy measurements never relayout, and values round up to a quarter point', async () => {
  expect(quantizeBlockReservation(10.1)).toBe(10.25);
  expect(quantizeBlockReservation(10.25)).toBe(10.25);
  expect(quantizeBlockReservation(10.2500001)).toBe(10.25);
  for (const invalid of [0, -4, Number.NaN, Number.POSITIVE_INFINITY])
    expect(quantizeBlockReservation(invalid)).toBe(0);
  expect(quantizeBlockReservation(1e9)).toBe(4 * 1584);
  const state = createBlockReservationState();
  expect(state.set(new Map([['a', 10.1]]))).toBe(true);
  const epoch = state.option()!.epoch;
  expect(state.set(new Map([['a', 10.2]]))).toBe(false);
  expect(state.option()!.epoch).toBe(epoch);
  expect(state.set(new Map([['a', -1]]))).toBe(true);
  expect(state.option()).toBeUndefined();
  expect(state.set(new Map())).toBe(false);

  const doc = open(docx(p('Alpha') + p('Beta')));
  try {
    const base = doc.yOf('Beta');
    doc.surface.setBlockReservations(new Map([[doc.idOf('Alpha'), 20.1]]));
    await settle();
    expect(doc.yOf('Beta')).toBe(base + 20.25);
    const frames = doc.frames();
    doc.surface.setBlockReservations(new Map([[doc.idOf('Alpha'), 20.2]]));
    doc.surface.setBlockReservations(new Map([[doc.idOf('Alpha'), 20.25]]));
    await settle();
    expect(doc.frames()).toBe(frames);
    expect(doc.yOf('Beta')).toBe(base + 20.25);
    // A reservation for an unknown paragraph is inert.
    doc.surface.setBlockReservations(new Map([['missing-paragraph', 40]]));
    await settle();
    expect(doc.yOf('Beta')).toBe(base);
  } finally {
    doc.close();
  }
});

test('a reservation larger than the rest of the page carries the following text to the next page', async () => {
  const lines = Array.from({ length: 8 }, (_, index) => p(`Line ${index + 1}`)).join('');
  const doc = open(docx(lines + p('Next')));
  try {
    expect(doc.caretOf('Next').pageIndex).toBe(0);
    doc.surface.setBlockReservations(new Map([[doc.idOf('Line 8'), 900]]));
    await settle();
    const next = doc.caretOf('Next');
    expect(next.pageIndex).toBe(1);
    // It opens the next page at its top line, with no blank page in between.
    expect(next.y).toBe(doc.caretOf('Line 1').y);
    expect(doc.surface.layout().pages).toHaveLength(2);
  } finally {
    doc.close();
  }
});

test('user history and saved files never see a reservation', async () => {
  const doc = open(docx(p('Alpha') + p('Beta') + p('Gamma')));
  try {
    const plain = await bytesOf(doc);
    const base = [doc.yOf('Beta'), doc.yOf('Gamma')];
    doc.surface.setBlockReservations(new Map([[doc.idOf('Alpha'), 24]]));
    await settle();
    expect(await bytesOf(doc)).toEqual(plain);

    const beta = doc.idOf('Beta');
    doc.surface.setSelection({
      anchor: { paragraphId: beta, offset: 4 },
      head: { paragraphId: beta, offset: 4 },
    });
    doc.surface.insertPlainText('!');
    doc.surface.flushPendingInput();
    expect(doc.yOf('Beta!')).toBe(base[0]! + 24);
    // One user edit, one undo step; the reservation stays in place throughout.
    expect(doc.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(doc.surface.state().canUndo).toBe(false);
    expect([doc.yOf('Beta'), doc.yOf('Gamma')]).toEqual([base[0]! + 24, base[1]! + 24]);
    expect(doc.editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(doc.yOf('Beta!')).toBe(base[0]! + 24);
    expect(doc.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(await bytesOf(doc)).toEqual(plain);

    const reopened = open(await bytesOf(doc));
    try {
      expect([reopened.yOf('Beta'), reopened.yOf('Gamma')]).toEqual(base);
    } finally {
      reopened.close();
    }
  } finally {
    doc.close();
  }
});

test('a destroyed surface ignores reservations', () => {
  const doc = open(docx(p('Alpha')));
  const id = doc.idOf('Alpha');
  doc.close();
  expect(() => doc.surface.setBlockReservations(new Map([[id, 10]]))).not.toThrow();
});
