import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
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

test('view header padding survives contextual list spacing and the first page edge', async () => {
  const item = (text: string) =>
    `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:contextualSpacing/>` +
    `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>` +
    `<w:spacing w:before="120" w:after="120"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const doc = open(withNumbering(docx(item('Question') + item('Answer') + item('Next'))));
  try {
    const before = ['Question', 'Answer', 'Next'].map((text) => doc.caretOf(text).y);
    const revision = doc.surface.session.packageRevision();
    const saved = new Uint8Array(await doc.editor.save());
    doc.surface.setBlockInsets(
      new Map([
        [doc.idOf('Question'), { left: 0, right: 0, top: 24 }],
        [doc.idOf('Answer'), { left: 0, right: 0, top: 30, bottom: 9 }],
      ])
    );
    await settle();
    expect(doc.caretOf('Question').y - before[0]!).toBeCloseTo(24, 6);
    expect(doc.caretOf('Answer').y - before[1]!).toBeCloseTo(54, 6);
    expect(doc.caretOf('Next').y - before[2]!).toBeCloseTo(63, 6);
    expect(doc.surface.session.packageRevision()).toBe(revision);
    expect(doc.surface.state().canUndo).toBe(false);
    expect(new Uint8Array(await doc.editor.save())).toEqual(saved);
  } finally {
    doc.close();
  }
});

test('view bottom padding stays separate from the following authored space before', async () => {
  const doc = open(
    docx(
      p('Answer') +
        '<w:p><w:pPr><w:spacing w:before="600" w:after="0"/></w:pPr><w:r><w:t>Next</w:t></w:r></w:p>'
    )
  );
  try {
    const gap = doc.caretOf('Next').y - doc.caretOf('Answer').y;
    doc.surface.setBlockInsets(new Map([[doc.idOf('Answer'), { left: 0, right: 0, bottom: 12 }]]));
    await settle();
    expect(doc.caretOf('Next').y - doc.caretOf('Answer').y).toBeCloseTo(gap + 12, 6);
  } finally {
    doc.close();
  }
});

test('a generated label stays with body padding when its kept group moves to the next page', async () => {
  const preceding = Array.from({ length: 18 }, (_, index) => p(`Line ${index}`)).join('');
  const label = '<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>Generated label</w:t></w:r></w:p>';
  const doc = open(docx(preceding + label + p('Body')));
  try {
    expect(doc.caretOf('Generated label').pageIndex).toBe(0);
    const body = doc.caretOf('Body');
    const room = doc.surface.layout().pages[0]!.contentBox.height - body.y;
    expect(room + 12).toBeLessThan(324);
    doc.surface.setBlockInsets(
      new Map([
        [doc.idOf('Generated label'), { left: 0, right: 0, top: 8, bottom: 8 }],
        [doc.idOf('Body'), { left: 0, right: 0, top: room + 12 }],
      ])
    );
    await settle();
    expect(doc.caretOf('Generated label').pageIndex).toBe(1);
    expect(doc.caretOf('Body').pageIndex).toBe(1);
  } finally {
    doc.close();
  }
});

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

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
/** `docx()` with a one-level decimal list in `numbering.xml`. */
function withNumbering(bytes: Uint8Array): Uint8Array {
  const files = unzipSync(bytes);
  const add = (name: string, find: string, insert: string) => {
    files[name] = strToU8(strFromU8(files[name]!).replace(find, insert + find));
  };
  add(
    '[Content_Types].xml',
    '</Types>',
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
  );
  add(
    'word/_rels/document.xml.rels',
    '</Relationships>',
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
  );
  files['word/numbering.xml'] = strToU8(
    `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
      '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
      '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
  );
  return zipSync(files);
}

test('a list marker moves with its inset paragraph', async () => {
  const item = (text: string) =>
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
    `<w:r><w:t>${text}</w:t></w:r></w:p>`;
  const doc = open(withNumbering(docx(item('First') + item('Second'))));
  try {
    const markerX = (prefix: string) => {
      doc.surface.flushPendingInput();
      const id = doc.idOf(prefix);
      for (const page of doc.surface.layout().pages)
        for (const fragment of page.fragments)
          if (fragment.kind === 'paragraph' && fragment.paragraphId === id && fragment.marker)
            return fragment.marker.box.x;
      throw new Error(`no marker for ${prefix}`);
    };
    const before = { marker: markerX('Second'), text: doc.caretOf('Second').x };
    doc.surface.setBlockInsets(new Map([[doc.idOf('Second'), { left: 24, right: 0 }]]));
    await settle();
    expect(markerX('Second')).toBe(before.marker + 24);
    expect(doc.caretOf('Second').x).toBe(before.text + 24);
    expect(markerX('First')).toBe(before.marker);
  } finally {
    doc.close();
  }
});

test('top and bottom view room remain separate from authored paragraph spacing', async () => {
  const doc = open(docx(p('Alpha') + p('Beta') + p('Gamma')));
  try {
    const y = (text: string) => doc.caretOf(text).y;
    // Equal lines: each step is the line height plus the paragraph spacing.
    const step = y('Beta') - y('Alpha');
    expect(y('Gamma') - y('Beta')).toBeCloseTo(step, 6);
    doc.surface.setBlockInsets(
      new Map([[doc.idOf('Beta'), { left: 0, right: 0, top: 40, bottom: 10 }]])
    );
    await settle();
    // Host padding surrounds the unchanged Word gap and cannot collapse into that gap.
    const after = y('Beta') - y('Alpha');
    expect(after).toBeCloseTo(step + 40, 6);
    expect(y('Gamma') - y('Beta')).toBeCloseTo(step + 10, 6);
    expect(doc.caretOf('Beta').x).toBe(doc.caretOf('Alpha').x);
    doc.surface.setBlockInsets(new Map());
    await settle();
    expect(y('Beta') - y('Alpha')).toBeCloseTo(step, 6);
  } finally {
    doc.close();
  }
});
