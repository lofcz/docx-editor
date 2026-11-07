import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createDocxEditor } from '../docx-editor.ts';
import {
  docx,
  p,
  pWithSection,
  sdt,
  table,
  row,
  cell,
} from '../../automation/__tests__/support/protocol.ts';
import { caretAt } from '../../layout/semantic-interaction.ts';
import { paragraphFragmentsOf } from '../../layout/semantic-records.ts';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  parentNodeOf,
  readOoxmlPackage,
  semanticDigest,
} from '@docx-editor.dev/core/store';

function open(body: string, bytes = docx(body)) {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: bytes });
  const surface = editor.surface!;
  const idOf = (text: string) =>
    surface.session
      .paragraphIds()
      .find((id) => paragraphTextOf(surface.session.part(), id) === text)!;
  const ownerOf = (text: string) =>
    parentNodeOf(surface.session.part(), parentNodeOf(surface.session.part(), idOf(text))!.id)!.id;
  const caret = (text: string) => {
    surface.flushPendingInput();
    return caretAt(surface.layout(), { paragraphId: idOf(text), offset: 0 });
  };
  return {
    editor,
    surface,
    container,
    bytes,
    idOf,
    ownerOf,
    caret,
    close: () => {
      editor.destroy();
      container.remove();
    },
  };
}
async function digest(bytes: Uint8Array) {
  const read = readOoxmlPackage(bytes);
  if (!read.ok) throw Error(read.reason);
  return semanticDigest(read.package.parts.values());
}

test('collapse removes body flow while canonical content, history, and saved semantics stay unchanged', async () => {
  const doc = open(p('Before') + sdt(p('Summary'), p('Hidden body')) + p('After'));
  try {
    const before = canonicalOoxmlFingerprint(doc.surface.session.part());
    const saved = await digest(new Uint8Array(await doc.editor.save()));
    const revision = doc.surface.session.packageRevision();
    const afterY = doc.caret('After')!.y;
    const summaryY = doc.caret('Summary')!.y;
    const requests = new Map([[doc.ownerOf('Summary'), doc.idOf('Summary')]]);
    expect(doc.surface.setCollapsedContentControls(requests)).toEqual({ ok: true });
    expect(doc.caret('After')!.y).toBeLessThan(afterY);
    expect(doc.caret('Summary')!.y).toBe(summaryY);
    expect(doc.caret('Hidden body')).toBeNull();
    expect(
      doc.container.querySelector(`[data-paragraph-id="${doc.idOf('Hidden body')}"]`)
    ).toBeNull();
    const snapshot = doc.surface.collapsedContentControls();
    expect(Object.isFrozen(snapshot)).toBe(true);
    expect(doc.surface.setCollapsedContentControls(requests)).toEqual({ ok: true });
    expect(doc.surface.collapsedContentControls()).toBe(snapshot);
    expect(canonicalOoxmlFingerprint(doc.surface.session.part())).toBe(before);
    expect(doc.surface.session.packageRevision()).toBe(revision);
    expect(doc.surface.state().canUndo).toBe(false);
    expect(await digest(new Uint8Array(await doc.editor.save()))).toEqual(saved);
    expect(doc.surface.collapsedContentControls()).toBe(snapshot);
    expect(doc.surface.setCollapsedContentControls(new Map())).toEqual({ ok: true });
    expect(doc.caret('After')!.y).toBe(afterY);
  } finally {
    doc.close();
  }
});

test('a null summary replaces the whole answer with a local header and reveals before typing', async () => {
  const doc = open(p('Before') + sdt(p('Secret answer'), p('Hidden details')) + p('After'));
  try {
    const original = canonicalOoxmlFingerprint(doc.surface.session.part());
    const saved = await digest(new Uint8Array(await doc.editor.save()));
    const beforeY = doc.caret('After')!.y;
    const answer = doc.idOf('Secret answer');
    const owner = doc.ownerOf('Secret answer');
    const point = { paragraphId: answer, offset: 6 };
    doc.surface.setSelection({ anchor: point, head: point });
    expect(doc.surface.setCollapsedContentControls(new Map([[owner, null]]))).toEqual({ ok: true });
    doc.surface.flushPendingInput();
    const header = doc.surface
      .layout()
      .pages.flatMap(paragraphFragmentsOf)
      .find((p) => p.paragraphId === answer)!;
    expect(header.box.height).toBeCloseTo(24, 8);
    expect(header.lines).toEqual([]);
    expect(caretAt(doc.surface.layout(), { paragraphId: answer, offset: 0 })).toBeNull();
    expect(
      doc.container
        .querySelector('[data-docx-collapse-placeholder]')
        ?.getAttribute('contenteditable')
    ).toBe('false');
    expect(doc.caret('After')!.y).toBeLessThan(beforeY);
    expect(doc.container.textContent).not.toContain('Secret');
    expect(doc.container.textContent).not.toContain('Hidden details');
    expect(canonicalOoxmlFingerprint(doc.surface.session.part())).toBe(original);
    expect(doc.surface.state().canUndo).toBe(false);
    expect(await digest(new Uint8Array(await doc.editor.save()))).toEqual(saved);
    doc.surface.type('new ');
    doc.surface.flushPendingInput();
    expect(doc.surface.collapsedContentControls()).toEqual([]);
    expect(paragraphTextOf(doc.surface.session.part(), answer)).toBe('Secretnew  answer');
    expect(doc.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(doc.surface.session.part())).toBe(original);
  } finally {
    doc.close();
  }
});

for (const ranged of [false, true])
  test(`typing reveals its exact hidden destination before editing (range=${ranged})`, () => {
    const doc = open(sdt(p('Summary'), p('Body')) + p('After'));
    try {
      const anchor = { paragraphId: doc.idOf(ranged ? 'Summary' : 'Body'), offset: ranged ? 3 : 4 };
      const head = { paragraphId: doc.idOf('Body'), offset: ranged ? 2 : 4 };
      doc.surface.setSelection({ anchor, head });
      const bodyId = doc.idOf('Body');
      doc.surface.setCollapsedContentControls(
        new Map([[doc.ownerOf('Summary'), doc.idOf('Summary')]])
      );
      doc.surface.flushPendingInput();
      expect(caretAt(doc.surface.layout(), head)).toBeNull();
      expect(doc.surface.state().selection).toEqual({ anchor, head });
      let visibleAtCommit = false;
      const beforeRevision = doc.surface.session.packageRevision();
      const stop = doc.surface.observePaint({
        afterPaint: ({ layout }) => {
          if (doc.surface.session.packageRevision() === beforeRevision)
            visibleAtCommit =
              doc.surface.collapsedContentControls().length === 0 && caretAt(layout, head) !== null;
        },
      });
      doc.surface.type('X');
      stop();
      expect(visibleAtCommit).toBe(true);
      expect(doc.surface.collapsedContentControls()).toEqual([]);
      expect(
        doc.surface.session
          .paragraphIds()
          .map((id) => paragraphTextOf(doc.surface.session.part(), id))
      ).toEqual(ranged ? ['SumXdy', 'After'] : ['Summary', 'BodyX', 'After']);
      expect(doc.editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(paragraphTextOf(doc.surface.session.part(), bodyId)).toBe('Body');
      expect(doc.surface.collapsedContentControls()).toEqual([]);
    } finally {
      doc.close();
    }
  });

test('Enter after a collapsed summary reveals the new body caret before further input', () => {
  const doc = open(sdt(p('Summary'), p('Body')) + p('After'));
  try {
    const point = { paragraphId: doc.idOf('Summary'), offset: 7 };
    doc.surface.setSelection({ anchor: point, head: point });
    expect(
      doc.surface.setCollapsedContentControls(
        new Map([[doc.ownerOf('Summary'), point.paragraphId]])
      )
    ).toEqual({ ok: true });
    doc.surface.splitParagraph();
    doc.surface.flushPendingInput();
    expect(doc.surface.collapsedContentControls()).toEqual([]);
    expect(caretAt(doc.surface.layout(), doc.surface.state().selection.head)).not.toBeNull();
    doc.surface.type('New body');
    doc.surface.flushPendingInput();
    expect(
      doc.surface.session
        .paragraphIds()
        .map((id) => paragraphTextOf(doc.surface.session.part(), id))
    ).toEqual(['Summary', 'New body', 'Body', 'After']);
  } finally {
    doc.close();
  }
});

test('a replacement spanning a fold reveals hidden content with both endpoints outside it', () => {
  const doc = open(p('Before') + sdt(p('Summary'), p('Body')) + p('After'));
  try {
    const before = canonicalOoxmlFingerprint(doc.surface.session.part());
    const anchor = { paragraphId: doc.idOf('Before'), offset: 3 };
    const head = { paragraphId: doc.idOf('After'), offset: 2 };
    const body = { paragraphId: doc.idOf('Body'), offset: 0 };
    doc.surface.setSelection({ anchor, head });
    doc.surface.setCollapsedContentControls(
      new Map([[doc.ownerOf('Summary'), doc.idOf('Summary')]])
    );
    doc.surface.flushPendingInput();
    expect(caretAt(doc.surface.layout(), body)).toBeNull();
    const revision = doc.surface.session.packageRevision();
    let revealedSelection: unknown;
    const stop = doc.surface.observePaint({
      afterPaint: ({ layout }) => {
        if (doc.surface.session.packageRevision() === revision && caretAt(layout, body))
          revealedSelection = doc.surface.state().selection;
      },
    });
    doc.surface.type('X');
    stop();
    expect(revealedSelection).toEqual({ anchor, head });
    expect(doc.surface.collapsedContentControls()).toEqual([]);
    expect(
      doc.surface.session
        .paragraphIds()
        .map((id) => paragraphTextOf(doc.surface.session.part(), id))
    ).toEqual(['BefXter']);
    expect(doc.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(doc.surface.session.part())).toBe(before);
    expect(doc.surface.state().canUndo).toBe(false);
    expect(doc.editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(
      doc.surface.session
        .paragraphIds()
        .map((id) => paragraphTextOf(doc.surface.session.part(), id))
    ).toEqual(['BefXter']);
  } finally {
    doc.close();
  }
});

test('hidden footnote references reserve no note pages and save their complete canonical notes', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const source = unzipSync(
    docx(
      sdt(p('Summary'), '<w:p><w:r><w:t>Body</w:t><w:footnoteReference w:id="1"/></w:r></w:p>') +
        p('After') +
        '<w:sectPr><w:pgSz w:w="6000" w:h="2400"/><w:pgMar w:top="240" w:right="240" w:bottom="240" w:left="240"/></w:sectPr>',
      undefined,
      { rels: `<Relationship Id="rIdFootnote" Type="${R}/footnotes" Target="footnotes.xml"/>` }
    )
  );
  source['[Content_Types].xml'] = strToU8(
    strFromU8(source['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'
    )
  );
  source['word/footnotes.xml'] = strToU8(
    `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">` +
      Array.from({ length: 20 }, (_, i) => p(`Note paragraph ${i}`)).join('') +
      '</w:footnote></w:footnotes>'
  );
  const doc = open('', zipSync(source));
  try {
    const saved = await digest(new Uint8Array(await doc.editor.save()));
    const pageCount = doc.surface.layout().pages.length;
    expect(pageCount).toBeGreaterThan(1);
    expect(doc.surface.layout().pages.some((page) => (page.footnotes?.notes.length ?? 0) > 0)).toBe(
      true
    );
    doc.surface.setCollapsedContentControls(
      new Map([[doc.ownerOf('Summary'), doc.idOf('Summary')]])
    );
    doc.surface.flushPendingInput();
    expect(doc.surface.layout().pages).toHaveLength(1);
    expect(
      doc.surface.layout().pages.every((page) => (page.footnotes?.notes.length ?? 0) === 0)
    ).toBe(true);
    expect(await digest(new Uint8Array(await doc.editor.save()))).toEqual(saved);
    doc.surface.setCollapsedContentControls(new Map());
    doc.surface.flushPendingInput();
    expect(doc.surface.layout().pages).toHaveLength(pageCount);
  } finally {
    doc.close();
  }
});

test('invalid summaries and unsupported shapes refuse the whole view replacement', () => {
  const doc = open(
    sdt(p('Summary'), p('Body')) +
      sdt(p('Section summary'), pWithSection('Boundary')) +
      table(row(cell(sdt(p('Cell summary'), p('Cell body')))))
  );
  try {
    const owner = doc.ownerOf('Summary');
    expect(
      doc.surface.setCollapsedContentControls(new Map([[owner, doc.idOf('Summary')]]))
    ).toEqual({ ok: true });
    const snapshot = doc.surface.collapsedContentControls();
    for (const [control, summary, reason] of [
      [owner, doc.idOf('Body'), 'invalid-summary'],
      [doc.idOf('Summary'), doc.idOf('Summary'), 'invalid-owner'],
      [doc.ownerOf('Section summary'), doc.idOf('Section summary'), 'unsupported-content'],
    ] as const) {
      expect(doc.surface.setCollapsedContentControls(new Map([[control, summary]]))).toEqual({
        ok: false,
        reason,
      });
      expect(doc.surface.collapsedContentControls()).toBe(snapshot);
    }
  } finally {
    doc.close();
  }
});

test('folds inside table cells compact nested content and preserve exact saved source', async () => {
  const nested = table(row(cell(p('Nested body'))));
  const doc = open(table(row(cell(sdt(p('Cell summary'), p('Cell body'), nested)))) + p('After'));
  try {
    const before = canonicalOoxmlFingerprint(doc.surface.session.part());
    const saved = await digest(new Uint8Array(await doc.editor.save()));
    const openY = doc.caret('After')!.y;
    expect(
      doc.surface.setCollapsedContentControls(
        new Map([[doc.ownerOf('Cell summary'), doc.idOf('Cell summary')]])
      )
    ).toEqual({ ok: true });
    expect(doc.caret('Cell summary')).not.toBeNull();
    expect(doc.caret('Cell body')).toBeNull();
    expect(doc.caret('Nested body')).toBeNull();
    expect(doc.caret('After')!.y).toBeLessThan(openY);
    expect(canonicalOoxmlFingerprint(doc.surface.session.part())).toBe(before);
    expect(await digest(new Uint8Array(await doc.editor.save()))).toEqual(saved);
    doc.surface.setCollapsedContentControls(new Map());
    expect(doc.caret('After')!.y).toBe(openY);
    expect(doc.caret('Cell body')).not.toBeNull();
    expect(doc.caret('Nested body')).not.toBeNull();
  } finally {
    doc.close();
  }
});

test('explicit local navigation reveals nested owners before native caret placement', () => {
  const doc = open(
    sdt(p('Outer summary'), sdt(p('Inner summary'), p('Hidden body')), p('Outer tail')) + p('After')
  );
  try {
    const requests = new Map([
      [doc.ownerOf('Outer summary'), doc.idOf('Outer summary')],
      [doc.ownerOf('Inner summary'), doc.idOf('Inner summary')],
    ]);
    expect(doc.surface.setCollapsedContentControls(requests)).toEqual({ ok: true });
    expect(doc.caret('Hidden body')).toBeNull();
    const head = { paragraphId: doc.idOf('Hidden body'), offset: 2 };
    doc.surface.setSelection({ anchor: head, head });
    expect(doc.surface.collapsedContentControls()).toEqual([]);
    expect(caretAt(doc.surface.layout(), head)).not.toBeNull();
    expect(doc.surface.state().selection.head).toEqual(head);
    expect(doc.surface.state().canUndo).toBe(false);
  } finally {
    doc.close();
  }
});

test('collapse survives host moves and resets on document replacement', async () => {
  const doc = open(sdt(p('Summary'), p('Body')) + p('After'));
  const moved = document.createElement('div');
  document.body.append(moved);
  try {
    doc.surface.setCollapsedContentControls(
      new Map([[doc.ownerOf('Summary'), doc.idOf('Summary')]])
    );
    doc.surface.flushPendingInput();
    const snapshot = doc.surface.collapsedContentControls();
    doc.editor.detach();
    doc.editor.attach(moved);
    expect(doc.surface.collapsedContentControls()).toBe(snapshot);
    expect(doc.caret('Body')).toBeNull();
    await doc.editor.load(doc.bytes);
    expect(doc.editor.surface!.collapsedContentControls()).toEqual([]);
  } finally {
    doc.close();
    moved.remove();
  }
});

test('a reference after folded content still resolves its hidden canonical bookmark', () => {
  const target =
    '<w:p><w:bookmarkStart w:id="1" w:name="Target"/><w:r><w:t>Secret</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>';
  const reference = '<w:p><w:fldSimple w:instr=" REF Target \\h "/></w:p>';
  const doc = open(sdt(p('Summary'), target) + reference);
  try {
    const refId = doc.surface.session.paragraphIds().at(-1)!;
    const painted = () =>
      doc.container.querySelector(`[data-paragraph-id="${refId}"]`)!.textContent;
    expect(painted()).toContain('Secret');
    doc.surface.setCollapsedContentControls(
      new Map([[doc.ownerOf('Summary'), doc.idOf('Summary')]])
    );
    doc.surface.flushPendingInput();
    expect(painted()).toContain('Secret');
    expect(doc.caret('Secret')).toBeNull();
    const before = canonicalOoxmlFingerprint(doc.surface.session.part());
    doc.container
      .querySelector('a.docx-hyperlink')!
      .dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
    expect(doc.surface.collapsedContentControls()).toEqual([]);
    expect(doc.surface.state().selection.head).toMatchObject({
      paragraphId: doc.idOf('Secret'),
      offset: 0,
    });
    expect(doc.caret('Secret')).not.toBeNull();
    expect(canonicalOoxmlFingerprint(doc.surface.session.part())).toBe(before);
  } finally {
    doc.close();
  }
});
