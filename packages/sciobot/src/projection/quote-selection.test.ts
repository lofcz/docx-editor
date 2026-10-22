import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  escapeXml,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import reference from '../__tests__/fixtures/quote-reference.json';
import { docxFixture } from '../../test/support/document';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { BlockPayload } from '../blocks/payload-schema';
import { GENERATED_LABEL_TAG } from './block-tags';
import { projectMdx } from './mdx';
import { readEditorRevision } from './revision-cache';
import { readDocument } from '../tools/read';
import {
  resolveQuoteSelection,
  type NativeQuote,
  type QuotePoint,
  type QuoteSelection,
  type QuoteSelectionContext,
} from './quote-selection';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
const p = (text: string) => `<w:p>${run(text)}</w:p>`;
const control = (id: string, content: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const label = (text: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${GENERATED_LABEL_TAG}"/></w:sdtPr><w:sdtContent>${p(text)}</w:sdtContent></w:sdt>`;
const payload = (
  id: string,
  type: string,
  attributes: BlockPayload['attributes'] = {}
): BlockPayload => ({ id, type, version: 1, attributes });
const removed = '<w:del w:id="1" w:author="AI"><w:r><w:delText>Old</w:delText></w:r></w:del>';
const inserted = `<w:ins w:id="2" w:author="AI">${run('New')}</w:ins>`;

async function imported(body: string, payloads: readonly BlockPayload[] = [], bytes?: Uint8Array) {
  const parsed = readOoxmlPackage(bytes ?? docxFixture(body));
  if (!parsed.ok) throw new Error(parsed.reason);
  let pkg = parsed.package;
  for (const value of payloads) {
    const result = withPayloadRecord(pkg, value, 'Readable Word fallback');
    if (!result.ok) throw new Error(result.reason);
    pkg = result.pkg;
  }
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(pkg),
  });
  const session = editor.surface!.session;
  const snapshots = () =>
    [...session.currentPackage().parts].map(([name, part]) => [
      name,
      canonicalOoxmlFingerprint(part),
    ]);
  const before = snapshots();
  const selection = JSON.stringify(editor.surface!.state().selection);
  const revision = readEditorRevision(editor, 'quote-doc', 'mount-1')!;
  const opened = openDocumentForExport(session);
  if (!opened.ok) {
    editor.destroy();
    throw new Error(opened.reason);
  }
  try {
    const [workingLayout, liveLayout] = await Promise.all([
      opened.session.layoutFor('proposed'),
      opened.session.layoutFor('all-markup'),
    ]);
    const working = projectMdx(session.currentPackage(), workingLayout, revision);
    const live = projectMdx(session.currentPackage(), liveLayout, revision);
    if (!working.success) throw new Error(working.error);
    if (!live.success) throw new Error(live.error);
    const context: QuoteSelectionContext = {
      revision,
      part: session.part(),
      live: live.projection,
    };
    return {
      editor,
      working: working.projection,
      context,
      point(text: string, offset: number): QuotePoint {
        const id = session
          .paragraphIds()
          .find((id) => paragraphTextOf(session.part(), id) === text);
        if (!id) throw new Error(`Missing native paragraph ${text}`);
        return { paragraphId: id, offset };
      },
      quote(selection: QuoteSelection, overrides: Partial<QuoteSelectionContext> = {}) {
        return resolveQuoteSelection(working.projection, selection, { ...context, ...overrides });
      },
      dispose() {
        try {
          expect(snapshots()).toEqual(before);
          expect(session.packageRevision()).toBe(revision.revision);
          expect(JSON.stringify(editor.surface!.state().selection)).toBe(selection);
          expect([session.canUndo(), session.canRedo()]).toEqual([false, false]);
        } finally {
          opened.session.dispose();
          editor.destroy();
        }
      },
    };
  } catch (error) {
    opened.session.dispose();
    editor.destroy();
    throw error;
  }
}
const expected = (
  text: string,
  startLine = 1,
  endLine = startLine,
  containsMdx = false
): NativeQuote => ({ text, startLine, endLine, containsMdx, resolved: true });
const sourceCase = (name: string) => reference.cases.find((item) => item.name === name)!.expected;
function assertWorking(
  value: Awaited<ReturnType<typeof imported>>,
  selection: QuoteSelection,
  quote: NativeQuote
) {
  const result = value.quote(selection);
  if (!result.success) throw new Error(result.error);
  expect(result.view).toBe('working');
  expect(result.quote).toEqual(quote);
  expect(result.revision).toEqual(value.context.revision);
  expect(result.revision === value.context.revision).toBe(false);
  expect(Object.isFrozen(result.quote)).toBe(true);
  expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  if (quote.resolved)
    expect(
      readDocument(value.working, {
        offset: quote.startLine,
        limit: quote.endLine - quote.startLine + 1,
      }).content
    ).toBe(quote.text);
  return result;
}

test('partial native text and heading selections match the source quote resolver', async () => {
  const value = await imported(
    p('Hello world') + '<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr>' + run('Title') + '</w:p>'
  );
  try {
    assertWorking(
      value,
      { anchor: value.point('Hello world', 2), head: value.point('Hello world', 7) },
      sourceCase('single paragraph partial selection')
    );
    assertWorking(
      value,
      { anchor: value.point('Title', 0), head: value.point('Title', 5) },
      { ...sourceCase('heading'), startLine: 3, endLine: 3 }
    );
  } finally {
    value.dispose();
  }
});

test('backward cross-paragraph selections retain complete source lines', async () => {
  const value = await imported(p('Alpha') + p('Beta') + p('Tail'));
  try {
    const forward = { anchor: value.point('Alpha', 2), head: value.point('Beta', 3) };
    assertWorking(
      value,
      { anchor: forward.head, head: forward.anchor },
      sourceCase('backward cross paragraph')
    );
    assertWorking(value, forward, sourceCase('backward cross paragraph'));
  } finally {
    value.dispose();
  }
});

test('native escaping and Unicode quote the landed read_document text', async () => {
  const value = await imported(p('A😀&*B'));
  try {
    expect(sourceCase('escaped emoji').text).toBe('A😀&\\*B');
    assertWorking(
      value,
      { anchor: value.point('A😀&*B', 1), head: value.point('A😀&*B', 3) },
      expected('A😀&amp;\\*B')
    );
    const result = value.quote({
      anchor: value.point('A😀&*B', 2),
      head: value.point('A😀&*B', 3),
    });
    expect(!result.success && result.code).toBe('invalid-input');
  } finally {
    value.dispose();
  }
});

test('manual-break selection narrows to its actual Markdown line', async () => {
  const value = await imported(
    `<w:p>${run('First')}<w:r><w:br/></w:r>${run('Second')}<w:r><w:br/></w:r>${run('Third')}</w:p>`
  );
  try {
    expect(sourceCase('multiline paragraph')).toEqual(expected('Second\\', 2));
    assertWorking(
      value,
      {
        anchor: value.point('First\nSecond\nThird', 6),
        head: value.point('First\nSecond\nThird', 12),
      },
      expected('Second  ', 2)
    );
  } finally {
    value.dispose();
  }
});

test('pending replacement text quotes the working result exactly once', async () => {
  const value = await imported(`<w:p>${run('A')}${removed}${inserted}${run('Z')}</w:p>`);
  try {
    assertWorking(
      value,
      { anchor: value.point('AOldNewZ', 4), head: value.point('AOldNewZ', 7) },
      sourceCase('new pending text')
    );
    expect(collectRevisionSites(value.editor.surface!.session.part())).toHaveLength(2);
  } finally {
    value.dispose();
  }
});

test('pending block insertion after a deleted paragraph mark uses its working source identity', async () => {
  const value = await imported(
    '<w:p><w:pPr><w:rPr><w:del w:id="10" w:author="AI"/></w:rPr></w:pPr>' +
      removed +
      '</w:p>' +
      '<w:p><w:pPr><w:rPr><w:ins w:id="11" w:author="AI"/></w:rPr></w:pPr>' +
      inserted +
      '</w:p>' +
      p('Tail')
  );
  try {
    assertWorking(
      value,
      { anchor: value.point('New', 0), head: value.point('New', 3) },
      sourceCase('new pending block')
    );
    const fallback = value.quote({ anchor: value.point('Old', 1), head: value.point('Old', 2) });
    if (!fallback.success) throw new Error(fallback.error);
    expect(fallback.view).toBe('live');
    expect(sourceCase('removed pending block fallback').text).toBe('Old');
    expect(fallback.quote).toEqual(expected('~~Old~~'));
  } finally {
    value.dispose();
  }
});

test('selected deleted text uses an explicit live-view fallback, including boundary-only selections', async () => {
  const value = await imported(`<w:p>${run('A')}${removed}${inserted}${run('Z')}</w:p>`);
  try {
    // The source fallback can resolve a removed leaf to empty text. Native reads retain its visible deletion.
    expect(sourceCase('removed pending text fallback')).toEqual(expected(''));
    for (const [start, end] of [
      [1, 4],
      [2, 3],
      [3, 2],
    ]) {
      const result = value.quote({
        anchor: value.point('AOldNewZ', start!),
        head: value.point('AOldNewZ', end!),
      });
      if (!result.success) throw new Error(result.error);
      expect(result.view).toBe('live');
      expect(result.quote).toEqual(expected('A~~Old~~NewZ')); // Native all-markup reads retain deletion markup.
    }
    const result = value.quote(
      { anchor: value.point('AOldNewZ', 2), head: value.point('AOldNewZ', 3) },
      { live: undefined }
    );
    expect(result.success && result.quote.resolved).toBe(false);
    expect(result.success && result.quote.text).toBe('');
  } finally {
    value.dispose();
  }
});

test('a nested activity body quotes its line without expanding its ancestors', async () => {
  const value = await imported(
    control('phase', control('activity', p('Question') + p('Answer'))) + p('Tail'),
    [
      payload('phase', 'lesson_phase', { phaseType: 'evocation' }),
      payload('activity', 'lesson_activity', { name: 'Warmup', duration: '5' }),
    ]
  );
  try {
    expect(sourceCase('nested activity body').text.trim()).toBe('Answer');
    expect(value.working.text).toBe(
      '<phase type="evocation">\n\n<activity name="Warmup" duration="5">\n\nQuestion\n\nAnswer\n\n</activity>\n\n</phase>\n\nTail'
    );
    assertWorking(
      value,
      { anchor: value.point('Answer', 1), head: value.point('Answer', 5) },
      expected('Answer', 7)
    );
  } finally {
    value.dispose();
  }
});

test('explicit block edges and generated native labels quote the full activity', async () => {
  const value = await imported(
    control('activity', label('Activity heading') + p('Question') + p('Answer')) + p('Tail'),
    [payload('activity', 'lesson_activity', { name: 'Warmup', duration: '5' })]
  );
  try {
    const quote = expected(
      '<activity name="Warmup" duration="5">\n\nQuestion\n\nAnswer\n\n</activity>',
      1,
      7,
      true
    );
    expect(sourceCase('nested activity container').containsMdx).toBe(true);
    assertWorking(
      value,
      {
        anchor: { blockId: 'activity', edge: 'start' },
        head: { blockId: 'activity', edge: 'end' },
      },
      quote
    );
    assertWorking(
      value,
      { anchor: value.point('Activity heading', 0), head: value.point('Activity heading', 16) },
      quote
    );
  } finally {
    value.dispose();
  }
});

test('table cell selections match source line extraction without selecting another row', async () => {
  const cell = (text: string) => `<w:tc>${p(text)}</w:tc>`;
  const row = (a: string, b: string) => `<w:tr>${cell(a)}${cell(b)}</w:tr>`;
  const value = await imported(
    '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      row('A', 'B') +
      row('One', 'Two') +
      '</w:tbl>' +
      p('Tail')
  );
  try {
    assertWorking(
      value,
      { anchor: value.point('Two', 0), head: value.point('Two', 3) },
      sourceCase('table cell')
    );
  } finally {
    value.dispose();
  }
});

test('inline atomic source fallback and adjacent text retain MDX quote syntax', async () => {
  const value = await imported(
    '<w:p>' + run('Before ') + control('citation', run('Citation')) + run(' after') + '</w:p>',
    [payload('citation', 'source', { url: 'https://example.com', title: 'Citation' })]
  );
  try {
    const quote = sourceCase('inline source');
    assertWorking(
      value,
      {
        anchor: value.point('Before Citation after', 1),
        head: value.point('Before Citation after', 17),
      },
      quote
    );
    assertWorking(
      value,
      {
        anchor: value.point('Before Citation after', 8),
        head: value.point('Before Citation after', 12),
      },
      quote
    );
  } finally {
    value.dispose();
  }
});

test('one stale endpoint uses mapped context; two unmapped endpoints stay unresolved', async () => {
  const value = await imported(p('Actual'));
  try {
    assertWorking(
      value,
      { anchor: value.point('Actual', 0), head: { paragraphId: 'gone', offset: 2 } },
      sourceCase('one endpoint maps')
    );
    assertWorking(
      value,
      { anchor: { paragraphId: 'gone', offset: 0 }, head: { paragraphId: 'also-gone', offset: 2 } },
      sourceCase('neither endpoint maps')
    );
  } finally {
    value.dispose();
  }
});

test('stale mount/revision, invalid ranges, wrong views, and getter inputs never produce quotes', async () => {
  const value = await imported(p('Actual'));
  try {
    const selection = { anchor: value.point('Actual', 0), head: value.point('Actual', 6) };
    for (const patch of [
      { documentId: 'other' },
      { epoch: 'other' },
      { revision: value.context.revision.revision + 1 },
    ]) {
      const result = value.quote(selection, { revision: { ...value.context.revision, ...patch } });
      expect(!result.success && result.code).toBe('stale-revision');
    }
    for (const offset of [-1, 1.5, Infinity, NaN]) {
      const result = value.quote({ ...selection, head: { paragraphId: 'missing', offset } });
      expect(!result.success && result.code).toBe('invalid-input');
    }
    expect(
      value.quote({ anchor: value.point('Actual', 0), head: value.point('Actual', 7) }).success
    ).toBe(false);
    let calls = 0;
    const getter = {
      get anchor() {
        calls++;
        return selection.anchor;
      },
      head: selection.head,
    };
    expect(value.quote(getter).success).toBe(false);
    expect(calls).toBe(0);
    const wrongMode = resolveQuoteSelection(value.context.live!, selection, value.context);
    expect(!wrongMode.success && wrongMode.code).toBe('invalid-document');
    const staleLive = value.quote(selection, {
      live: { ...value.context.live!, revision: { ...value.context.revision, epoch: 'old' } },
    });
    expect(!staleLive.success && staleLive.code).toBe('stale-revision');
  } finally {
    value.dispose();
  }
});

test('quote reads preserve an existing undo entry and pending review state through save/reopen', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(`<w:p>${run('A')}${removed}${inserted}${run('Z')}</w:p>`),
  });
  const surface = editor.surface!,
    session = surface.session;
  const id = session.paragraphIds()[0]!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: 0 },
    head: { paragraphId: id, offset: 0 },
  });
  surface.insertPlainText('User ');
  const revision = readEditorRevision(editor, 'quote-doc', 'mount-1')!;
  const opened = openDocumentForExport(session, { displayMode: 'proposed' });
  if (!opened.ok) {
    editor.destroy();
    throw new Error(opened.reason);
  }
  try {
    const layout = await opened.session.layout();
    const working = projectMdx(session.currentPackage(), layout, revision);
    if (!working.success) throw new Error(working.error);
    const before = canonicalOoxmlFingerprint(session.part());
    const pending = collectRevisionSites(session.part()).length;
    const selection = JSON.stringify(surface.state().selection);
    for (let read = 0; read < 3; read++) {
      const result = resolveQuoteSelection(
        working.projection,
        { anchor: { paragraphId: id, offset: 9 }, head: { paragraphId: id, offset: 12 } },
        { revision, part: session.part() }
      );
      expect(result.success && result.quote).toEqual(expected('User ANewZ'));
    }
    expect(session.packageRevision()).toBe(revision.revision);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
    expect(collectRevisionSites(session.part())).toHaveLength(pending);
    expect(JSON.stringify(surface.state().selection)).toBe(selection);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), id)).toBe('AOldNewZ');
    expect(session.canUndo()).toBe(false);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(paragraphTextOf(session.part(), id)).toBe('User AOldNewZ');
    const reopened = await imported('', [], new Uint8Array(await editor.save()));
    try {
      assertWorking(
        reopened,
        { anchor: reopened.point('User AOldNewZ', 9), head: reopened.point('User AOldNewZ', 12) },
        expected('User ANewZ')
      );
      expect(collectRevisionSites(reopened.editor.surface!.session.part())).toHaveLength(pending);
    } finally {
      reopened.dispose();
    }
  } finally {
    opened.session.dispose();
    editor.destroy();
  }
});
