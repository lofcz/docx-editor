import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createDocxEditor } from '../docx-editor.ts';
import { setPaginatedSurfaceScale } from '../surface-scale.ts';
import {
  setSurfaceInlineWidgetMetrics,
  SurfaceInlineWidgetMetrics,
} from '../surface-inline-widget-metrics.ts';
import { updateSurfaceMeasurement } from '../surface-measurement.ts';
import { docx } from '../../automation/__tests__/support/protocol.ts';
import { linesOf, type TextMeasurer } from '../../layout/semantic-records.ts';
import { createFixedMeasurer } from '../../layout/fixed-measurer.ts';
import { DEFAULT_RUN_STYLE } from '../../layout/run-style.ts';
import { canonicalOoxmlFingerprint, paragraphTextOf, readOoxmlPackage } from '../../store/index.ts';

const run = (text: string) => `<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>${text}</w:t></w:r>`;
const body = `<w:p>${run('A')}<w:sdt><w:sdtPr><w:tag w:val="chip:source"/></w:sdtPr><w:sdtContent>${run('source')}</w:sdtContent></w:sdt>${run('B')}</w:p>`;
const metrics = (width: number, revision: string, ascent = 20) => ({
  measure: (request: { tag: string | undefined }) =>
    request.tag === 'chip:source' ? { width, ascent, descent: 4 } : null,
  revision,
});

test('live metric publication retains native content, selection, one-step history, zoom, and font replacement', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docx(body),
  });
  try {
    const surface = editor.surface!,
      id = surface.session.paragraphIds()[0]!;
    const selection = {
      anchor: { paragraphId: id, offset: 8 },
      head: { paragraphId: id, offset: 8 },
    };
    surface.setSelection(selection);
    surface.type('!');
    const before = canonicalOoxmlFingerprint(surface.session.part().root),
      revision = surface.session.packageRevision();
    const selected = surface.state().selection;
    const widget = () =>
      linesOf(surface.layout())
        .flatMap((line) => line.spans)
        .find((span) => span.inlineWidget)!;
    setSurfaceInlineWidgetMetrics(surface, metrics(40, '1'));
    expect(widget().box.width).toBe(30);
    expect(surface.state().selection).toEqual(selected);
    expect(surface.session.packageRevision()).toBe(revision);
    expect(canonicalOoxmlFingerprint(surface.session.part().root)).toBe(before);
    expect(setPaginatedSurfaceScale(surface, 2)).toBe(true);
    expect(widget().box.width).toBe(30);
    updateSurfaceMeasurement(surface, {
      measurer: createFixedMeasurer(7, 16),
      producer: 'replacement-fonts',
    });
    expect(widget().box.width).toBe(30);
    setSurfaceInlineWidgetMetrics(surface, metrics(64, '2'));
    expect(widget().box.width).toBe(48);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(paragraphTextOf(surface.session.part(), id)).toBe('AsourceB');
    expect(widget().box.width).toBe(48);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(paragraphTextOf(surface.session.part(), id)).toBe('AsourceB!');
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    expect(
      canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!.root)
    ).toBe(canonicalOoxmlFingerprint(surface.session.part().root));
    setSurfaceInlineWidgetMetrics(surface, null);
    expect(
      linesOf(surface.layout())
        .flatMap((line) => line.spans)
        .some((span) => span.inlineWidget)
    ).toBe(false);
    expect(surface.textMeasurer!().measure('A', { ...DEFAULT_RUN_STYLE, fontSizePt: 11 })).toBe(7);
  } finally {
    editor.destroy();
  }
});

test('measurement wrappers preserve prototype method receivers and immutable prior snapshots', () => {
  class NativeMeasurer implements TextMeasurer {
    #advance = 7;
    measure(text: string) {
      return this.#advance * text.length;
    }
    lineMetrics() {
      return { height: this.#advance * 2, baseline: this.#advance };
    }
    caretAdvances(text: string) {
      return Array.from({ length: text.length + 1 }, (_, index) => this.#advance * index);
    }
  }
  const state = new SurfaceInlineWidgetMetrics(),
    base = new NativeMeasurer();
  expect(state.measurer(base)).toBe(base);
  state.set(metrics(40, '1'));
  const first = state.measurer(base);
  expect(first.measure('AB', DEFAULT_RUN_STYLE)).toBe(14);
  expect(first.lineMetrics(DEFAULT_RUN_STYLE)).toEqual({ height: 14, baseline: 7 });
  expect(first.caretAdvances!('AB', DEFAULT_RUN_STYLE)).toEqual([0, 7, 14]);
  const request = {
    controlId: 'id',
    tag: 'chip:source',
    fallbackText: 'source',
    availableWidth: 100,
  };
  expect(first.inlineWidgets!.measure(request)?.width).toBe(40);
  state.set(metrics(64, '2'));
  expect(state.measurer(base).inlineWidgets!.measure(request)?.width).toBe(64);
  expect(first.inlineWidgets!.measure(request)?.width).toBe(40);
  state.set(null);
  expect(state.measurer(base)).toBe(base);
});

test('Word save pagination ignores host-only inline boxes', async () => {
  const field =
    '<w:p><w:fldSimple w:instr="PAGEREF target"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>';
  const tail =
    '<w:p><w:bookmarkStart w:id="1" w:name="target"/><w:r><w:t>Tail</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>';
  const section =
    '<w:sectPr><w:pgSz w:w="4000" w:h="4000"/><w:pgMar w:top="200" w:right="200" w:bottom="200" w:left="200"/></w:sectPr>';
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docx(field + body + tail + section),
  });
  try {
    const before = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!before.ok) throw new Error(before.reason);
    expect(editor.surface!.layout().pages).toHaveLength(1);
    setSurfaceInlineWidgetMetrics(editor.surface!, metrics(40, '1', 400));
    expect(editor.surface!.layout().pages.length).toBeGreaterThan(1);
    const after = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!after.ok) throw new Error(after.reason);
    expect(
      canonicalOoxmlFingerprint(after.package.parts.get(after.package.mainDocumentPart)!.root)
    ).toBe(
      canonicalOoxmlFingerprint(before.package.parts.get(before.package.mainDocumentPart)!.root)
    );
  } finally {
    editor.destroy();
  }
});

test('header metric revisions invalidate cached furniture widths and retain native saved fallback', async () => {
  const entries = unzipSync(
    docx(
      '<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="header"/></w:sectPr>',
      undefined,
      {
        rels: '<Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      }
    )
  );
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
    )
  );
  entries['word/header1.xml'] = strToU8(
    `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${body}</w:hdr>`
  );
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: zipSync(entries),
  });
  try {
    const surface = editor.surface!;
    const headerWidget = () =>
      surface
        .layout()
        .pages[0]!.header!.fragments.flatMap((fragment) =>
          fragment.kind === 'paragraph' ? fragment.lines : []
        )
        .flatMap((line) => line.spans)
        .find((span) => span.inlineWidget);
    setSurfaceInlineWidgetMetrics(surface, metrics(40, '1'));
    expect(headerWidget()?.box.width).toBe(30);
    setSurfaceInlineWidgetMetrics(surface, metrics(64, '2'));
    expect(headerWidget()?.box.width).toBe(48);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    expect(canonicalOoxmlFingerprint(saved.package.parts.get('/word/header1.xml')!.root)).toBe(
      canonicalOoxmlFingerprint(
        surface.session.currentPackage().parts.get('/word/header1.xml')!.root
      )
    );
    setSurfaceInlineWidgetMetrics(surface, null);
    expect(headerWidget()).toBeUndefined();
  } finally {
    editor.destroy();
  }
});

test('measurement publication waits for native IME readback and applies the latest snapshot', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: docx(body) });
  try {
    const surface = editor.surface!,
      id = surface.session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId: id, offset: 1 },
      head: { paragraphId: id, offset: 1 },
    });
    surface.focus();
    await Promise.resolve();
    const page = container.querySelector<HTMLElement>('.docx-pages')!;
    const text = page.querySelector<HTMLElement>('.layout-run-text')!.firstChild!;
    const range = document.createRange();
    range.setStart(text, 1);
    range.collapse(true);
    const selection = document.defaultView!.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
    page.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }));
    text.textContent = 'AX';
    setSurfaceInlineWidgetMetrics(surface, metrics(40, '1'));
    setSurfaceInlineWidgetMetrics(surface, metrics(64, '2'));
    expect(text.isConnected).toBe(true);
    expect(text.textContent).toBe('AX');
    expect(paragraphTextOf(surface.session.part(), id)).toBe('AsourceB');
    page.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'X' }));
    await Promise.resolve();
    expect(paragraphTextOf(surface.session.part(), id)).toBe('AXsourceB');
    const widget = linesOf(surface.layout())
      .flatMap((line) => line.spans)
      .find((span) => span.inlineWidget)!;
    expect(widget.box.width).toBe(48);
  } finally {
    editor.destroy();
    container.remove();
  }
});
