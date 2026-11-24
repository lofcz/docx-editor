import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import type { ConfigureInitialBlockView } from '../surface-content-control-collapse.ts';
import { paragraphIdsInDocumentOrder } from '../../store/store/tree-op-blocks.ts';
import { paragraphTextOf, parentNodeOf } from '@docx-editor.dev/core/store';
import { docx, p, sdt } from '../../automation/__tests__/support/protocol.ts';
import { caretAt } from '../../layout/semantic-interaction.ts';

const bytes = docx(p('First') + sdt(p('Summary'), p('Answer')) + p('Last'));

test('the first layout applies the validated folds and insets from its own parsed package', async () => {
  const original = document.createElement('div');
  const moved = document.createElement('div');
  document.body.append(original, moved);
  let calls = 0;
  let retained: Parameters<ConfigureInitialBlockView>[1] | undefined;
  let owner = '',
    summary = '',
    answer = '';
  const editor = createDocxEditor({
    document: bytes,
    configureInitialBlockView(pkg, view) {
      calls++;
      retained = view;
      const part = pkg.parts.get(pkg.mainDocumentPart)!;
      const ids = paragraphIdsInDocumentOrder(part);
      summary = ids.find((id) => paragraphTextOf(part, id) === 'Summary')!;
      answer = ids.find((id) => paragraphTextOf(part, id) === 'Answer')!;
      owner = parentNodeOf(part, parentNodeOf(part, summary)!.id)!.id;
      expect(view.setCollapsedContentControls(new Map([[owner, summary]]))).toEqual({ ok: true });
      const accepted = view.collapsedContentControls();
      expect(view.setCollapsedContentControls(new Map([[owner, answer]]))).toEqual({
        ok: false,
        reason: 'invalid-summary',
      });
      expect(view.collapsedContentControls()).toBe(accepted);
      view.setBlockInsets(new Map([[summary, { left: 12, right: 3 }]]));
    },
  });
  try {
    expect(calls).toBe(0);
    editor.attach(original);
    const surface = editor.surface!;
    const position = (id: string) => caretAt(surface.layout(), { paragraphId: id, offset: 0 });
    const first = surface.session.paragraphIds()[0]!;
    const firstLayout = surface.layout();
    expect(surface.state().perf.fullPasses).toBe(1);
    expect(position(answer)).toBeNull();
    expect(position(summary)!.x - position(first)!.x).toBeCloseTo(12);
    expect(original.querySelector(`[data-paragraph-id="${answer}"]`)).toBeNull();
    const { setBlockInsets, setCollapsedContentControls } = retained!;
    expect(() => setBlockInsets(new Map())).toThrow('closed');
    expect(() => setCollapsedContentControls(new Map())).toThrow('closed');
    expect(surface.setCollapsedContentControls(new Map([[owner, summary]]))).toEqual({ ok: true });
    surface.setBlockInsets(new Map([[summary, { left: 12, right: 3 }]]));
    surface.flushPendingInput();
    expect(surface.layout()).toBe(firstLayout);
    expect(surface.state().canUndo).toBe(false);
    editor.detach();
    editor.attach(moved);
    expect(calls).toBe(1);
    expect(position(answer)).toBeNull();
    await editor.load(bytes);
    expect(calls).toBe(2);
    expect(editor.surface!.collapsedContentControls()).toHaveLength(1);
  } finally {
    editor.destroy();
    original.remove();
    moved.remove();
  }
});

test('an unknown initial owner refuses the map and leaves the document visible', () => {
  const container = document.createElement('div');
  const editor = createDocxEditor({
    container,
    document: bytes,
    configureInitialBlockView(_pkg, view) {
      expect(view.setCollapsedContentControls(new Map([['unknown', null]]))).toEqual({
        ok: false,
        reason: 'invalid-owner',
      });
      expect(view.collapsedContentControls()).toEqual([]);
    },
  });
  try {
    expect(container.textContent).toContain('Answer');
    expect(editor.surface!.state().perf.fullPasses).toBe(1);
  } finally {
    editor.destroy();
  }
});

test('a width-based widget takes its measured room on the first layout', () => {
  const host = document.createElement('div');
  const control = createDocxEditor({ container: document.createElement('div'), document: bytes });
  let width = 0;
  let id = '';
  const editor = createDocxEditor({
    container: host,
    document: bytes,
    configureInitialBlockView(pkg, view) {
      id = paragraphIdsInDocumentOrder(pkg.parts.get(pkg.mainDocumentPart)!)[0]!;
      view.setInitialBlockHeights(
        new Map([
          [
            id,
            (points) => {
              width = points;
              return points / 3;
            },
          ],
        ])
      );
    },
  });
  try {
    const surface = editor.surface!;
    const first = surface.layout();
    const nextId = surface.session.paragraphIds()[1]!;
    const y = (layout: typeof first) => caretAt(layout, { paragraphId: nextId, offset: 0 })!.y;
    const extra = y(first) - y(control.surface!.layout());
    expect(width).toBeGreaterThan(0);
    expect(extra).toBeGreaterThan(0);
    expect(extra).toBeLessThan(width / 3);
    expect(surface.state().perf.fullPasses).toBe(1);
    surface.setBlockReservations(new Map([[id, extra]]));
    surface.flushPendingInput();
    expect(surface.layout()).toBe(first);
    surface.setBlockReservations(new Map([[id, 0]]));
    surface.flushPendingInput();
    expect(y(surface.layout())).toBeCloseTo(y(control.surface!.layout()));
    expect(surface.state().canUndo).toBe(false);
  } finally {
    editor.destroy();
    control.destroy();
  }
});

test.each([false, true])(
  'an unmeasured opening box follows its owner after an edit (owner check: %s)',
  (checked) => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docx(p('Outside') + p('Figure') + p('After')),
      configureInitialBlockView(pkg, view) {
        const [, figure] = paragraphIdsInDocumentOrder(pkg.parts.get(pkg.mainDocumentPart)!);
        view.setInitialBlockHeights(
          new Map([[figure!, () => 150]]),
          checked
            ? (current, id) =>
                paragraphTextOf(current.parts.get(current.mainDocumentPart)!, id) === 'Figure'
            : undefined
        );
      },
    });
    try {
      const surface = editor.surface!;
      const [outside, figure, after] = surface.session.paragraphIds();
      const position = () => caretAt(surface.layout(), { paragraphId: after!, offset: 0 })!.y;
      const opening = position();
      const edit = (id: string) => {
        const caret = { paragraphId: id, offset: 0 };
        surface.setSelection({ anchor: caret, head: caret });
        surface.type('x');
        surface.flushPendingInput();
      };
      edit(outside!);
      // A later measurement can reflow an offscreen paragraph after the edit.
      surface.setBlockInsets(new Map([[figure!, { left: 3, right: 1 }]]));
      surface.flushPendingInput();
      if (checked) expect(position()).toBeCloseTo(opening);
      else expect(position()).toBeLessThan(opening - 100);
      edit(figure!);
      expect(position()).toBeLessThan(opening - 100);
    } finally {
      editor.destroy();
    }
  }
);

test('opening boxes use section column widths in points independently of zoom', () => {
  for (const [pageWidth, zoom] of [
    [7800, 0.75],
    [12240, 1.5],
  ] as const) {
    const available = (pageWidth - 1000) / 20;
    const widthCalls: number[] = [];
    const word = docx(
      p('Figure') +
        p('After') +
        `<w:sectPr><w:pgSz w:w="${pageWidth}" w:h="15840"/><w:pgMar w:left="500" w:right="500"/></w:sectPr>`
    );
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: word,
      zoom,
      configureInitialBlockView(pkg, view) {
        const [id] = paragraphIdsInDocumentOrder(pkg.parts.get(pkg.mainDocumentPart)!);
        view.setInitialBlockHeights(
          new Map([
            [
              id!,
              (widthPt) => {
                widthCalls.push(widthPt);
                return (widthPt * 2) / 3;
              },
            ],
          ])
        );
      },
    });
    try {
      expect(widthCalls[0]).toBeCloseTo(available);
      expect(editor.getZoom()).toBe(zoom);
      expect(editor.surface!.state().perf.fullPasses).toBe(1);
    } finally {
      editor.destroy();
    }
  }
});
