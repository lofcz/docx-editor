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
