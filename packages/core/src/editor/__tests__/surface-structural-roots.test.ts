import { expect, test } from 'bun:test';
import { docx, p } from '../../automation/__tests__/support/protocol.ts';
import { contentControlsIn, paragraphTextOf } from '../../store/index.ts';
import { createDocxEditor } from '../docx-editor.ts';

test('a borrowed summary stays in its host while native page siblings leave and return', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const bytes = docx(
    p('Before') +
      '<w:sdt><w:sdtPr/><w:sdtContent>' +
      p('Summary') +
      p('Answer') +
      '</w:sdtContent></w:sdt>' +
      p('After')
  );
  const editor = createDocxEditor({ container, document: bytes });
  const surface = editor.surface!;
  const part = surface.session.part();
  const owner = contentControlsIn(part.root)[0]!.node.id;
  const caption = surface.session
    .paragraphIds()
    .find((id) => paragraphTextOf(part, id) === 'Summary')!;
  const original = container.querySelector<HTMLElement>(
    `.docx-paragraph-fragment[data-paragraph-id="${caption}"]`
  )!;
  const marker = document.createComment('summary position');
  const view = document.createElement('section');
  const last = container.querySelectorAll<HTMLElement>('.docx-paragraph-fragment')[3]!;
  const endMarker = document.createComment('end position');
  const endView = document.createElement('aside');
  const homes = new Map([
    [original, marker],
    [last, endMarker],
  ]);
  const accepted: boolean[] = [];
  const stop = surface.observePaint({
    replayCurrentFrame: true,
    beforePaint(whole, preserve) {
      if (whole?.has(0)) return;
      const keep = preserve?.(
        0,
        (_changed, structural) => structural === true,
        (element) => homes.get(element) ?? element
      );
      if (preserve) accepted.push(keep === true);
      if (!keep && marker.parentNode) {
        marker.replaceWith(original);
        view.remove();
      }
    },
    afterPaint() {
      if (original.parentElement === view) return;
      original.before(marker, view);
      view.append(original);
    },
  });
  const stopEnd = surface.observePaint({
    replayCurrentFrame: true,
    beforePaint(_whole, reserve) {
      if (reserve)
        expect(
          reserve(
            0,
            () => false,
            (element) => homes.get(element) ?? element
          )
        ).toBe(false);
      if (endMarker.parentNode) {
        endMarker.replaceWith(last);
        endView.remove();
      }
    },
    afterPaint() {
      last.before(endMarker, endView);
      endView.append(last);
    },
  });
  try {
    for (const folded of [true, false]) {
      expect(
        surface.setCollapsedContentControls(folded ? new Map([[owner, caption]]) : new Map()).ok
      ).toBe(true);
      surface.flushPendingInput();
      expect(original.parentElement).toBe(view);
      expect(view.isConnected).toBe(true);
      expect(marker.parentNode).toBe(view.parentNode);
      expect(last.parentElement).toBe(endView);
      expect(endView.isConnected).toBe(true);
      expect(
        [...container.querySelectorAll('.docx-paragraph-fragment')].map((node) => node.textContent)
      ).toEqual(folded ? ['Before', 'Summary', 'After'] : ['Before', 'Summary', 'Answer', 'After']);
    }
    expect(accepted).toEqual([true, true]);
  } finally {
    stopEnd();
    stop();
    editor.destroy();
    container.remove();
  }
});

test.each([true, false])(
  'structural observers retain their own loans (first accepts: %s)',
  (firstAccepts) => {
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({
      container,
      document: docx(
        '<w:sdt><w:sdtPr/><w:sdtContent>' +
          p('Summary') +
          p('Answer') +
          '</w:sdtContent></w:sdt>' +
          p('After')
      ),
    });
    const surface = editor.surface!;
    const part = surface.session.part();
    const owner = contentControlsIn(part.root)[0]!.node.id;
    const summary = surface.session.paragraphIds()[0]!;
    const decisions: boolean[] = [];
    const callbacks: number[] = [];
    const stops = [firstAccepts, !firstAccepts].map((allow, index) =>
      surface.observePaint({
        replayCurrentFrame: true,
        beforePaint(_whole, reserve) {
          if (reserve)
            decisions.push(
              reserve(
                0,
                () => {
                  callbacks.push(index);
                  return allow;
                },
                (root) => root
              )
            );
        },
        afterPaint() {},
      })
    );
    try {
      surface.setCollapsedContentControls(new Map([[owner, summary]]));
      surface.flushPendingInput();
      expect(decisions).toEqual([firstAccepts, false]);
      expect(callbacks).toEqual(firstAccepts ? [0, 1] : [0]);
      expect(
        [...container.querySelectorAll('.docx-paragraph-fragment')].map((el) => el.textContent)
      ).toEqual(['Summary', 'After']);
    } finally {
      stops.forEach((stop) => stop());
      editor.destroy();
      container.remove();
    }
  }
);

test('a host can change paragraph children after accepting a translated paragraph', () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(
      '<w:sdt><w:sdtPr/><w:sdtContent>' +
        p('Summary') +
        p('Answer') +
        '</w:sdtContent></w:sdt>' +
        p('After')
    ),
  });
  const surface = editor.surface!;
  const part = surface.session.part();
  const owner = contentControlsIn(part.root)[0]!.node.id;
  const summary = surface.session.paragraphIds()[0]!;
  const root = container.querySelectorAll<HTMLElement>('.docx-paragraph-fragment')[2]!;
  const host = document.createElement('i');
  const previousTop = root.style.top;
  const stop = surface.observePaint({
    replayCurrentFrame: true,
    beforePaint(_whole, reserve) {
      reserve?.(
        0,
        (changed, structural) => {
          if (!structural || !changed.has(root)) return false;
          root.prepend(host);
          return true;
        },
        (node) => node
      );
    },
    afterPaint() {},
  });
  try {
    surface.setCollapsedContentControls(new Map([[owner, summary]]));
    surface.flushPendingInput();
    expect(host.parentNode).toBe(root);
    expect(root.isConnected).toBe(true);
    expect(root.textContent).toBe('After');
    expect(root.style.top).not.toBe(previousTop);
    expect(host.style.cssText).toBe('');
  } finally {
    stop();
    editor.destroy();
    container.remove();
  }
});
