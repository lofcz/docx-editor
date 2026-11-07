import { expect, test } from 'bun:test';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { createContentControlCollapseState } from '../content-control-collapse-state.ts';
import { createLayoutSession } from '../layout-session.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';
import { anchor, drawings, filler, lineAt, text } from './anchor-page-fit-fixtures.ts';
import { layoutContext, load } from './anchored-drawing-test-fixtures.ts';

function documentWith(body: string) {
  const part = load(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
      ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<w:body>${body}</w:body></w:document>`
  );
  return {
    part,
    options: {
      compatibilityMode: 15,
      measurer: createFixedMeasurer(6, 14),
      inlineDrawingLayout: layoutContext(part),
    },
  };
}

test.each([
  { reason: 'the widget reservation', height: 40, reserve: 120, basePage: 0 },
  { reason: 'the wrapping object', height: 300, reserve: 40, basePage: 1 },
])(
  '$reason moves the paragraph with both view constraints intact',
  ({ height, reserve, basePage }) => {
    const { part, options } = documentWith(
      filler() +
        `<w:p>${anchor({ y: 8, height, x: 368 })}<w:r><w:t>Widget</w:t></w:r></w:p>` +
        text('After')
    );
    const fingerprint = canonicalOoxmlFingerprint(part);
    const xml = serializeOoxmlPart(part);
    const widgetId = bodyStoryRoot(part)!.children[1]!.id;
    const session = createLayoutSession();
    const plain = layoutSemanticDocument(part, 1, { ...options, session });
    expect(lineAt(plain, 'Widget')!.page).toBe(basePage);
    const blockReservations = {
      epoch: 'reserved',
      extraAfter: (id: string) => (id === widgetId ? reserve : 0),
    };
    const reserved = layoutSemanticDocument(part, 1, { ...options, session, blockReservations });
    expect(lineAt(reserved, 'Widget')).toEqual({ page: 1, y: 0 });
    expect(drawings(reserved)).toEqual([{ page: 1, y: 8 }]);
    const nativeGap = lineAt(plain, 'After')!.y - lineAt(plain, 'Widget')!.y;
    expect(lineAt(reserved, 'After')!.page).toBe(1);
    expect(lineAt(reserved, 'After')!.y).toBeCloseTo(reserve + nativeGap, 6);
    const fresh = layoutSemanticDocument(part, 1, { ...options, blockReservations });
    expect(reserved.pages).toEqual(fresh.pages);
    expect(layoutSemanticDocument(part, 1, { ...options, session }).pages).toEqual(plain.pages);
    expect(canonicalOoxmlFingerprint(part)).toEqual(fingerprint);
    expect(serializeOoxmlPart(part)).toBe(xml);
  }
);

test('collapse removes hidden wrapping objects from paginated flow and reopening restores them', () => {
  const { part, options } = documentWith(
    filler() +
      '<w:sdt><w:sdtPr/><w:sdtContent>' +
      text('Summary') +
      `<w:p>${anchor({ y: 8, height: 300, x: 368 })}<w:r><w:t>Hidden</w:t></w:r></w:p>` +
      '</w:sdtContent></w:sdt>' +
      text('After')
  );
  const fingerprint = canonicalOoxmlFingerprint(part);
  const xml = serializeOoxmlPart(part);
  const control = bodyStoryRoot(part)!.children[1]!;
  const summary = contentControlContentNodeOf(control)!.children[0]!;
  const collapse = createContentControlCollapseState();
  const session = createLayoutSession();
  const opened = layoutSemanticDocument(part, 1, { ...options, session });
  expect(drawings(opened)).toHaveLength(1);
  expect(collapse.set(part, new Map([[control.id, summary.id]])).ok).toBe(true);
  const closed = layoutSemanticDocument(part, 1, {
    ...options,
    session,
    contentControlCollapse: collapse.option(),
  });
  expect(drawings(closed)).toEqual([]);
  expect(lineAt(closed, 'Hidden')).toBeNull();
  expect(lineAt(closed, 'After')!.page).toBe(0);
  expect(closed.pages).toEqual(
    layoutSemanticDocument(part, 1, { ...options, contentControlCollapse: collapse.option() }).pages
  );
  expect(collapse.set(part, new Map()).ok).toBe(true);
  expect(
    layoutSemanticDocument(part, 1, {
      ...options,
      session,
      contentControlCollapse: collapse.option(),
    }).pages
  ).toEqual(opened.pages);
  expect(canonicalOoxmlFingerprint(part)).toEqual(fingerprint);
  expect(serializeOoxmlPart(part)).toBe(xml);
});
