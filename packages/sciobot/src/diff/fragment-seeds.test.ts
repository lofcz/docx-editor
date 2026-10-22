import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  directParagraphProperties,
  findNode,
  paragraphTextOf,
  storyParagraphs,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { W } from '../../test/support/document';
import { listDocument } from '../../test/support/list-document';
import { targetParagraph as p } from '../../test/support/target-document';
import { parseTargetMarkdown } from '../projection/parse-target';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from '../ai/apply-direct';
import { decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';
import { planTargetDocument } from './plan-target';

const STYLES =
  `<w:styles xmlns:w="${W}">` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/></w:style>' +
  '</w:styles>';
const context = { documentId: 'fragments', epoch: 'mount' };
const nameOf = (node: OoxmlNode) => (node.kind === 'textValue' ? '' : node.localName);
/** Native Markdown always separates list items and table cells loosely. */
function ast(markdown: string) {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  return JSON.parse(
    JSON.stringify(result.target.normalizedAst, (key, value) =>
      key === 'spread' ? undefined : value
    )
  ) as unknown;
}
function mount(body: string, styles: string | null = STYLES) {
  const container = document.createElement('div');
  document.body.append(container);
  return createDocxEditor({
    container,
    document: listDocument(
      body +
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440"/></w:sectPr>',
      styles ?? undefined
    ),
    author: 'Teacher',
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
}
async function source(editor: DocxEditorInstance) {
  const token = readEditorRevision(editor, context.documentId, context.epoch)!,
    view = await readWorkingView(editor.surface!.session, token);
  if (!view.success) throw new Error(view.error);
  return {
    revision: token,
    pkg: editor.surface!.session.currentPackage(),
    projection: view.view.projection,
  };
}
async function plan(editor: DocxEditorInstance, target: string): Promise<CanonicalEditPlan> {
  const result = planTargetDocument(await source(editor), target);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
const body = (editor: DocxEditorInstance) => {
  const root = bodyStoryRoot(editor.surface!.session.part())!;
  if (root.kind === 'textValue') throw new Error('Invalid body');
  return root.children.filter((node) => node.kind !== 'textValue' && node.localName !== 'sectPr');
};
const cellTexts = (editor: DocxEditorInstance, table: OoxmlNode) => {
  const part = editor.surface!.session.part();
  if (table.kind === 'textValue') return [];
  return table.children
    .filter((row) => nameOf(row) === 'tr')
    .map((row) =>
      row.kind === 'textValue'
        ? []
        : row.children
            .filter((cell) => nameOf(cell) === 'tc')
            .map((cell) =>
              storyParagraphs(cell)
                .map((node) => paragraphTextOf(part, node.id))
                .join('|')
            )
    );
};
/** Apply directly, then prove one undo step, redo, and exact reopen against the target. */
async function land(editor: DocxEditorInstance, target: string, planned: CanonicalEditPlan) {
  const result = applyDirectPlan(editor, planned, context);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  expect(ast((await source(editor)).projection.text)).toEqual(ast(target));
  expect((await plan(editor, target)).operations).toEqual([]);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(editor.surface!.session.canUndo()).toBe(false);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  const container = document.createElement('div');
  document.body.append(container);
  const reopened = createDocxEditor({ container, document: new Uint8Array(await editor.save()) });
  try {
    expect(ast((await source(reopened)).projection.text)).toEqual(ast(target));
  } finally {
    reopened.destroy();
    container.remove();
  }
}

test('a new GFM table with alignment, marks, and links lands as one native Word table', async () => {
  const editor = mount(p('A.') + p('B.'));
  try {
    const [a, b] = body(editor),
      target =
        'A.\n\n| Name | Value |\n| :---: | ---: |\n| **One** | [two](https://example.org) |\n| three | *4* |\n\nB.',
      planned = await plan(editor, target);
    expect(planned.operations).toEqual([
      expect.objectContaining({
        kind: 'siblings',
        beforeId: a!.id,
        afterId: b!.id,
        removedIds: [],
        inserted: [expect.objectContaining({ type: 'table', align: ['center', 'right'] })],
        lists: [],
      }),
    ]);
    expect(decodeEditPlan(serializeEditPlan(planned))).toEqual({ success: true, plan: planned });
    await land(editor, target, planned);
    const blocks = body(editor);
    expect(blocks.map((node) => (node.kind === 'textValue' ? '' : node.localName))).toEqual([
      'p',
      'tbl',
      'p',
    ]);
    expect(blocks[0]!.id).toBe(a!.id);
    expect(blocks[2]!.id).toBe(b!.id);
    expect(cellTexts(editor, blocks[1]!)).toEqual([
      ['Name', 'Value'],
      ['One', 'two'],
      ['three', '4'],
    ]);
    const part = editor.surface!.session.part();
    const firstRow =
      blocks[1]!.kind === 'textValue'
        ? []
        : blocks[1]!.children.filter((row) => nameOf(row) === 'tr');
    const alignments = firstRow.flatMap((row) =>
      row.kind === 'textValue'
        ? []
        : row.children
            .filter((cell) => nameOf(cell) === 'tc')
            .map(
              (cell) =>
                directParagraphProperties(part, storyParagraphs(cell)[0]!.id).find(
                  (item) => item.localName === 'jc'
                )?.attributes?.val
            )
    );
    expect(alignments).toEqual(['center', 'right', 'center', 'right', 'center', 'right']);
  } finally {
    editor.destroy();
  }
}, 60_000);

test('a trailing table and adjacent tables receive the paragraphs Word requires', async () => {
  const editor = mount(p('A.'));
  try {
    const target = 'A.\n\n| X |\n| --- |\n| 1 |\n\n| Y |\n| --- |\n| 2 |';
    await land(editor, target, await plan(editor, target));
    expect(body(editor).map((node) => (node.kind === 'textValue' ? '' : node.localName))).toEqual([
      'p',
      'tbl',
      'p',
      'tbl',
      'p',
    ]);
    const part = editor.surface!.session.part();
    expect(
      body(editor)
        .filter((node) => node.kind === 'paragraph')
        .map((node) => paragraphTextOf(part, node.id))
    ).toEqual(['A.', '', '']);
  } finally {
    editor.destroy();
  }
}, 60_000);

test('a new block quote uses the document Quote style and native paragraphs', async () => {
  const editor = mount(p('Intro.') + p('Outro.'));
  try {
    const target = 'Intro.\n\n> First *quoted* [line](https://example.org).\n\nOutro.',
      planned = await plan(editor, target);
    expect(planned.operations).toEqual([
      expect.objectContaining({
        kind: 'siblings',
        inserted: [expect.objectContaining({ type: 'blockquote' })],
        styles: [{ path: [0], styleId: 'Quote' }],
      }),
    ]);
    await land(editor, target, planned);
    const part = editor.surface!.session.part(),
      paragraphs = body(editor);
    expect(paragraphs.map((node) => paragraphTextOf(part, node.id))).toEqual([
      'Intro.',
      'First quoted line.',
      'Outro.',
    ]);
    for (const node of paragraphs.slice(1, 2))
      expect(
        directParagraphProperties(part, node.id).find((item) => item.localName === 'pStyle')
          ?.attributes
      ).toEqual({
        val: 'Quote',
      });
    expect(findNode(part, paragraphs[0]!.id)).toBeTruthy();
  } finally {
    editor.destroy();
  }
}, 60_000);

test('a new multi-paragraph block quote lands as consecutive Quote paragraphs', async () => {
  const editor = mount(p('Intro.') + p('Outro.'));
  try {
    const target =
        'Intro.\n\n> First *quoted*.\n>\n> Second [line](https://example.org).\n\nOutro.',
      planned = await plan(editor, target);
    expect(planned.operations).toEqual([
      expect.objectContaining({
        kind: 'siblings',
        inserted: [expect.objectContaining({ type: 'blockquote' })],
        styles: [{ path: [0], styleId: 'Quote' }],
      }),
    ]);
    await land(editor, target, planned);
    const part = editor.surface!.session.part(),
      paragraphs = body(editor);
    expect(paragraphs.map((node) => paragraphTextOf(part, node.id))).toEqual([
      'Intro.',
      'First quoted.',
      'Second line.',
      'Outro.',
    ]);
    for (const node of paragraphs.slice(1, 3))
      expect(
        directParagraphProperties(part, node.id).find((item) => item.localName === 'pStyle')
          ?.attributes
      ).toEqual({ val: 'Quote' });
  } finally {
    editor.destroy();
  }
}, 60_000);

test('unresolved quote styles, rich cells, and unrepresented blocks refuse atomically', async () => {
  const plain = mount(p('Intro.'), null);
  const styled = mount(p('Intro.'));
  try {
    for (const [editor, target] of [
      [plain, 'Intro.\n\n> Quote.'],
      [styled, 'Intro.\n\n> - item'],
      [styled, 'Intro.\n\n```js\nx()\n```'],
      [styled, 'Intro.\n\n---'],
    ] as const) {
      const result = planTargetDocument(await source(editor), target);
      expect({ target, success: result.success }).toEqual({ target, success: false });
      expect(Object.hasOwn(result, 'plan')).toBe(false);
    }
  } finally {
    plain.destroy();
    styled.destroy();
  }
}, 60_000);
