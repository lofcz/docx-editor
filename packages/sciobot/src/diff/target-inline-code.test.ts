import { expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { applyDirectPlan } from '../ai/apply-direct';
import { docxFixture } from '../../test/support/document';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readWorkingView } from '../review/working-view';
import { planTargetDocument } from './plan-target';
import type { CanonicalEdit } from './plan';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const context = { documentId: 'inline-code-plan', epoch: 'mount-1' };
const STYLES =
  `<w:styles xmlns:w="${W}">` +
  '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/></w:style>' +
  '<w:style w:type="character" w:customStyle="1" w:styleId="SciobotCode"><w:name w:val="Sciobot Code"/></w:style>' +
  '</w:styles>';
const p = (runs: string) => `<w:p>${runs}</w:p>`;
const r = (text: string, props = '') =>
  `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const code = (text: string, props = '') => r(text, `<w:rStyle w:val="SciobotCode"/>${props}`);

function mount(bytes: Uint8Array): DocxEditorInstance {
  const container = document.createElement('div');
  document.body.append(container);
  return createDocxEditor({
    container,
    document: bytes,
    author: 'Teacher',
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
}
const token = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, context.documentId, context.epoch)!;
async function projection(editor: DocxEditorInstance) {
  const view = await readWorkingView(editor.surface!.session, token(editor));
  if (!view.success) throw new Error(view.error);
  return view.view.projection;
}
async function plan(editor: DocxEditorInstance, target: string) {
  const result = planTargetDocument(
    {
      revision: token(editor),
      pkg: editor.surface!.session.currentPackage(),
      projection: await projection(editor),
    },
    target
  );
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
const styleChanges = (operations: readonly CanonicalEdit[]) =>
  operations.flatMap((edit) =>
    edit.kind === 'format'
      ? edit.changes
          .filter((change) => change.name === 'rStyle')
          .map((change) => ({
            range: edit.target.kind === 'run' ? [edit.target.start, edit.target.end] : null,
            before: change.before,
            after: change.after,
          }))
      : []
  );

/** Plan, apply directly, replan to nothing, undo exactly, and reopen to the same Markdown. */
async function land(body: string, before: string, target: string, styles?: string) {
  const editor = mount(docxFixture(body, styles ? { styles } : {}));
  try {
    expect((await projection(editor)).text).toBe(before);
    const original = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const planned = await plan(editor, target);
    const result = applyDirectPlan(editor, planned, context);
    if (!result.success) throw new Error(`${result.code}: ${result.error}`);
    expect((await projection(editor)).text).toBe(target);
    expect((await plan(editor, target)).operations).toEqual([]);
    const saved = new Uint8Array(await editor.save());
    const styleXml = unzipSync(saved)['word/styles.xml'];
    const definitions = styleXml
      ? (strFromU8(styleXml).match(/w:styleId="SciobotCode"/g) ?? []).length
      : 0;
    const reopened = mount(saved);
    try {
      expect((await projection(reopened)).text).toBe(target);
      expect((await plan(reopened, target)).operations).toEqual([]);
    } finally {
      reopened.destroy();
    }
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(original);
    return { operations: planned.operations, definitions };
  } finally {
    editor.destroy();
  }
}

test('adding inline code to retained text plans an explicit character style change', async () => {
  const { operations, definitions } = await land(
    p(r('alpha beta gamma', '<w:b/>')),
    '**alpha beta gamma**',
    '**alpha `beta` gamma**'
  );
  expect(operations.filter((edit) => edit.kind === 'text')).toEqual([]);
  expect(styleChanges(operations)).toEqual([
    {
      range: [6, 10],
      before: { present: false },
      after: { present: true, value: { val: 'SciobotCode' } },
    },
  ]);
  expect(definitions).toBe(1);
});

test('removing inline code restores ordinary retained text', async () => {
  const { operations } = await land(
    p(r('alpha ') + code('beta') + r(' gamma')),
    'alpha `beta` gamma',
    'alpha beta gamma',
    STYLES
  );
  expect(styleChanges(operations)).toEqual([
    {
      range: [6, 10],
      before: { present: true, value: { val: 'SciobotCode' } },
      after: { present: false },
    },
  ]);
});

test('inserted text inside and beside code lands with the intended code identity', async () => {
  await land(p(r('alpha gamma')), 'alpha gamma', 'alpha `beta` gamma');
  await land(
    p(r('see ') + code('code') + r(' here')),
    'see `code` here',
    'see `codes` and `more` here',
    STYLES
  );
});

test('code combines with toggles and link text, and returns link text to its style', async () => {
  const { operations } = await land(p(r('alpha beta')), 'alpha beta', 'alpha **`beta`**');
  // One run slice carries both the toggle and the character style change.
  const formats = operations.filter((edit) => edit.kind === 'format');
  expect(formats).toHaveLength(1);
  expect(
    formats[0]!.kind === 'format' && formats[0]!.changes.map((change) => change.name).sort()
  ).toEqual(['b', 'rStyle']);
  const link =
    '<w:bookmarkStart w:id="1" w:name="here"/>' +
    r('Go ') +
    '<w:hyperlink w:anchor="here">' +
    r('there', '<w:rStyle w:val="Hyperlink"/>') +
    '</w:hyperlink><w:bookmarkEnd w:id="1"/>';
  await land(p(link), 'Go [there](#here)', 'Go [`there`](#here)', STYLES);
  const back = await land(
    p(
      '<w:bookmarkStart w:id="1" w:name="here"/>' +
        r('Go ') +
        `<w:hyperlink w:anchor="here">${code('there')}</w:hyperlink><w:bookmarkEnd w:id="1"/>`
    ),
    'Go [`there`](#here)',
    'Go [there](#here)',
    STYLES
  );
  expect(styleChanges(back.operations)).toEqual([
    {
      range: [3, 8],
      before: { present: true, value: { val: 'SciobotCode' } },
      after: { present: true, value: { val: 'Hyperlink' } },
    },
  ]);
});

test('new paragraphs and table cells carry inline code and gain the style definition', async () => {
  const { definitions } = await land(
    p(r('First.')),
    'First\\.',
    'First\\.\n\nRun `npm test` now\\.\n\n| A | B |\n| --- | --- |\n| `x` | y |'
  );
  expect(definitions).toBe(1);
});

test('code plans preserve an imported character style and restore it on removal', async () => {
  const styles = `<w:styles xmlns:w="${W}"><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style></w:styles>`;
  const added = await land(
    p(r('strong', '<w:rStyle w:val="Strong"/>')),
    '**strong**',
    '**`strong`**',
    styles
  );
  expect(styleChanges(added.operations)[0]).toMatchObject({
    before: { present: true, value: { val: 'Strong' } },
    after: { present: true, value: { val: 'SciobotCode__Strong' } },
  });
  const derived = styles.replace(
    '</w:styles>',
    '<w:style w:type="character" w:styleId="SciobotCode__Strong"><w:name w:val="Sciobot Code"/><w:basedOn w:val="Strong"/></w:style></w:styles>'
  );
  const removed = await land(
    p(r('strong', '<w:rStyle w:val="SciobotCode__Strong"/>')),
    '**`strong`**',
    '**strong**',
    derived
  );
  expect(styleChanges(removed.operations)[0]).toMatchObject({
    before: { present: true, value: { val: 'SciobotCode__Strong' } },
    after: { present: true, value: { val: 'Strong' } },
  });
});
