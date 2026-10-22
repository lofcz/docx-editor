import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { readOoxmlPackage, writeOoxmlPackage, type OoxmlNode } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { parseTargetMarkdown } from '../projection/parse-target';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from '../ai/apply-direct';
import { decodeEditPlan, serializeEditPlan, type CanonicalEdit } from './plan';
import { planTargetDocument } from './plan-target';

const context = { documentId: 'links', epoch: 'mount' };
const REL =
  '<Relationship Id="link" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>';
const linked = (before: string, text: string, after: string, tooltip = '') =>
  `<w:p>${before ? `<w:r><w:t xml:space="preserve">${before}</w:t></w:r>` : ''}<w:hyperlink r:id="link"${tooltip ? ` w:tooltip="${tooltip}"` : ''}><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:hyperlink>${after ? `<w:r><w:t xml:space="preserve">${after}</w:t></w:r>` : ''}</w:p>`;
const ast = (markdown: string) => {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  return result.target.normalizedAst;
};
function mount(body: string) {
  const container = document.createElement('div');
  document.body.append(container);
  return createDocxEditor({
    container,
    document: docxFixture(body, { relationships: REL }),
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
const wrappers = (editor: DocxEditorInstance) => {
  const out: string[] = [],
    stack: OoxmlNode[] = [editor.surface!.session.part().root];
  while (stack.length) {
    const node = stack.pop()!;
    if (node.kind === 'hyperlink') out.push(node.id);
    else if (node.kind !== 'textValue') stack.push(...node.children);
  }
  return out;
};

const cases: readonly [string, string, string, readonly CanonicalEdit['kind'][], boolean][] = [
  // name, body, target, expected operation kinds, original wrapper survives
  [
    'link around existing words',
    '<w:p><w:r><w:t>See the docs now.</w:t></w:r></w:p>',
    'See the [docs](https://example.org) now.',
    ['link'],
    false,
  ],
  ['link removal', linked('See ', 'docs', ' now.'), 'See docs now.', ['unlink'], false],
  [
    'changed link text',
    linked('See ', 'docs', ' now.'),
    'See [the manual](https://example.com) now.',
    ['text', 'text', 'text', 'unlink', 'link'],
    false,
  ],
  [
    'text after a link',
    linked('', 'Link', ' tail'),
    '[Link](https://example.com) new tail',
    ['text'],
    true,
  ],
  [
    'text before a retargeted link',
    linked('See ', 'docs', ''),
    'Please see [docs](https://example.org)',
    ['text', 'text', 'text', 'format'],
    true,
  ],
  [
    'tooltip removal',
    linked('See ', 'docs', '', 'Old tip'),
    'See [docs](https://example.com)',
    ['unlink', 'link'],
    false,
  ],
];
for (const [name, body, target, kinds, survives] of cases)
  test(`${name}: plans explicit link operations and lands the target`, async () => {
    const editor = mount(body);
    try {
      const original = wrappers(editor),
        result = planTargetDocument(await source(editor), target);
      if (!result.success) throw new Error(`${result.code}: ${result.error}`);
      expect(result.plan.operations.map((edit) => edit.kind)).toEqual([...kinds]);
      expect(decodeEditPlan(serializeEditPlan(result.plan))).toEqual({
        success: true,
        plan: result.plan,
      });
      const applied = applyDirectPlan(editor, result.plan, context);
      if (!applied.success) throw new Error(`${applied.code}: ${applied.error}`);
      expect(ast((await source(editor)).projection.text)).toEqual(ast(target));
      const replanned = planTargetDocument(await source(editor), target);
      expect(replanned.success && replanned.plan.operations).toEqual([]);
      if (survives) expect(wrappers(editor)).toEqual(original);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(editor.surface!.session.canUndo()).toBe(false);
      expect(wrappers(editor)).toEqual(original);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      const container = document.createElement('div');
      document.body.append(container);
      const again = createDocxEditor({ container, document: writeOoxmlPackage(reopened.package) });
      try {
        expect(ast((await source(again)).projection.text)).toEqual(ast(target));
      } finally {
        again.destroy();
        container.remove();
      }
    } finally {
      editor.destroy();
    }
  }, 60_000);

test('link plans carry no relationship IDs and refuse empty or unsafe targets', async () => {
  const editor = mount('<w:p><w:r><w:t>Word</w:t></w:r></w:p>');
  try {
    const result = planTargetDocument(await source(editor), '[Word](https://example.org/x)');
    if (!result.success) throw new Error(result.error);
    expect(result.plan.operations).toEqual([
      expect.objectContaining({
        kind: 'link',
        start: 0,
        end: 4,
        target: { kind: 'external', url: 'https://example.org/x' },
        title: null,
      }),
    ]);
    expect(JSON.stringify(result.plan)).not.toContain('rId');
    const unsafe = planTargetDocument(await source(editor), '[Word](javascript:alert(1))');
    if (unsafe.success) {
      const applied = applyDirectPlan(editor, unsafe.plan, context);
      expect(applied.success).toBe(false);
      expect(editor.surface!.session.canUndo()).toBe(false);
    }
  } finally {
    editor.destroy();
  }
}, 60_000);
