import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput, targetParagraph as p } from '../../test/support/target-document';
import { applyTargetPlanOracle } from '../../test/support/target-plan-oracle';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { parseTargetMarkdown } from '../projection/parse-target';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';
import type { TargetSourceInput } from './target-source';
import type { CanonicalEdit } from './plan';

const ast = (markdown: string) => {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  return result.target.normalizedAst;
};
const run = (text: string, properties = '') =>
  `<w:r><w:rPr>${properties}<w:color w:val="1F4E79"/><keep:extra xmlns:keep="urn:keep" keep:flag="run"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const firstParagraph = (input: TargetSourceInput) => {
  const root = bodyStoryRoot(input.pkg.parts.get(input.pkg.mainDocumentPart)!)!;
  if (root.kind === 'textValue') throw new Error('Invalid body');
  return root.children.find((node) => node.kind === 'paragraph')!;
};
const runs = (paragraph: OoxmlNode) =>
  paragraph.kind === 'textValue' ? [] : paragraph.children.filter((node) => node.kind === 'run');

test('a bold prefix inserted before plain text formats only the inserted characters', async () => {
  const input = await targetDocumentInput(p('Tail.')),
    paragraph = firstParagraph(input),
    result = planTargetDocument(input, '**New** Tail.');
  if (!result.success) throw new Error(result.error);
  expect(result.plan.operations).toEqual([
    expect.objectContaining({ kind: 'text', start: 0, end: 0, before: '', after: 'New ' }),
    {
      kind: 'format',
      partName: '/word/document.xml',
      target: { kind: 'run', paragraphId: paragraph.id, start: 0, end: 3 },
      changes: [
        { name: 'b', before: { present: false }, after: { present: true, value: { val: '1' } } },
      ],
    },
  ]);
});

const cases: readonly [string, string, string, readonly CanonicalEdit['kind'][]][] = [
  ['bold prefix', p('Tail.'), '**New** Tail.', ['text', 'format']],
  ['italic word in the middle', p('One two.'), 'One *big* two.', ['text', 'format']],
  [
    'plain replacement of a bold word',
    `<w:p>${run('Old', '<w:b/>')}${run(' tail')}</w:p>`,
    'New tail',
    ['text', 'format'],
  ],
  [
    'bold replacement of a plain word',
    `<w:p>${run('Old tail')}</w:p>`,
    '**New** tail',
    ['text', 'format'],
  ],
  [
    'bold italic insertion',
    `<w:p>${run('Hello world')}</w:p>`,
    'Hello ***brave*** world',
    ['text', 'format'],
  ],
  [
    'struck suffix and plain prefix',
    `<w:p>${run('Mid', '<w:i/>')}</w:p>`,
    'Pre *Mid* ~~post~~',
    ['text', 'text', 'format', 'format', 'format'],
  ],
];
for (const [name, body, target, kinds] of cases)
  test(`${name}: the plan lands the exact target natively, directly, and through save/reopen`, async () => {
    const input = await targetDocumentInput(body),
      original = runs(firstParagraph(input)).map((node) => canonicalOoxmlFingerprint(node)),
      result = planTargetDocument(input, target);
    if (!result.success) throw new Error(result.error);
    expect(result.plan.operations.map((edit) => edit.kind)).toEqual([...kinds]);
    // Immutable acceptance oracle.
    const pkg = applyTargetPlanOracle(input.pkg, result.plan),
      reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(ast((await projectSourcePackage(reopened.package)).text)).toEqual(ast(target));
    // Every landed run keeps the inherited unknown run property beside the new marks.
    const landedRuns = runs(firstParagraph({ ...input, pkg }));
    for (const node of landedRuns) {
      expect(canonicalOoxmlFingerprint(node)).toContain('urn:keep');
      if (body.includes('1F4E79')) expect(canonicalOoxmlFingerprint(node)).toContain('1F4E79');
    }
    expect(original.length).toBeGreaterThan(0);

    // The live direct executor reaches the same result in one transaction.
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({ container, document: writeOoxmlPackage(input.pkg) });
    try {
      const token = readEditorRevision(editor, 'marks', 'mount')!,
        view = await readWorkingView(editor.surface!.session, token);
      if (!view.success) throw new Error(view.error);
      const live = planTargetDocument(
        {
          revision: token,
          pkg: editor.surface!.session.currentPackage(),
          projection: view.view.projection,
        },
        target
      );
      if (!live.success) throw new Error(live.error);
      const applied = applyDirectPlan(editor, live.plan, { documentId: 'marks', epoch: 'mount' });
      if (!applied.success) throw new Error(applied.error);
      const after = readEditorRevision(editor, 'marks', 'mount')!,
        landed = await readWorkingView(editor.surface!.session, after);
      if (!landed.success) throw new Error(landed.error);
      expect(ast(landed.view.projection.text)).toEqual(ast(target));
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(editor.surface!.session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
      container.remove();
    }
  }, 60_000);

test('inline code insertion plans the native character style after the text edit (D31)', async () => {
  const input = await targetDocumentInput(p('Use now.')),
    result = planTargetDocument(input, 'Use `x` now.');
  if (!result.success) throw new Error(result.error);
  const kinds = result.plan.operations.map((edit) => edit.kind);
  expect(kinds.indexOf('text')).toBeLessThan(kinds.indexOf('format'));
  expect(
    result.plan.operations.flatMap((edit) =>
      edit.kind === 'format' ? edit.changes.map((change) => change.name) : []
    )
  ).toContain('rStyle');
});
