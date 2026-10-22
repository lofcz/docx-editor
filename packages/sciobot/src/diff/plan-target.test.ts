import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  targetDocumentInput,
  targetLessonInput,
  targetParagraph as p,
  applyTargetRetained,
} from '../../test/support/target-document';
import {
  projectSourcePackage,
  sourceRevision,
  sourceControl,
} from '../../test/support/target-source-fixture';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { protectedDocument } from '../../test/support/protected-document';
import { checklistParagraph } from '../../test/support/list-document';
import pairs from '../__tests__/fixtures/ai-pairs.json';
import { buildBlockIndex } from '../blocks/block-index';
import { compileInlineTarget } from './inline-target';
import { parseTargetMarkdown } from '../projection/parse-target';
import { planTargetDocument } from './plan-target';
import { decodeEditPlan, serializeEditPlan } from './plan';
import type { TargetSourceInput } from './target-source';

const part = (input: TargetSourceInput) => input.pkg.parts.get(input.pkg.mainDocumentPart)!;
const paragraphIds = (input: TargetSourceInput): string[] => {
  const root = bodyStoryRoot(part(input))!;
  if (root.kind === 'textValue') throw new Error('Invalid body.');
  return root.children.filter((node) => node.kind === 'paragraph').map((node) => node.id);
};
function planned(input: TargetSourceInput, target: string) {
  const result = planTargetDocument(input, target);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result;
}
const nodeText = (ast: Parameters<typeof compileInlineTarget>[0]) => {
  const result = compileInlineTarget(ast);
  if (!result.success) throw new Error(result.error);
  return result.target.plainText;
};
function parsedAst(markdown: string) {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  return result.target.normalizedAst;
}

test('unchanged full targets preserve imported opaque body/parts through an empty plan', async () => {
  const loaded = readOoxmlPackage(protectedDocument());
  if (!loaded.ok) throw new Error(loaded.reason);
  const input = {
      revision: sourceRevision,
      pkg: loaded.package,
      projection: await projectSourcePackage(loaded.package),
    },
    before = writeOoxmlPackage(input.pkg);
  expect(planned(input, input.projection.text).plan.operations).toEqual([]);
  expect(writeOoxmlPackage(input.pkg)).toEqual(before);
});

test('C01 section insertion retains both paragraphs and uses their original physical gap', async () => {
  const input = await targetDocumentInput(p('Opening.') + p('Closing.')),
    ids = paragraphIds(input),
    before = writeOoxmlPackage(input.pkg),
    result = planned(input, pairs.cases[2]!.afterMarkdown);
  expect(result.plan.operations).toHaveLength(1);
  const edit = result.plan.operations[0]!;
  expect(edit).toMatchObject({
    kind: 'siblings',
    partName: input.pkg.mainDocumentPart,
    parentId: bodyStoryRoot(part(input))!.id,
    beforeId: ids[0],
    afterId: ids[1],
    removedIds: [],
  });
  if (edit.kind !== 'siblings') throw new Error('Expected sibling plan.');
  expect(edit.inserted.map((ast) => [ast.type, nodeText(ast)])).toEqual([
    ['heading', 'New section'],
    ['paragraph', 'First addition.'],
    ['paragraph', 'Second addition.'],
  ]);
  expect(result.pairs.map((pair) => pair.id)).toEqual(ids);
  expect(writeOoxmlPackage(input.pkg)).toEqual(before);
});

test('C01 deletion names only the removed paragraph and retains original neighbors', async () => {
  const input = await targetDocumentInput(['Keep.', 'Remove.', 'Tail.'].map(p).join('')),
    ids = paragraphIds(input),
    result = planned(input, pairs.cases[5]!.afterMarkdown);
  expect(result.plan.operations).toEqual([
    {
      kind: 'siblings',
      partName: input.pkg.mainDocumentPart,
      parentId: bodyStoryRoot(part(input))!.id,
      beforeId: ids[0]!,
      afterId: ids[2]!,
      removedIds: [ids[1]!],
      inserted: [],
    },
  ]);
  expect(result.pairs.map((pair) => pair.id)).toEqual([ids[0]!, ids[2]!]);
});

test('C01 complete rewrite retains reusable paragraph identities and inserts only the new heading', async () => {
  const input = await targetDocumentInput(p('Old document.') + p('Old end.')),
    ids = paragraphIds(input),
    result = planned(input, pairs.cases[6]!.afterMarkdown);
  const structure = result.plan.operations.filter((edit) => edit.kind === 'siblings');
  expect(structure).toHaveLength(1);
  expect(structure[0]).toMatchObject({
    beforeId: null,
    afterId: ids[0],
    removedIds: [],
    inserted: [{ type: 'heading', depth: 1 }],
  });
  const retainedPlan = {
      ...result.plan,
      operations: result.plan.operations.filter((edit) => edit.kind !== 'siblings'),
    },
    applied = applyTargetRetained(input.pkg, retainedPlan),
    nextPart = applied.parts.get(input.pkg.mainDocumentPart)!;
  expect(ids.map((id) => paragraphTextOf(nextPart, id))).toEqual(['New body.', 'New ending.']);
  expect(result.pairs.map((pair) => pair.id)).toEqual(ids);
  expect(structure[0]!.inserted.map(nodeText)).toEqual(['Replacement']);
  expect(paragraphTextOf(part(input), ids[0]!)).toBe('Old document.');
});

test('native paragraph text plans preserve inherited italic formatting and unchanged native subtrees', async () => {
  const input = await targetDocumentInput(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:r><w:rPr><w:i/><keep:extra xmlns:keep="urn:keep"/></w:rPr><w:t>italic</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>' +
        p('Tail.')
    ),
    ids = paragraphIds(input),
    result = planned(input, '**Bold** and *updated italic*.\n\nTail.');
  expect(result.plan.operations).toEqual([
    {
      kind: 'text',
      partName: input.pkg.mainDocumentPart,
      paragraphId: ids[0]!,
      start: 9,
      end: 9,
      before: '',
      after: 'updated ',
      bias: 'right',
    },
  ]);
  const pkg = applyTargetRetained(input.pkg, result.plan),
    next = await projectSourcePackage(pkg);
  expect(parsedAst(next.text)).toEqual(parsedAst('**Bold** and *updated italic*.\n\nTail.'));
  expect(findNode(pkg.parts.get(pkg.mainDocumentPart)!, ids[1]!)).toBe(
    findNode(part(input), ids[1]!)
  );
  expect(canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!)).toContain('extra');
});

test('same-text Markdown mark changes compose native run property deltas without replacing paragraphs', async () => {
  const input = await targetDocumentInput(
      '<w:p><w:r><w:rPr><w:i/><w:color w:val="123456"/><keep:extra xmlns:keep="urn:keep"/></w:rPr><w:t>Keep</w:t></w:r></w:p>'
    ),
    id = paragraphIds(input)[0]!,
    result = planned(input, '**Keep**');
  expect(result.plan.operations).toMatchObject([
    {
      kind: 'format',
      target: { kind: 'run', paragraphId: id, start: 0, end: 4 },
      changes: [
        { name: 'i', after: { present: true, value: { val: '0' } } },
        { name: 'b', after: { present: true, value: { val: '1' } } },
      ],
    },
  ]);
  const pkg = applyTargetRetained(input.pkg, result.plan),
    projection = await projectSourcePackage(pkg);
  expect(projection.text).toBe('**Keep**');
  expect(canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!)).toContain('123456');
  expect(canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!)).toContain('extra');
});

test('URL-only full targets address the original hyperlink wrapper and retain native link content', async () => {
  const input = await targetDocumentInput(
      '<w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:hyperlink r:id="link"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>',
      '<Relationship Id="link" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>'
    ),
    result = planned(input, '[Link](https://example.org)');
  expect(result.plan.operations).toMatchObject([
    {
      kind: 'format',
      target: { kind: 'hyperlink' },
      changes: [
        {
          name: 'target',
          after: { present: true, value: { kind: 'external', authored: 'https://example.org' } },
        },
      ],
    },
  ]);
  const pkg = applyTargetRetained(input.pkg, result.plan);
  expect((await projectSourcePackage(pkg)).text).toBe('[Link](https://example.org)');
  expect(paragraphTextOf(pkg.parts.get(pkg.mainDocumentPart)!, paragraphIds(input)[0]!)).toBe(
    'Link'
  );
});

test('C01 nested lesson changes compose activity metadata and child text while retaining native wrappers', async () => {
  const input = await targetLessonInput(),
    result = planned(input, pairs.cases[3]!.afterMarkdown),
    index = buildBlockIndex(input.pkg);
  expect(result.plan.operations.find((edit) => edit.kind === 'payload')).toMatchObject({
    kind: 'payload',
    blockId: 'activity',
    controlId: index.byId.get('activity')!.controlId,
    path: ['attributes', 'duration'],
    before: { present: true, value: 10 },
    after: { present: true, value: 15 },
  });
  expect(result.plan.operations.some((edit) => edit.kind === 'siblings')).toBe(false);
  const pkg = applyTargetRetained(input.pkg, result.plan),
    next = buildBlockIndex(pkg);
  expect(next.byId.get('activity')!.payload.attributes).toEqual({
    name: 'Warmup',
    duration: 15,
    privateData: { preserve: true },
  });
  for (const id of ['title', 'phase', 'activity', 'source'])
    expect(next.byId.get(id)!.controlId).toBe(index.byId.get(id)!.controlId);
  const nativePart = pkg.parts.get(pkg.mainDocumentPart)!;
  expect(
    storyParagraphs(bodyStoryRoot(nativePart)!).map((node) => paragraphTextOf(nativePart, node.id))
  ).toEqual(['Lesson', 'A better question.', 'Source', 'Last question.']);
});

test('C01 source content changes preserve opaque stored attributes while updating the owned source payload', async () => {
  const input = await targetLessonInput(),
    result = planned(input, pairs.cases[4]!.afterMarkdown);
  expect(result.plan.operations.filter((edit) => edit.kind === 'payload')).toMatchObject([
    {
      kind: 'payload',
      blockId: 'source',
      path: [],
      after: {
        present: true,
        value: { attributes: { url: 'https://example.org', privateData: { preserve: true } } },
      },
    },
  ]);
  expect(result.plan.operations.some((edit) => edit.kind === 'siblings')).toBe(false);
  const pkg = applyTargetRetained(input.pkg, result.plan),
    projection = await projectSourcePackage(pkg);
  expect(projection.text).toContain('url="https://example.org"');
  const nativePart = pkg.parts.get(pkg.mainDocumentPart)!;
  expect(
    storyParagraphs(bodyStoryRoot(nativePart)!).map((node) => paragraphTextOf(nativePart, node.id))
  ).toEqual(['Lesson', 'First question.', 'Source', 'Final reflection.']);
});

test('generated task placeholders stay outside authored text during a checklist state change', async () => {
  const input = await targetDocumentInput(checklistParagraph(false).replace('Task', '')),
    result = planned(input, '- [x]  ');
  expect(result.plan.operations).toMatchObject([
    {
      kind: 'list',
      before: { checkbox: { checked: false } },
      after: { kind: 'checklist', checked: true, level: 0 },
    },
  ]);
  expect(result.plan.operations.some((edit) => edit.kind === 'text')).toBe(false);
});

test('nested checklist intentions cannot disappear behind equal container content keys', async () => {
  const base = await targetDocumentInput(
    sourceControl('recap', checklistParagraph(false).replace('Task', ''))
  );
  const stored = withPayloadRecord(
    base.pkg,
    { version: 1, id: 'recap', type: 'recap_goals', attributes: {} },
    'recap'
  );
  if (!stored.ok) throw new Error(stored.reason);
  const input = { ...base, pkg: stored.pkg, projection: await projectSourcePackage(stored.pkg) },
    result = planned(input, '<recap_goals>\n\n- [x]  \n\n</recap_goals>');
  expect(result.plan.operations).toMatchObject([
    {
      kind: 'list',
      before: { checkbox: { checked: false } },
      after: { kind: 'checklist', checked: true },
    },
  ]);
});

test('paragraph HTML metadata is explicit while a bare wrapper preserves semantic identity', async () => {
  const input = await targetDocumentInput(p('Keep'));
  expect(planTargetDocument(input, '<p data-extra="keep">Keep</p>').success).toBe(false);
  expect(planned(input, '<p>Keep</p>').plan.operations).toEqual([]);
});

test('heading changes use an explicit resolved native style instead of replacing a retained paragraph', async () => {
  const input = await targetDocumentInput(p('Keep')),
    id = paragraphIds(input)[0]!;
  expect(planTargetDocument(input, '## Keep').success).toBe(false);
  const result = planTargetDocument(input, '## Keep', {
    headingStyle: (depth) => (depth === 2 ? 'Heading2' : null),
  });
  if (!result.success) throw new Error(result.error);
  expect(result.plan.operations).toMatchObject([
    {
      kind: 'format',
      target: { kind: 'paragraph', paragraphId: id },
      changes: [
        {
          name: 'pStyle',
          before: { present: false },
          after: { present: true, value: { val: 'Heading2' } },
        },
      ],
    },
  ]);
  expect(result.pairs.map((pair) => pair.id)).toEqual([id]);
});

test('unknown rich insertions refuse atomically, and inserted inline code plans (D31)', async () => {
  const input = await targetDocumentInput(p('Tail.'));
  const result = planTargetDocument(input, '<Unknown value={danger()} />\n\nTail.');
  expect(result.success).toBe(false);
  expect(Object.hasOwn(result, 'plan')).toBe(false);
  expect((globalThis as Record<string, unknown>).danger).toBeUndefined();
  // N18 supplies the native code identity, so inserted code is an explicit style change.
  expect(planTargetDocument(input, '`New`Tail.').success).toBe(true);
});

test('plans freeze/serialize plain intentions, and malformed or stale input never returns a partial plan', async () => {
  const input = await targetDocumentInput(p('Before.')),
    result = planned(input, 'After.'),
    bytes = writeOoxmlPackage(input.pkg);
  expect(Object.isFrozen(result.plan) && Object.isFrozen(result.plan.operations)).toBe(true);
  expect(decodeEditPlan(serializeEditPlan(result.plan))).toEqual({
    success: true,
    plan: result.plan,
  });
  for (const value of [
    planTargetDocument(input, '<Open>'),
    planTargetDocument({ ...input, revision: { ...sourceRevision, revision: 99 } }, 'After.'),
  ]) {
    expect(value.success).toBe(false);
    expect(Object.hasOwn(value, 'plan')).toBe(false);
  }
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
});

test('planning preserves live bytes, revision, selection, undo/redo history, and imported paragraph XML', async () => {
  const input = await targetDocumentInput(p('Before.')),
    editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(input.pkg),
    });
  try {
    const surface = editor.surface!,
      session = surface.session;
    surface.applyAutomationOps(() => [
      { op: 'insertText', paragraphId: session.paragraphIds()[0]!, offset: 0, text: 'User ' },
    ]);
    const bytes = writeOoxmlPackage(session.currentPackage()),
      revision = session.packageRevision(),
      xml = canonicalOoxmlFingerprint(session.part()),
      selection = surface.state().selection,
      history = [session.canUndo(), session.canRedo()];
    const projection = await projectSourcePackage(session.currentPackage());
    expect(
      planned(
        { revision: sourceRevision, pkg: session.currentPackage(), projection },
        'User After.'
      ).plan.operations.length
    ).toBeGreaterThan(0);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(session.packageRevision()).toBe(revision);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(xml);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
  } finally {
    editor.destroy();
  }
});
