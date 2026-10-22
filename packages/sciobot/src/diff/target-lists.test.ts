import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  contentControlPropertiesOf,
  directParagraphProperties,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { listDocument, listParagraph, checklistParagraph } from '../../test/support/list-document';
import { targetDocumentInput, targetParagraph as p } from '../../test/support/target-document';
import { applyTargetPlanOracle } from '../../test/support/target-plan-oracle';
import { projectSourcePackage, sourceRevision } from '../../test/support/target-source-fixture';
import reference from '../__tests__/fixtures/ai-pairs.json';
import { planTargetDocument } from './plan-target';
import type { TargetSourceInput } from './target-source';
import type { CanonicalEditPlan } from './plan';

async function numbered(body: string): Promise<TargetSourceInput> {
  const loaded = readOoxmlPackage(listDocument(body));
  if (!loaded.ok) throw new Error(loaded.reason);
  return {
    revision: sourceRevision,
    pkg: loaded.package,
    projection: await projectSourcePackage(loaded.package),
  };
}
const part = (input: TargetSourceInput) => input.pkg.parts.get(input.pkg.mainDocumentPart)!;
const paragraphs = (input: TargetSourceInput) => storyParagraphs(bodyStoryRoot(part(input))!);
function plan(input: TargetSourceInput, markdown: string): CanonicalEditPlan {
  const result = planTargetDocument(input, markdown);
  if (!result.success) throw new Error(result.error);
  return result.plan;
}

test('C01 formatting/list insertion carries the original numbering instance without edits for later ordinals', async () => {
  const input = await numbered(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>' +
        listParagraph('First', '1') +
        listParagraph('Second', '1') +
        listParagraph('Apple', '2') +
        listParagraph('Pear', '2')
    ),
    ids = paragraphs(input).map((node) => node.id),
    before = writeOoxmlPackage(input.pkg),
    result = plan(input, reference.cases[0]!.afterMarkdown);
  expect(result.operations).toHaveLength(2);
  expect(result.operations[0]).toMatchObject({
    kind: 'text',
    paragraphId: ids[0],
    start: 9,
    end: 9,
    after: 'updated ',
    bias: 'right',
  });
  expect(result.operations[1]).toMatchObject({
    kind: 'siblings',
    beforeId: ids[1],
    afterId: ids[2],
    removedIds: [],
    inserted: [{ type: 'paragraph', children: [{ type: 'text', value: 'Added' }] }],
    lists: [
      {
        path: [0],
        target: { kind: 'ordered', level: 0, instance: { kind: 'existing', numId: '1' } },
      },
    ],
  });
  expect(result.operations.some((edit) => edit.kind === 'list')).toBe(false);
  const pkg = applyTargetPlanOracle(input.pkg, result),
    landed = pkg.parts.get(pkg.mainDocumentPart)!;
  expect(
    storyParagraphs(bodyStoryRoot(landed)!).map((node) => paragraphTextOf(landed, node.id))
  ).toEqual(['Bold and updated italic.', 'First', 'Added', 'Second', 'Apple', 'Pear']);
  for (const id of ids.slice(1)) expect(findNode(landed, id)).toBe(findNode(part(input), id));
  expect(pkg.parts.get('/word/numbering.xml')).toBe(input.pkg.parts.get('/word/numbering.xml'));
  const projection = await projectSourcePackage(pkg);
  expect(projection.text).toContain('2. Added');
  expect(projection.text).toContain('3. Second');
  const reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect((await projectSourcePackage(reopened.package)).text).toBe(projection.text);
  expect(writeOoxmlPackage(input.pkg)).toEqual(before);
});

test('retained ordinals remain presentation data, including changed first markers', async () => {
  const input = await numbered(listParagraph('First', '1') + listParagraph('Second', '1'));
  const result = plan(input, '99. First\n100. Added\n101. Second');
  expect(result.operations).toHaveLength(1);
  expect(result.operations[0]).toMatchObject({
    kind: 'siblings',
    lists: [{ target: { kind: 'ordered', instance: { kind: 'existing', numId: '1' } } }],
  });
  expect((await projectSourcePackage(applyTargetPlanOracle(input.pkg, result))).text).toContain(
    '3. Second'
  );
});

test('new groups share deferred keys and preserve an authored start only for the new instance', async () => {
  const input = await targetDocumentInput(p('Tail')),
    result = plan(input, '7. A\n8. B\n\nTail');
  expect(result.operations).toMatchObject([
    {
      kind: 'siblings',
      lists: [
        {
          path: [0],
          target: {
            kind: 'ordered',
            level: 0,
            instance: { kind: 'new', key: 'target-list:0', start: 7 },
          },
        },
        {
          path: [1],
          target: {
            kind: 'ordered',
            level: 0,
            instance: { kind: 'new', key: 'target-list:0', start: 7 },
          },
        },
      ],
    },
  ]);
  const pkg = applyTargetPlanOracle(input.pkg, result),
    projection = await projectSourcePackage(pkg);
  expect(projection.text).toContain('7. A');
  expect(projection.text).toContain('8. B');
  expect(input.pkg.parts.has('/word/numbering.xml')).toBe(false);
});

test('equal inserted ASTs in separate reversed gaps keep their distinct list targets', async () => {
  const input = await targetDocumentInput(p('A') + p('B') + p('C')),
    result = plan(input, '1. Same\n\nA\n\n- Same\n\nB\n\nC');
  expect(result.operations).toHaveLength(2);
  expect(result.operations[0]).toMatchObject({
    kind: 'siblings',
    lists: [
      { path: [0], target: { kind: 'bullet', instance: { kind: 'new', key: 'target-list:2' } } },
    ],
  });
  expect(result.operations[1]).toMatchObject({
    kind: 'siblings',
    lists: [
      {
        path: [0],
        target: { kind: 'ordered', instance: { kind: 'new', key: 'target-list:0', start: 1 } },
      },
    ],
  });
  const projection = await projectSourcePackage(applyTargetPlanOracle(input.pkg, result));
  expect(projection.text).toContain('1. Same');
  expect(projection.text).toContain('- Same');
});

test('adjacent retained Word instances stay distinct and inserted items use the preceding compatible neighbor', async () => {
  const input = await numbered(listParagraph('First', '1') + listParagraph('Second', '3')),
    result = plan(input, '1. First\n2. Added\n3. Second');
  // Markdown ordinals retain the original Word format and disconnected numbering instance.
  expect(result.operations.some((edit) => edit.kind === 'list')).toBe(false);
  expect(result.operations[0]).toMatchObject({
    lists: [{ target: { instance: { kind: 'existing', numId: '1' } } }],
  });
  const landed = applyTargetPlanOracle(input.pkg, result).parts.get(input.pkg.mainDocumentPart)!;
  expect(findNode(landed, paragraphs(input)[1]!.id)).toBe(paragraphs(input)[1]);
  const source = await numbered(listParagraph('First', '1') + listParagraph('Second', '1'));
  const result2 = plan(source, '1. Added\n2. First\n3. Second');
  expect(result2.operations).toMatchObject([
    { kind: 'siblings', lists: [{ target: { instance: { kind: 'existing', numId: '1' } } }] },
  ]);
});

test('retained paragraphs joining one new list share a key with its inserted items', async () => {
  const input = await targetDocumentInput(p('Kept') + p('Tail')),
    result = plan(input, '5. Added\n6. Kept\n\nTail');
  expect(result.operations).toMatchObject([
    {
      kind: 'list',
      after: { kind: 'ordered', instance: { kind: 'new', key: 'target-list:0', start: 5 } },
    },
    {
      kind: 'siblings',
      lists: [
        { target: { kind: 'ordered', instance: { kind: 'new', key: 'target-list:0', start: 5 } } },
      ],
    },
  ]);
  expect((await projectSourcePackage(applyTargetPlanOracle(input.pkg, result))).text).toContain(
    '6. Kept'
  );
});

test('checklist removal follows body text edits so original offsets remain valid', async () => {
  const input = await targetDocumentInput(checklistParagraph(false)),
    result = plan(input, 'Changed');
  expect(result.operations.at(-1)).toMatchObject({ kind: 'list', after: { kind: 'none' } });
  const text = result.operations.filter((edit) => edit.kind === 'text');
  expect(text.length).toBeGreaterThan(0);
  for (const edit of text) if (edit.kind === 'text') expect(edit.start).toBeGreaterThanOrEqual(1);
  const pkg = applyTargetPlanOracle(input.pkg, result),
    native = pkg.parts.get(pkg.mainDocumentPart)!;
  expect(paragraphTextOf(native, paragraphs(input)[0]!.id)).toBe('Changed');
  expect(
    directParagraphProperties(native, paragraphs(input)[0]!.id).some(
      (property) => property.localName === 'ind'
    )
  ).toBe(false);
});

test('nested inserted numbering retains its explicit level and new-group start through reopen', async () => {
  const input = await targetDocumentInput(p('Tail')),
    result = plan(input, '- Parent\n\n  7. Child\n  8. Next\n\nTail');
  const edit = result.operations[0]!;
  if (edit.kind !== 'siblings') throw new Error('Expected insertion');
  expect(edit.lists?.map((binding) => binding.target)).toMatchObject([
    { kind: 'bullet', level: 0 },
    { kind: 'ordered', level: 1, instance: { kind: 'new', start: 7 } },
    { kind: 'ordered', level: 1, instance: { kind: 'new', start: 7 } },
  ]);
  const applied = applyTargetPlanOracle(input.pkg, result),
    projection = await projectSourcePackage(applied);
  expect(projection.text).toContain('  7. Child');
  expect(projection.text).toContain('  8. Next');
  const reopened = readOoxmlPackage(writeOoxmlPackage(applied));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect((await projectSourcePackage(reopened.package)).text).toBe(projection.text);
});

test('retained checklist nesting and state preserve the owned control identity', async () => {
  const input = await targetDocumentInput(checklistParagraph(false)),
    controlId = input.projection.checklists[0]!.controlId,
    result = plan(input, '- [ ] Outer\n  - [x] Task');
  expect(result.operations[0]).toMatchObject({
    kind: 'list',
    before: { checkbox: { controlId } },
    after: { kind: 'checklist', level: 1, checked: true },
  });
  const applied = applyTargetPlanOracle(input.pkg, result),
    projection = await projectSourcePackage(applied);
  expect(projection.checklists.map((task) => [task.checked, task.level])).toEqual([
    [false, 0],
    [true, 1],
  ]);
  expect(projection.checklists[1]!.controlId).toBe(controlId);
  expect(
    directParagraphProperties(
      applied.parts.get(input.pkg.mainDocumentPart)!,
      paragraphs(input)[0]!.id
    )
  ).toContainEqual({
    localName: 'ind',
    attributes: { left: '1440', hanging: '360' },
  });
});

test('inserted nested items reuse a compatible parent Word instance and preserve authored numbering', async () => {
  const input = await numbered(listParagraph('Parent', '1') + p('Tail')),
    result = plan(input, '1. Parent\n\n   1. Child\n\nTail');
  expect(result.operations).toMatchObject([
    {
      kind: 'siblings',
      lists: [
        { target: { kind: 'ordered', level: 1, instance: { kind: 'existing', numId: '1' } } },
      ],
    },
  ]);
  const applied = applyTargetPlanOracle(input.pkg, result);
  expect(applied.parts.get('/word/numbering.xml')).toBe(input.pkg.parts.get('/word/numbering.xml'));
  expect((await projectSourcePackage(applied)).text).toContain('   1. Child');
});

test('unresolved nested formats and conflicting new starts refuse without rewriting existing definitions', async () => {
  const input = await numbered(listParagraph('Parent', '2') + p('Tail')),
    before = writeOoxmlPackage(input.pkg),
    result = planTargetDocument(input, '- Parent\n\n  7. Child\n\nTail');
  expect(result.success).toBe(false);
  expect('plan' in result).toBe(false);
  expect(writeOoxmlPackage(input.pkg)).toEqual(before);
  const plain = await targetDocumentInput(p('Tail')),
    conflicting = planTargetDocument(plain, '- A\n\n  7. First\n\n- B\n\n  2. Second\n\nTail');
  expect(conflicting.success).toBe(false);
  expect('plan' in conflicting).toBe(false);
  expect(planTargetDocument(plain, '- Parent\n  - [ ] Task\n    1. Child\n\nTail').success).toBe(
    false
  );
});

test('a new nested group preserves an imported start and cannot silently change that definition', async () => {
  const entries = unzipSync(listDocument(listParagraph('Parent', '1') + p('Tail')));
  entries['word/numbering.xml'] = strToU8(
    strFromU8(entries['word/numbering.xml']!).replace(
      '<w:lvl w:ilvl="1"><w:start w:val="1"/>',
      '<w:lvl w:ilvl="1"><w:start w:val="7"/>'
    )
  );
  const loaded = readOoxmlPackage(zipSync(entries));
  if (!loaded.ok) throw new Error(loaded.reason);
  const input = {
      pkg: loaded.package,
      revision: sourceRevision,
      projection: await projectSourcePackage(loaded.package),
    },
    incompatible = planTargetDocument(input, '1. Parent\n\n   1. Child\n\nTail');
  expect(incompatible.success).toBe(false);
  expect('plan' in incompatible).toBe(false);
  const result = plan(input, '1. Parent\n\n   7. Child\n\nTail'),
    applied = applyTargetPlanOracle(input.pkg, result);
  expect(applied.parts.get('/word/numbering.xml')).toBe(input.pkg.parts.get('/word/numbering.xml'));
  expect((await projectSourcePackage(applied)).text).toContain('   7. Child');
  const ordinary = await numbered(listParagraph('Parent', '1') + p('Tail'));
  expect(planTargetDocument(ordinary, '1. Parent\n\n   7. Child\n\nTail').success).toBe(false);
});

test('list nesting beyond the native ninth level refuses without a partial plan or source writes', async () => {
  const input = await targetDocumentInput(p('Tail')),
    bytes = writeOoxmlPackage(input.pkg),
    target =
      Array.from({ length: 10 }, (_, level) => `${'  '.repeat(level)}- Item ${level}`).join('\n') +
      '\n\nTail',
    result = planTargetDocument(input, target);
  expect(result.success).toBe(false);
  expect('plan' in result).toBe(false);
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
});

test('inserted tasks carry state independently of authored NBSP and allocate native controls during application', async () => {
  const input = await targetDocumentInput(p('Tail')),
    result = plan(input, '- [x]  \n- [ ] Work\n\nTail');
  expect(result.operations).toMatchObject([
    {
      kind: 'siblings',
      lists: [
        { target: { kind: 'checklist', checked: true, level: 0 } },
        { target: { kind: 'checklist', checked: false, level: 0 } },
      ],
    },
  ]);
  const pkg = applyTargetPlanOracle(input.pkg, result),
    projection = await projectSourcePackage(pkg);
  expect(projection.checklists.map((task) => [task.checked, task.empty])).toEqual([
    [true, false],
    [false, false],
  ]);
  expect(
    contentControlPropertiesOf(
      findNode(pkg.parts.get(pkg.mainDocumentPart)!, projection.checklists[0]!.controlId)!
    ).checkbox?.checked
  ).toBe(true);
  expect(
    paragraphTextOf(
      pkg.parts.get(pkg.mainDocumentPart)!,
      projection.checklists[0]!.paragraphId
    )?.slice(1)
  ).toBe(' ');
});

test('source state, fingerprints, and native allocation remain unchanged during list planning', async () => {
  const input = await numbered(listParagraph('First', '1') + listParagraph('Second', '1')),
    bytes = writeOoxmlPackage(input.pkg),
    fingerprint = canonicalOoxmlFingerprint(part(input));
  plan(input, '1. First\n2. Added\n3. Second');
  expect(writeOoxmlPackage(input.pkg)).toEqual(bytes);
  expect(canonicalOoxmlFingerprint(part(input))).toBe(fingerprint);
});
