import { expect, test } from 'bun:test';
import { readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { planInsertedAtoms, type InlineInsertionLocation } from './inserted-atoms';
const packageOf = () => {
  const read = readOoxmlPackage(docxFixture('<w:p><w:r><w:t>Keep</w:t></w:r></w:p>'));
  if (!read.ok) throw Error(read.reason);
  return read.package;
};
const paragraphOf = (markdown: string): TargetAstNode => {
  const read = parseTargetMarkdown(markdown);
  if (!read.success) throw Error(read.error);
  return read.target.normalizedAst.children![0] as TargetAstNode;
};
for (const location of [
  { kind: 'new-paragraph', path: [2, 0] },
  { kind: 'retained-paragraph', paragraphId: 'existing-native-paragraph' },
] as const)
  test(`all four atom kinds have explicit deferred operations in ${location.kind} fragments`, () => {
    const pkg = packageOf(),
      before = writeOoxmlPackage(pkg);
    let allocated = 0;
    const result = planInsertedAtoms(
      { pkg, location, offset: 10, allocateBlockId: () => `atom-${++allocated}` },
      paragraphOf(
        'Lead ![α](asset:img "Figure") $x^2$ [Ada Lovelace](mention:user-7) <source title="Paper" url="https://x.test" /> end'
      )
    );
    if (!result.success) throw Error(result.error);
    expect(result.plan.modelLength).toBe(33);
    const ops = result.plan.operations;
    expect(
      ops.map((op) => ({ kind: op.kind, atom: op.atom, offset: op.offset, location: op.location }))
    ).toEqual([
      { kind: 'drawing', atom: 'image', offset: 15, location },
      { kind: 'inline-control', atom: 'math', offset: 17, location },
      { kind: 'inline-control', atom: 'mention', offset: 21, location },
      { kind: 'inline-control', atom: 'source', offset: 34, location },
    ]);
    expect(ops[0]).toMatchObject({
      resource: { reference: 'asset:img', alt: 'α', title: 'Figure' },
      modelLength: 1,
    });
    const controls = ops.filter((op) => op.kind === 'inline-control');
    expect(controls.map((op) => op.payload)).toEqual([
      { id: 'atom-1', type: 'inline_equation', version: 1, attributes: { texExpression: 'x^2' } },
      {
        id: 'atom-2',
        type: 'mention',
        version: 1,
        attributes: { key: 'user-7', value: 'Ada Lovelace' },
      },
      {
        id: 'atom-3',
        type: 'source',
        version: 1,
        attributes: { title: 'Paper', url: 'https://x.test' },
      },
    ]);
    expect(controls.map((op) => op.fallback)).toEqual(['x^2', 'Ada Lovelace', 'Paper']);
    expect(allocated).toBe(3);
    expect(writeOoxmlPackage(pkg)).toEqual(before);
    expect(Object.isFrozen(result.plan.operations)).toBe(true);
    expect(Object.isFrozen(controls[0]!.payload.attributes)).toBe(true);
  });

test('atom placement retains UTF-16 text widths, marks, link ownership, and source identity', () => {
  const pkg = packageOf();
  const location: InlineInsertionLocation = { kind: 'new-paragraph', path: [3] };
  const result = planInsertedAtoms(
    { pkg, location, offset: 0, allocateBlockId: () => 'math' },
    paragraphOf('𝛼 [**$z$**](https://example.test "link")')
  );
  if (!result.success) throw Error(result.error);
  expect(result.plan.operations).toHaveLength(1);
  expect(result.plan.operations[0]).toMatchObject({
    offset: 3,
    marks: ['bold'],
    hyperlink: { target: { kind: 'external', url: 'https://example.test' }, title: 'link' },
  });
  expect(result.plan.modelLength).toBe(4);
  (location.path as number[])[0] = 99;
  expect(result.plan.operations[0]!.location).toEqual({ kind: 'new-paragraph', path: [3] });
});

test('document mention identifiers and at-mention labels remain distinct from persistent block IDs', () => {
  for (const [markdown, key, value] of [
    ['[Lesson One](mention:%2Fdocs%2Flesson-1)', '/docs/lesson-1', 'Lesson One'],
    ['@ada', 'ada', '@ada'],
  ]) {
    const result = planInsertedAtoms(
      {
        pkg: packageOf(),
        location: { kind: 'retained-paragraph', paragraphId: 'p' },
        offset: 0,
        allocateBlockId: () => 'block-identity',
      },
      paragraphOf(markdown!)
    );
    if (!result.success) throw Error(result.error);
    expect(result.plan.operations[0]).toMatchObject({
      payload: { id: 'block-identity', type: 'mention', attributes: { key, value } },
      fallback: value,
    });
  }
});

test('empty and multiline application math preserve source while carrying explicit Word fallback widths', () => {
  for (const [source, fallback] of [
    ['', '□'],
    [' x\r\n y\n', ' x  y '],
  ]) {
    const result = planInsertedAtoms(
      {
        pkg: packageOf(),
        location: { kind: 'new-paragraph', path: [0] },
        offset: 0,
        allocateBlockId: () => 'm',
      },
      { type: 'paragraph', children: [{ type: 'inlineMath', value: source! }] }
    );
    if (!result.success) throw Error(result.error);
    expect(result.plan.operations[0]).toMatchObject({
      payload: { attributes: { texExpression: source } },
      fallback,
    });
    expect(result.plan.modelLength).toBe(fallback!.length);
  }
});

test('image references stay deferred and ordinary text allocates no control IDs or media parts', () => {
  const pkg = packageOf(),
    before = writeOoxmlPackage(pkg);
  const result = planInsertedAtoms(
    {
      pkg,
      location: { kind: 'new-paragraph', path: [0] },
      offset: 0,
      allocateBlockId: () => {
        throw Error('No control allocation expected');
      },
    },
    paragraphOf(
      'Text ![one](https://example.invalid/one.png) ![two](https://example.invalid/one.png)'
    )
  );
  if (!result.success) throw Error(result.error);
  expect(result.plan.operations.map((op) => op.atom)).toEqual(['image', 'image']);
  expect(
    result.plan.operations.every(
      (op) => op.kind === 'drawing' && !Object.hasOwn(op.resource, 'bytes')
    )
  ).toBe(true);
  expect(writeOoxmlPackage(pkg)).toEqual(before);
});

test('duplicate, invalid, or existing IDs and unsupported atoms refuse without package mutation', () => {
  const initial = packageOf();
  const stored = withPayloadRecord(
    initial,
    { id: 'existing', type: 'mention', version: 1, attributes: { key: 'old', value: 'Old' } },
    'Old'
  );
  if (!stored.ok) throw Error(stored.reason);
  for (const [id, markdown] of [
    ['same', '$a$ $b$'],
    ['bad id', '$a$'],
    ['existing', '$a$'],
    ['fresh', '<unknown inert={throwing()} />'],
  ]) {
    const pkg = stored.pkg,
      before = writeOoxmlPackage(pkg);
    const result = planInsertedAtoms(
      {
        pkg,
        location: { kind: 'retained-paragraph', paragraphId: 'p' },
        offset: 0,
        allocateBlockId: () => id!,
      },
      paragraphOf(markdown!)
    );
    expect(result.success).toBe(false);
    expect(writeOoxmlPackage(pkg)).toEqual(before);
  }
});

test('unsafe coordinates and invalid source characters refuse instead of truncating or normalizing', () => {
  const pkg = packageOf();
  for (const offset of [-1, 0.5, Number.MAX_SAFE_INTEGER]) {
    expect(
      planInsertedAtoms(
        { pkg, location: { kind: 'retained-paragraph', paragraphId: 'p' }, offset },
        paragraphOf('x ![image](asset:one)')
      ).success
    ).toBe(false);
  }
  const invalidAtoms: readonly TargetAstNode[] = [
    { type: 'inlineMath', value: '\u0000' },
    { type: 'mention', username: 'key', displayText: 'split\nlabel' },
    { type: 'image', url: 'asset:one', alt: '\ud800' },
  ];
  for (const atom of invalidAtoms) {
    expect(
      planInsertedAtoms(
        { pkg, location: { kind: 'new-paragraph', path: [0] }, offset: 0 },
        { type: 'paragraph', children: [atom] }
      ).success
    ).toBe(false);
  }
});

test('unknown semantic atom fields and nested source children never disappear into fallback', () => {
  const unknownAtoms: readonly TargetAstNode[] = [
    { type: 'image', url: 'asset:one', width: 123 },
    { type: 'inlineMath', value: 'x', data: { private: 'retain' } },
    { type: 'mention', username: 'ada', children: [{ type: 'text', value: 'Ada', private: true }] },
    {
      type: 'mdxJsxTextElement',
      name: 'source',
      attributes: [],
      children: [{ type: 'mdxJsxTextElement', name: 'widget', attributes: [], children: [] }],
    },
  ];
  for (const atom of unknownAtoms) {
    const result = planInsertedAtoms(
      { pkg: packageOf(), location: { kind: 'new-paragraph', path: [0] }, offset: 0 },
      { type: 'paragraph', children: [atom] }
    );
    expect(result).toMatchObject({ success: false, code: 'unsupported-operation' });
  }
});
