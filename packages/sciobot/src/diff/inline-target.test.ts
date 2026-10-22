import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import fixture from '../__tests__/fixtures/ai-pairs.json';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import {
  compileInlineTarget,
  INLINE_TARGET_LIMITS,
  type CompiledInlineTarget,
} from './inline-target';

const p = (children: readonly TargetAstNode[]): TargetAstNode => ({ type: 'paragraph', children });
const t = (value: string): TargetAstNode => ({ type: 'text', value });
const md = (markdown: string): TargetAstNode => {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  return result.target.ast;
};
const inline = (node: TargetAstNode): CompiledInlineTarget => {
  const result = compileInlineTarget(node);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.target;
};
const first = (markdown: string) => inline((md(markdown).children as TargetAstNode[])[0]!);
const plain = (target: CompiledInlineTarget) =>
  target.tokens
    .filter((token) => token.kind === 'text')
    .map((token) => ({ value: token.value, marks: token.marks }));
function paragraphLike(node: TargetAstNode): TargetAstNode[] {
  if (['paragraph', 'heading', 'tableCell'].includes(node.type)) return [node];
  return Array.isArray(node.children)
    ? (node.children as TargetAstNode[]).flatMap(paragraphLike)
    : [];
}

test('nested Markdown marks compile to canonical semantic sets, literal code, and exact UTF-16 text', () => {
  const target = first('Plain ***both*** ~~removed~~ `x < y` **`code`**.');
  expect(target.plainText).toBe('Plain both removed x < y code.');
  expect(plain(target)).toEqual([
    { value: 'Plain ', marks: [] },
    { value: 'both', marks: ['bold', 'italic'] },
    { value: ' ', marks: [] },
    { value: 'removed', marks: ['strike'] },
    { value: ' ', marks: [] },
    { value: 'x < y', marks: ['code'] },
    { value: ' ', marks: [] },
    { value: 'code', marks: ['bold', 'code'] },
    { value: '.', marks: [] },
  ]);
  expect(plain(first('*__both__*'))).toEqual(plain(first('***both***')));
  expect(target.hyperlinks).toEqual([]);
  expect(target.tokens[1]!.path).toEqual([1, 0, 0]);
});

test('manual breaks, tabs, hyphen model tokens, escapes, entities, Unicode, and authored NBSP stay exact', () => {
  expect(first('First  \nSecond\\* &amp; 🧑🏽‍💻 é').plainText).toBe('First\nSecond* & 🧑🏽‍💻 é');
  const target = inline(p([t('A\tB\u001eC\u001fD '), { type: 'break' }, t('👩‍👩‍👧‍👦')]));
  expect(target.plainText).toBe('A\tB\u001eC\u001fD \n👩‍👩‍👧‍👦');
  expect(inline(p([t('\u00a0')])).plainText).toBe('\u00a0');
  expect(inline(p([])).plainText).toBe('');
  expect(inline(p([t('')])).tokens).toMatchObject([{ kind: 'text', value: '', marks: [] }]);
});

test('native HTML emphasis fallback and empty br map to the same semantic inline data', () => {
  const target = first('A<strong>B<em>C</em></strong><del>D</del><br />E');
  expect(target.plainText).toBe('ABCD\nE');
  expect(plain(target)).toEqual([
    { value: 'A', marks: [] },
    { value: 'B', marks: ['bold'] },
    { value: 'C', marks: ['bold', 'italic'] },
    { value: 'D', marks: ['strike'] },
    { value: '\n', marks: [] },
    { value: 'E', marks: [] },
  ]);
  expect(plain(first('A<b>B<i>C</i></b><s>D</s>'))).toEqual(
    plain(first('A<strong>B<em>C</em></strong><del>D</del>'))
  );
});

test('hyperlink owners retain separate identity, titles, native safe targets, and mark ownership', () => {
  const target = first(
    '[**One**](https://example.com "Title")[Two](https://example.com) [Go](#bookmark)'
  );
  expect(target.plainText).toBe('OneTwo Go');
  expect(target.hyperlinks).toEqual([
    { path: [0], target: { kind: 'external', url: 'https://example.com' }, title: 'Title' },
    { path: [1], target: { kind: 'external', url: 'https://example.com' }, title: null },
    { path: [3], target: { kind: 'bookmark', anchor: 'bookmark' }, title: null },
  ]);
  expect(target.tokens.map((token) => token.hyperlink)).toEqual([[0], [1], null, [3]]);
  expect(target.tokens[0]).toMatchObject({ value: 'One', marks: ['bold'], path: [0, 0, 0] });
  const empty = inline(p([{ type: 'link', url: 'https://example.com', children: [] }]));
  expect(empty.plainText).toBe('');
  expect(empty.tokens).toEqual([]);
  expect(empty.hyperlinks).toHaveLength(1);
});

test('unsafe link targets refuse through native authorability without losing inert MDX source', () => {
  for (const url of [
    'javascript:alert(1)',
    'data:text/html,hello',
    'file:///secret',
    'relative/path',
    '',
    '#',
  ]) {
    const result = compileInlineTarget(p([{ type: 'link', url, children: [t('Click')] }]));
    expect(result.success).toBe(false);
    expect(Object.hasOwn(result, 'target')).toBe(false);
  }
  for (const url of [
    'https://example.org/a?q=x&b=2',
    'mailto:hello@example.org',
    'tel:+420123456789',
    'ftp://example.org/a',
  ])
    expect(inline(p([{ type: 'link', url, children: [t('Safe')] }])).hyperlinks[0]!.target).toEqual(
      { kind: 'external', url }
    );
  const ast = {
    type: 'mdxJsxTextElement',
    name: 'Widget',
    attributes: [{ type: 'mdxJsxAttribute', name: 'href', value: 'javascript:alert(1)' }],
    children: [],
  } as TargetAstNode;
  expect(inline(p([ast])).tokens).toMatchObject([{ kind: 'atom', ast }]);
});

test('images, mentions, math, references, unknown nodes, and MDX remain complete atoms with outer marks', () => {
  const atoms: TargetAstNode[] = [
    { type: 'image', url: 'https://example.com/pic.png', alt: 'Asset', title: 'Keep' },
    { type: 'mention', username: 'person', displayText: 'Person', children: [t('Person')] },
    { type: 'inlineMath', value: 'x^2' },
    {
      type: 'linkReference',
      identifier: 'ref',
      referenceType: 'full',
      label: 'Ref',
      children: [t('Ref')],
    },
    { type: 'footnoteReference', identifier: 'note', label: 'note' },
    {
      type: 'futureInline',
      value: 'do not flatten',
      future: { type: 'payload', position: { keep: true } },
      children: [t('Fallback')],
    },
    {
      type: 'mdxTextExpression',
      value: 'globalThis.__inlineTargetExecuted = true',
      data: { estree: { raw: 'inert' } },
    },
    {
      type: 'mdxJsxTextElement',
      name: 'Widget',
      attributes: [
        { type: 'mdxJsxExpressionAttribute', value: '...danger()' },
        {
          type: 'mdxJsxAttribute',
          name: 'value',
          value: { type: 'mdxJsxAttributeValueExpression', value: 'danger()' },
        },
      ],
      children: [{ type: 'mdxTextExpression', value: 'danger()' }],
    },
  ];
  const target = inline(p([{ type: 'strong', children: atoms }]));
  expect(target.plainText).toBeNull();
  expect(target.tokens).toHaveLength(atoms.length);
  target.tokens.forEach((token, index) => {
    expect(token.kind).toBe('atom');
    expect(token.marks).toEqual(['bold']);
    expect(token.path).toEqual([0, index]);
    if (token.kind === 'atom') expect(token.ast).toEqual(atoms[index]);
  });
  expect(target.tokens.some((token) => token.kind === 'text')).toBe(false);
  expect((globalThis as Record<string, unknown>).__inlineTargetExecuted).toBeUndefined();
});

test('known tags with attributes or unknown metadata stay atomic rather than dropping meaning', () => {
  const nodes: TargetAstNode[] = [
    { type: 'strong', children: [t('Keep')], future: { keep: true } },
    { type: 'text', value: 'Annotated', data: { future: 'keep' } },
    {
      type: 'mdxJsxTextElement',
      name: 'strong',
      attributes: [{ type: 'mdxJsxAttribute', name: 'className', value: 'custom' }],
      children: [t('Styled')],
    },
    {
      type: 'mdxJsxTextElement',
      name: 'br',
      attributes: [{ type: 'mdxJsxAttribute', name: 'data-extra', value: 'retain' }],
      children: [],
    },
  ];
  const root = {
    ...p(nodes),
    data: { blockMetadata: { keep: true } },
    position: { start: { offset: 3 }, end: { offset: 100 } },
  };
  const target = inline(root);
  expect(target.ast).toEqual(root);
  expect(target.plainText).toBeNull();
  target.tokens.forEach((token, index) =>
    expect(token).toMatchObject({ kind: 'atom', ast: nodes[index] })
  );
  expect(inline(p([{ type: 'strong', children: [t('Ordinary')], data: {} }])).plainText).toBe(
    'Ordinary'
  );
});

test('atomic content inside a hyperlink keeps its owner and never becomes fallback text', () => {
  const target = inline(
    p([
      {
        type: 'link',
        url: 'https://example.com',
        children: [t('Before'), { type: 'image', url: 'image.png', alt: 'Fallback' }, t('After')],
      },
    ])
  );
  expect(target.plainText).toBeNull();
  expect(target.tokens.map((token) => token.kind)).toEqual(['text', 'atom', 'text']);
  expect(target.tokens.map((token) => token.hyperlink)).toEqual([[0], [0], [0]]);
  expect(target.hyperlinks).toHaveLength(1);
});

test('opaque prototype-shaped keys stay data without affecting object prototypes', () => {
  const data = JSON.parse(
    '{"__proto__":{"polluted":"no"},"constructor":{"prototype":{"polluted":"no"}},"prototype":{"keep":true}}'
  );
  const target = inline(p([{ type: 'futureInline', data }]));
  expect(target.tokens[0]).toMatchObject({ kind: 'atom', ast: { type: 'futureInline', data } });
  expect(Object.getPrototypeOf(target.ast)).toBe(Object.prototype);
  expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  if (target.tokens[0]!.kind === 'atom')
    expect(Object.hasOwn(target.tokens[0]!.ast.data as object, '__proto__')).toBe(true);
});

const expected: Record<string, { before: (string | null)[]; after: (string | null)[] }> = {
  'formatting and lists': {
    before: ['Bold and italic.', 'First', 'Second', 'Apple', 'Pear'],
    after: ['Bold and updated italic.', 'First', 'Added', 'Second', 'Apple', 'Pear'],
  },
  'tables and links': {
    before: ['A', 'B', 'One', 'Two', 'Link'],
    after: ['A', 'B', 'One', 'Changed', 'Three', 'Four', 'Link'],
  },
  'section insertion': {
    before: ['Opening.', 'Closing.'],
    after: ['Opening.', 'New section', 'First addition.', 'Second addition.', 'Closing.'],
  },
  'nested MDX and attributes': {
    before: [null, 'First question.', 'Last question.'],
    after: [null, 'A better question.', 'Last question.'],
  },
  'source content': {
    before: [null, 'First question.', 'Last question.'],
    after: [null, 'First question.', 'Final reflection.'],
  },
  deletion: { before: ['Keep.', 'Remove.', 'Tail.'], after: ['Keep.', 'Tail.'] },
  'complete rewrite': {
    before: ['Old document.', 'Old end.'],
    after: ['Replacement', 'New body.', 'New ending.'],
  },
};
test.each(fixture.cases)('C01 inline content compiles explicitly on both sides: $name', (value) => {
  const before = paragraphLike(md(value.beforeMarkdown)).map(inline),
    after = paragraphLike(md(value.afterMarkdown)).map(inline);
  expect(before.map((target) => target.plainText)).toEqual(expected[value.name]!.before);
  expect(after.map((target) => target.plainText)).toEqual(expected[value.name]!.after);
  for (const target of [...before, ...after]) expect(target.ast.children).toBeDefined();
  if (value.name === 'formatting and lists')
    expect(plain(after[0]!)).toEqual([
      { value: 'Bold', marks: ['bold'] },
      { value: ' and ', marks: [] },
      { value: 'updated italic', marks: ['italic'] },
      { value: '.', marks: [] },
    ]);
  if (value.name === 'tables and links') {
    expect(before.at(-1)!.hyperlinks[0]!.target).toEqual({
      kind: 'external',
      url: 'https://example.com',
    });
    expect(after.at(-1)!.hyperlinks[0]!.target).toEqual({
      kind: 'external',
      url: 'https://example.org',
    });
  }
  if (value.name === 'nested MDX and attributes' || value.name === 'source content')
    expect(after[0]!.tokens[0]).toMatchObject({
      kind: 'atom',
      ast: {
        type: 'mdxJsxTextElement',
        name: 'lesson_title',
        attributes: [{ type: 'mdxJsxAttribute', name: 'grade', value: '7' }],
        children: [{ type: 'text', value: 'Lesson' }],
      },
    });
});

test('detached outputs freeze every field without freezing or mutating caller data', () => {
  const input = p([
    { type: 'strong', children: [t('Keep')] },
    { type: 'future', children: [t('Opaque')], data: { position: { keep: true } } },
  ]);
  const before = JSON.stringify(input),
    target = inline(input);
  function frozen(value: unknown): boolean {
    return (
      value === null ||
      typeof value !== 'object' ||
      (Object.isFrozen(value) && Object.values(value).every(frozen))
    );
  }
  expect(frozen(target)).toBe(true);
  expect(Object.isFrozen(input)).toBe(false);
  expect(JSON.stringify(input)).toBe(before);
  (input.children as Record<string, unknown>[])[0]!.children = [t('Changed')];
  expect(target.tokens[0]).toMatchObject({ kind: 'text', value: 'Keep' });
  expect(target.ast).not.toEqual(input);
});

test('malformed known nodes, blocks, nested links, and invalid text refuse atomically', () => {
  const invalid: unknown[] = [
    null,
    [],
    { type: 'root', children: [] },
    { type: { toString: null, valueOf: null }, children: [] },
    { ...p([]), kind: 'paragraph' },
    p([{ type: 'text', value: 4 }]),
    p([{ type: 'strong' }]),
    p([{ type: 'heading', depth: 2, children: [t('Block')] }]),
    p([
      {
        type: 'link',
        url: 'https://example.com',
        children: [{ type: 'link', url: 'https://example.org', children: [t('Nested')] }],
      },
    ]),
    p([{ type: 'link', url: 'https://example.com', title: 'bad\u001f', children: [] }]),
    p([t('\u0000')]),
    p([t('\r')]),
    p([t('\uD800')]),
    p([t('\uDC00')]),
  ];
  for (const value of invalid) {
    const result = compileInlineTarget(value as TargetAstNode);
    expect(result.success).toBe(false);
    expect(Object.hasOwn(result, 'target')).toBe(false);
  }
});

test('getters, toJSON, hidden/symbol properties, cycles, sparse arrays, and resource excess never evaluate', () => {
  let calls = 0;
  const getter = Object.defineProperty(p([]), 'children', {
      enumerable: true,
      get() {
        calls++;
        return [];
      },
    }),
    toJSON = {
      ...p([]),
      toJSON() {
        calls++;
        return p([]);
      },
    },
    hidden = Object.defineProperty(p([]), 'hidden', { value: true }),
    symbol = { ...p([]), [Symbol('hidden')]: true };
  const cycle: Record<string, unknown> = { type: 'paragraph', children: [] };
  cycle.children = [cycle];
  let deep: TargetAstNode = t('Leaf');
  for (let i = 0; i < INLINE_TARGET_LIMITS.maxDepth; i++)
    deep = { type: 'strong', children: [deep] };
  const values = [
    getter,
    toJSON,
    hidden,
    symbol,
    cycle,
    p(new Array(2)),
    p([deep]),
    p([t('x'.repeat(INLINE_TARGET_LIMITS.maxBytes))]),
    p([t('🧑'.repeat(INLINE_TARGET_LIMITS.maxBytes / 4))]),
  ];
  for (const value of values)
    expect(compileInlineTarget(value as TargetAstNode).success).toBe(false);
  expect(calls).toBe(0);
  const excess = compileInlineTarget(p(new Array(INLINE_TARGET_LIMITS.maxNodes + 1).fill(t(''))));
  expect(excess.success).toBe(false);
  if (!excess.success) expect(excess.error).toBe('node-limit');
});

test('compilation preserves live source bytes, revision, selection, undo history, and imported XML', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(
      '<w:p><w:pPr><keep:extra xmlns:keep="urn:keep" keep:attr="retain"/></w:pPr><w:r><w:t>Original.</w:t></w:r></w:p>'
    ),
  });
  try {
    const surface = editor.surface!,
      session = surface.session;
    surface.applyAutomationOps(() => [
      { op: 'insertText', paragraphId: session.paragraphIds()[0]!, offset: 0, text: 'User ' },
    ]);
    const revision = session.packageRevision(),
      bytes = writeOoxmlPackage(session.currentPackage()),
      xml = canonicalOoxmlFingerprint(session.part()),
      selection = surface.state().selection,
      history = [session.canUndo(), session.canRedo()];
    expect(first('A **new** [link](https://example.org).').plainText).toBe('A new link.');
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(xml);
    expect(session.packageRevision()).toBe(revision);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});
