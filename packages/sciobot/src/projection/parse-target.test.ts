import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import fixture from '../__tests__/fixtures/target-parsing.json';
import type { JsonValue } from '../blocks/payload-schema';
import {
  parseTargetMarkdown,
  type ParsedTargetMarkdown,
  type TargetAstNode,
  type TargetParseLimits,
} from './parse-target';

function parsed(markdown: string): ParsedTargetMarkdown {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  return result.target;
}
function nodesOfType(node: TargetAstNode, type: string): TargetAstNode[] {
  const children = Array.isArray(node.children) ? (node.children as TargetAstNode[]) : [];
  return [
    ...(node.type === type ? [node] : []),
    ...children.flatMap((child) => nodesOfType(child, type)),
  ];
}
function allFrozen(value: JsonValue): boolean {
  return (
    value === null ||
    typeof value !== 'object' ||
    (Object.isFrozen(value) && Object.values(value).every(allFrozen))
  );
}

test.each(fixture.cases)('target parsing matches the production pipeline: $name', (value) => {
  const first = parsed(value.input);
  const second = parsed(value.input);
  expect(first.sourceMarkdown).toBe(value.input);
  expect(first.markdown).toBe(value.markdown);
  expect<unknown>(first.normalizedAst).toEqual(value.normalizedAst);
  expect(first).toEqual(second);
  expect(first === second).toBe(false);
  expect(allFrozen(first.ast)).toBe(true);
  expect(allFrozen(first.normalizedAst)).toBe(true);
});

test.each(fixture.failures)(
  'malformed MDX returns the production syntax location: $input',
  (value) => {
    const result = parseTargetMarkdown(value.input);
    const { input: _input, ...issue } = value;
    expect(result).toMatchObject({
      success: false,
      code: 'invalid-input',
      issues: [
        {
          code: 'syntax',
          ...issue,
        },
      ],
    });
    expect(Object.hasOwn(result, 'target')).toBe(false);
  }
);

test('C01 targets retain explicit native formatting, table, link, and domain semantics', () => {
  const reference = (name: string) => fixture.cases.find((value) => value.name === name)!.input;
  const formatted = parsed(reference('formatting and lists:after')).normalizedAst;
  expect(nodesOfType(formatted, 'strong')).toHaveLength(1);
  expect(nodesOfType(formatted, 'emphasis')[0]?.children).toEqual([
    { type: 'text', value: 'updated italic' },
  ]);
  expect(nodesOfType(formatted, 'list').map((node) => node.ordered)).toEqual([true, false]);
  expect(nodesOfType(formatted, 'listItem')).toHaveLength(5);
  const table = parsed(reference('tables and links:after')).normalizedAst;
  expect(nodesOfType(table, 'tableRow')).toHaveLength(3);
  expect(nodesOfType(table, 'tableCell')).toHaveLength(6);
  expect(nodesOfType(table, 'link')[0]?.url).toBe('https://example.org');
  const nested = parsed(reference('nested MDX and attributes:after')).normalizedAst;
  const activity = nodesOfType(nested, 'mdxJsxFlowElement').find(
    (node) => node.name === 'activity'
  )!;
  expect(activity.attributes).toEqual([
    { type: 'mdxJsxAttribute', name: 'name', value: 'Warmup' },
    { type: 'mdxJsxAttribute', name: 'duration', value: '15' },
  ]);
  expect(nodesOfType(activity, 'text').map((node) => node.value)).toContain('A better question.');
});

test('untransformed text retains exact parser positions inside GFM cells and escaped text', () => {
  const markdown = '| A | B |\n| --- | --- |\n| One | Two |\n\nA \\*literal\\* value.',
    target = parsed(markdown),
    texts = nodesOfType(target.ast, 'text');
  for (const value of ['A', 'B', 'One', 'Two', 'A *literal* value.']) {
    const node = texts.find((item) => item.value === value)!;
    const position = node.position as { start: { offset: number }; end: { offset: number } };
    expect(position).toBeDefined();
    expect(target.markdown.slice(position.start.offset, position.end.offset)).toBe(
      value === 'A *literal* value.' ? 'A \\*literal\\* value.' : value
    );
  }
  expect(
    nodesOfType(target.normalizedAst, 'text').every((node) => !Object.hasOwn(node, 'position'))
  ).toBe(true);
});

test('source positions address normalized Markdown and symmetric comparison ignores positions', () => {
  const target = parsed('Label:\n6. First\n7. Second');
  const list = nodesOfType(target.ast, 'list')[0]!;
  const position = list.position as { start: { offset: number }; end: { offset: number } };
  expect(target.markdown.slice(position.start.offset, position.end.offset)).toBe(
    '6. First\n7. Second'
  );
  expect(target.markdown).toBe('Label:\n\n6. First\n7. Second');
  expect(parsed('**Bold**').normalizedAst).toEqual(parsed('__Bold__').normalizedAst);
  expect(parsed('Before\n\nAfter').normalizedAst).toEqual(
    parsed('Before\n\n\nAfter').normalizedAst
  );
  expect(parsed('3. Third\n4. Fourth').normalizedAst).not.toEqual(
    parsed('1. Third\n2. Fourth').normalizedAst
  );
});

test('normalization keeps inline images and code while grouping only image-only block runs', () => {
  const target = parsed(
    '![One](one.png)\n\n![Two](two.png)\n\nText ![Inline](inline.png).\n\n`![Code](code.png)`'
  );
  const galleries = nodesOfType(target.normalizedAst, 'mdxJsxFlowElement').filter(
    (node) => node.name === 'image_gallery'
  );
  expect(galleries).toHaveLength(1);
  const attributes = galleries[0]!.attributes as {
    name: string;
    value: { value?: string } | string;
  }[];
  const imageAttribute = attributes.find((attribute) => attribute.name === 'images')!;
  expect(JSON.parse((imageAttribute.value as { value: string }).value)).toEqual([
    { url: 'one.png', caption: 'One' },
    { url: 'two.png', caption: 'Two' },
  ]);
  expect(nodesOfType(target.normalizedAst, 'image').map((node) => node.url)).toEqual([
    'inline.png',
  ]);
  expect(nodesOfType(target.normalizedAst, 'inlineCode')[0]?.value).toBe('![Code](code.png)');
});

test('nested opaque MDX retains its entire AST and expression source without evaluation', () => {
  const scope = globalThis as typeof globalThis & { __targetProbe?: () => void };
  let calls = 0;
  scope.__targetProbe = () => {
    calls += 1;
  };
  try {
    const target = parsed(
      '<suggestions>\n\n<suggestion message="Continue">{globalThis.__targetProbe()}</suggestion>\n\n</suggestions>'
    );
    const tags = [
      ...nodesOfType(target.ast, 'mdxJsxFlowElement'),
      ...nodesOfType(target.ast, 'mdxJsxTextElement'),
    ];
    expect(tags.map((node) => node.name).sort()).toEqual(['suggestion', 'suggestions']);
    const expressions = [
      ...nodesOfType(target.ast, 'mdxFlowExpression'),
      ...nodesOfType(target.ast, 'mdxTextExpression'),
    ];
    expect(expressions[0]?.value).toBe('globalThis.__targetProbe()');
    expect(calls).toBe(0);
    // Compiler-only ESTree caches can hold non-JSON values; the expression source remains lossless.
    expect(JSON.stringify(parsed('<future value={123n} expression={/foo/u} />').ast)).toContain(
      '123n'
    );
    expect(JSON.stringify(target.ast)).not.toContain('"estree"');
  } finally {
    delete scope.__targetProbe;
  }
});

test('input, byte, node, and depth limits refuse without a partial target', () => {
  const samples: [unknown, Partial<TargetParseLimits>, string][] = [
    [42, {}, 'invalid-input'],
    ['😀😀', { maxBytes: 7 }, 'byte-limit'],
    ['One paragraph', { maxNodes: 2 }, 'node-limit'],
    ['**Bold**', { maxDepth: 2 }, 'depth-limit'],
    ['Text', { maxBytes: 0 }, 'invalid-input'],
  ];
  for (const [input, limits, code] of samples) {
    const result = parseTargetMarkdown(input as string, limits);
    expect(result).toMatchObject({ success: false, code: 'invalid-input', issues: [{ code }] });
    expect(Object.hasOwn(result, 'target')).toBe(false);
  }
});

test('success and syntax failures preserve existing native bytes, revisions, selection, and history', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture('<w:p><w:r><w:t>Kept 😀</w:t></w:r></w:p>'),
  });
  try {
    const surface = editor.surface!;
    const session = surface.session;
    const paragraphId = session.paragraphIds()[0]!;
    surface.setSelection({ anchor: { paragraphId, offset: 0 }, head: { paragraphId, offset: 0 } });
    surface.insertPlainText('User ');
    surface.flushPendingInput();
    const before = [...session.currentPackage().parts].map(([name, part]) => [
      name,
      canonicalOoxmlFingerprint(part),
    ]);
    const bytes = writeOoxmlPackage(session.currentPackage());
    const revision = session.packageRevision();
    const selection = JSON.stringify(surface.state().selection);
    const history = [session.canUndo(), session.canRedo()];
    expect(history).toEqual([true, false]);
    for (const input of ['Replacement.', '<activity>\n\nBroken.', '# Heading\n\n**New**']) {
      const result = parseTargetMarkdown(input);
      expect(result.success).toBe(!input.includes('Broken'));
      expect(
        [...session.currentPackage().parts].map(([name, part]) => [
          name,
          canonicalOoxmlFingerprint(part),
        ])
      ).toEqual(before);
      expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
      expect(session.packageRevision()).toBe(revision);
      expect(JSON.stringify(surface.state().selection)).toBe(selection);
      expect([session.canUndo(), session.canRedo()]).toEqual(history);
    }
  } finally {
    editor.destroy();
  }
});
