import { expect, test } from 'bun:test';
import { createMathPayload } from '../blocks/math-codec';
import { parseTargetMarkdown, type TargetAstNode } from './parse-target';
import { serializeMath } from './math';
import { domainNodeFromAst } from '../diff/domain-ast';
import { normalizeLogicalBlocks } from '../diff/logical-blocks';
import { insertedPayload } from '../diff/inserted-blocks';

function equation(markdown: string) {
  const read = parseTargetMarkdown(markdown);
  if (!read.success) throw Error(read.error);
  expect(read.target.normalizedAst.children).toHaveLength(1);
  let node = read.target.normalizedAst.children[0]!;
  if (node.type === 'paragraph') {
    const children = node.children as TargetAstNode[];
    expect(children).toHaveLength(1);
    node = children[0]!;
  }
  if (node.type === 'inlineMath' || node.type === 'math')
    return { kind: node.type === 'inlineMath' ? 'inline' : 'block', source: node.value };
  const decoded = domainNodeFromAst(node);
  if (!decoded) throw Error('Expected registered equation syntax');
  return {
    kind: decoded.type === 'inline_equation' ? 'inline' : 'block',
    source: decoded.texExpression,
  };
}

test('ordinary inline and display equations retain standard math syntax', () => {
  expect(serializeMath(createMathPayload('i', 'inline', 'x^2')).text).toBe('$x^2$');
  expect(serializeMath(createMathPayload('b', 'block', '\\frac{x}{y}')).text).toBe(
    '$$\n\\frac{x}{y}\n$$'
  );
});

test('every equation source and kind roundtrip exactly, including empty, multiline and delimiter sources', () => {
  for (const kind of ['inline', 'block'] as const)
    for (const source of [
      '',
      ' ',
      '  x  ',
      '1',
      'true',
      'null',
      '[]',
      '"quoted"',
      'a\nb',
      'a\r\nb\r',
      'a\n\nb\n',
      '\t𝛼\t',
      '$',
      '$$\n$$$',
      '\\frac{x}{y}',
      '<br>',
      '\\text{A<br>B}',
      'a|b',
      'a\\|b',
      '<entity foo="x"> &copy; {inert()}\r\n',
    ]) {
      const payload = createMathPayload('math', kind, source);
      const before = JSON.stringify(payload);
      const serialized = serializeMath(payload);
      expect(equation(serialized.text)).toEqual({ kind, source });
      expect(serialized.properties).toEqual([
        {
          name: 'texExpression',
          property: 'texExpression',
          path: ['texExpression'],
          start: 0,
          end: serialized.text.length,
          valueStart: null,
          valueEnd: null,
        },
      ]);
      expect(JSON.stringify(payload)).toBe(before);
    }
});

test('fallback syntax stays inert, preserves numeric-looking strings, and retains inline logical ownership', () => {
  expect(serializeMath(createMathPayload('i', 'inline', '')).text).toBe(
    '<scio-inline-equation texExpression="" />'
  );
  for (const source of ['1', 'true', 'null', '[]', '"quoted"']) {
    const payload = insertedPayload(
      {
        type: 'mdxJsxTextElement',
        name: 'scio-inline-equation',
        attributes: [{ type: 'mdxJsxAttribute', name: 'texExpression', value: source }],
        children: [],
      },
      'fresh'
    );
    expect(payload?.attributes.texExpression).toBe(source);
  }
  const parsed = parseTargetMarkdown(
    serializeMath(createMathPayload('i', 'inline', 'a\r\nb')).text
  );
  if (!parsed.success) throw Error(parsed.error);
  const logical = normalizeLogicalBlocks(parsed.target);
  if (!logical.success) throw Error(logical.error);
  expect(logical.document.blocks[0]!.kind).toBe('paragraph');
  expect(
    insertedPayload(
      { type: 'mdxJsxFlowElement', name: 'scio-equation', attributes: [], children: [] },
      'bad'
    )
  ).toBeNull();
});

test('inline source pipes remain exact inside GFM cells without adding columns', () => {
  for (const source of ['a|b', 'a\\|b', 'a\\\\|b', 'x\r\n|𝛼']) {
    const serialized = serializeMath(createMathPayload('i', 'inline', source));
    const parsed = parseTargetMarkdown(`| H |\n| - |\n| ${serialized.text} |`);
    if (!parsed.success) throw Error(parsed.error);
    const table = parsed.target.normalizedAst.children[0]!,
      rows = table.children as TargetAstNode[];
    expect(table.type).toBe('table');
    const cells = rows[1]!.children as TargetAstNode[];
    expect(cells).toHaveLength(1);
    const child = (cells[0]!.children as TargetAstNode[])[0]!;
    expect(
      child.type === 'inlineMath' ? child.value : domainNodeFromAst(child)?.texExpression
    ).toBe(source);
  }
});

test('unknown equation metadata, duplicate attributes, and expressions stay explicit and never disappear into payloads', () => {
  for (const text of [
    '<scio-equation texExpression="x" private="keep" />',
    '<scio-equation texExpression="x" texExpression="y" />',
    '<scio-equation texExpression={inert()} />',
    '<scio-equation texExpression="x"><widget /></scio-equation>',
  ]) {
    const parsed = parseTargetMarkdown(text);
    if (!parsed.success) throw Error(parsed.error);
    const node = parsed.target.normalizedAst.children[0]!;
    expect(node.type).toBe('mdxJsxFlowElement');
    expect(insertedPayload(node, 'fresh')).toBeNull();
  }
  for (const node of [
    { type: 'math', value: 'x', meta: 'private-format' },
    { type: 'math', value: 'x', private: true },
    { type: 'inlineMath', value: 'x', data: { private: true } },
  ] as TargetAstNode[])
    expect(insertedPayload(node, 'fresh')).toBeNull();
});

test('ordinary React equation tags remain opaque JSX instead of becoming native math', () => {
  for (const tag of ['equation', 'inline_equation']) {
    const parsed = parseTargetMarkdown(`<${tag} texExpression="x" />`);
    if (!parsed.success) throw Error(parsed.error);
    expect(parsed.target.normalizedAst.children[0]).toMatchObject({
      type: 'mdxJsxFlowElement',
      name: tag,
      attributes: [{ name: 'texExpression', value: 'x' }],
    });
  }
});
