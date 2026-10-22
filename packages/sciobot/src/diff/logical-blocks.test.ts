import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import pairs from '../__tests__/fixtures/ai-pairs.json';
import reference from '../__tests__/fixtures/logical-blocks.json';
import { DOMAIN_BLOCK_TAGS } from '../projection/block-tags';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { compileInlineTarget } from './inline-target';
import {
  LOGICAL_BLOCK_LIMITS,
  normalizeLogicalBlocks,
  type LogicalBlock,
  type LogicalBlockDocument,
  type LogicalGroup,
  type LogicalParagraph,
} from './logical-blocks';

const t = (value: string): TargetAstNode => ({ type: 'text', value });
const p = (value: string): TargetAstNode => ({ type: 'paragraph', children: [t(value)] });
const raw = (children: readonly TargetAstNode[]) => ({
  ast: { type: 'root' as const, children },
  markdown: '',
});
function parsed(markdown: string) {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  return result.target;
}
function logical(markdown: string): LogicalBlockDocument {
  const result = normalizeLogicalBlocks(parsed(markdown));
  if (!result.success) throw new Error(result.error);
  return result.document;
}
function tree(children: readonly TargetAstNode[]): LogicalBlockDocument {
  const result = normalizeLogicalBlocks(raw(children));
  if (!result.success) throw new Error(result.error);
  return result.document;
}
function paragraph(block: LogicalBlock): LogicalParagraph {
  if (block.kind !== 'paragraph') throw new Error('Expected paragraph.');
  return block;
}
function group(block: LogicalBlock): LogicalGroup {
  if (!('children' in block)) throw new Error('Expected group.');
  return block;
}
function text(block: LogicalBlock): string | null {
  const result = compileInlineTarget(paragraph(block).ast);
  if (!result.success) throw new Error(result.error);
  return result.target.plainText;
}
function shape(block: LogicalBlock): unknown {
  if (block.kind === 'paragraph')
    return [block.ast.type === 'heading' ? `h${block.ast.depth}` : 'p', text(block)];
  if ('children' in block) return [block.domainType ?? block.kind, block.children.map(shape)];
  return [block.kind, block.ast.name ?? block.ast.type];
}
const shapes = (markdown: string) => logical(markdown).blocks.map(shape);
const para = (value: string | null) => ['p', value];
const domain = (type: string, children: unknown[]) => [type, children];
const tableShape = (rows: string[][]) => [
  'table',
  rows.map((row) => ['row', row.map((cell) => ['cell', [para(cell)]])]),
];
const lesson = (question: string, tail: string) => [
  domain('lesson_title', [para('Lesson')]),
  domain('lesson_phase', [domain('lesson_activity', [para(question), para(null), para(tail)])]),
];

test('all seven C01 pairs use one normalizer with explicit before and after block streams', () => {
  const expected = [
    {
      name: 'formatting and lists',
      before: ['Bold and italic.', 'First', 'Second', 'Apple', 'Pear'].map(para),
      after: ['Bold and updated italic.', 'First', 'Added', 'Second', 'Apple', 'Pear'].map(para),
    },
    {
      name: 'tables and links',
      before: [
        tableShape([
          ['A', 'B'],
          ['One', 'Two'],
        ]),
        para('Link'),
      ],
      after: [
        tableShape([
          ['A', 'B'],
          ['One', 'Changed'],
          ['Three', 'Four'],
        ]),
        para('Link'),
      ],
    },
    {
      name: 'section insertion',
      before: [para('Opening.'), para('Closing.')],
      after: [
        para('Opening.'),
        ['h2', 'New section'],
        para('First addition.'),
        para('Second addition.'),
        para('Closing.'),
      ],
    },
    {
      name: 'nested MDX and attributes',
      before: lesson('First question.', 'Last question.'),
      after: lesson('A better question.', 'Last question.'),
    },
    {
      name: 'source content',
      before: lesson('First question.', 'Last question.'),
      after: lesson('First question.', 'Final reflection.'),
    },
    {
      name: 'deletion',
      before: ['Keep.', 'Remove.', 'Tail.'].map(para),
      after: ['Keep.', 'Tail.'].map(para),
    },
    {
      name: 'complete rewrite',
      before: ['Old document.', 'Old end.'].map(para),
      after: [['h1', 'Replacement'], para('New body.'), para('New ending.')],
    },
  ];
  expect(pairs.cases.map((entry) => entry.name)).toEqual(expected.map((entry) => entry.name));
  pairs.cases.forEach((entry, index) => {
    expect(shapes(entry.beforeMarkdown)).toEqual(expected[index]!.before);
    expect(shapes(entry.afterMarkdown)).toEqual(expected[index]!.after);
    for (const markdown of [entry.beforeMarkdown, entry.afterMarkdown]) {
      const target = parsed(markdown),
        result = normalizeLogicalBlocks(target);
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.document.ast).toEqual(target.ast);
        expect(result.document.markdown).toBe(target.markdown);
      }
    }
  });
  const lists = logical(pairs.cases[0]!.afterMarkdown).blocks;
  expect(lists.map((block) => block.list?.marker ?? null)).toEqual([
    null,
    'ordered',
    'ordered',
    'ordered',
    'bullet',
    'bullet',
  ]);
  expect(lists.map((block) => block.list?.start ?? null)).toEqual([null, 1, 1, 1, null, null]);
  expect(lists[1]!.list!.group).toEqual(lists[3]!.list!.group);
  expect(lists[1]!.list!.item).not.toEqual(lists[2]!.list!.item);
});

test('C01 metadata, marks, source chips, and changed link targets retain complete meaning', () => {
  const before = logical(pairs.cases[3]!.beforeMarkdown).blocks,
    after = logical(pairs.cases[3]!.afterMarkdown).blocks;
  const attr = (block: LogicalBlock, name: string) =>
    (group(block).header.attributes as Record<string, unknown>[]).find(
      (entry) => entry.name === name
    )?.value;
  expect(attr(before[0]!, 'grade')).toBe('7');
  expect(attr(before[1]!, 'type')).toBe('evocation');
  const activityBefore = group(before[1]!).children[0]!,
    activityAfter = group(after[1]!).children[0]!;
  expect(attr(activityBefore, 'duration')).toBe('10');
  expect(attr(activityAfter, 'duration')).toBe('15');
  const chip = paragraph(group(activityBefore).children[1]!);
  expect(chip.ast.children).toMatchObject([{ type: 'mdxJsxTextElement', name: 'source' }]);
  expect(text(chip)).toBeNull();
  for (const side of ['beforeMarkdown', 'afterMarkdown'] as const) {
    const listEntry = paragraph(logical(pairs.cases[0]![side]).blocks[0]!);
    const inline = compileInlineTarget(listEntry.ast);
    expect(inline.success).toBe(true);
    if (inline.success)
      expect(
        inline.target.tokens.filter((token) => token.kind === 'text').map((token) => token.marks)
      ).toEqual([['bold'], [], ['italic'], []]);
    const link = compileInlineTarget(paragraph(logical(pairs.cases[1]![side]).blocks[1]!).ast);
    expect(link.success).toBe(true);
    if (link.success)
      expect(link.target.hyperlinks[0]!.target).toEqual({
        kind: 'external',
        url: side === 'beforeMarkdown' ? 'https://example.com' : 'https://example.org',
      });
  }
});

type SlateNode = { type?: string; text?: string; children?: SlateNode[]; [key: string]: unknown };
function slateShape(node: SlateNode): unknown {
  if (node.type === 'p') {
    const values = node.children ?? [];
    const plain = (nodes: SlateNode[]): string =>
      nodes.map((child) => child.text ?? plain(child.children ?? [])).join('');
    return para(values.some((child) => child.type === 'source') ? null : plain(values));
  }
  return [node.type, (node.children ?? []).map(slateShape)];
}
test('independently executed application fixtures establish title, source, and recap normalization', () => {
  const expected = [
    [domain('lesson_title', [para('Lesson')])],
    [domain('lesson_title', [para('Lesson')])],
    [para(null)],
    [domain('source_excerpt', [para('Quoted words.')])],
    [para(null)],
    [domain('recap_goals', [para('First goal'), para('Second goal')])],
    [domain('recap_goals', [para('First goal'), para('Second goal')])],
    [domain('recap_curriculum', [para('One'), para('Two')])],
    [domain('recap_materials', [para('One'), para('Two')])],
    [domain('recap_adaptations', [para('One'), para('Two')])],
    [domain('lesson_activity', [para('One'), para('Two')])],
    ['First', 'Continuation.', 'Child', 'Second'].map(para),
  ];
  expect(reference.schemaVersion).toBe(1);
  expect(reference.cases).toHaveLength(12);
  expect(reference.source.files['src/components/lesson-plan/md-rules.ts']).toBe(
    pairs.sources.application.files['src/components/lesson-plan/md-rules.ts']
  );
  reference.cases.forEach((entry, index) => {
    expect((entry.normalized as SlateNode[]).map(slateShape)).toEqual(expected[index]);
    expect(shapes(entry.markdown)).toEqual(expected[index]);
  });
  // This reference asserts text order only for nested lists. Native list membership uses AST metadata.
  // The installed Plate pipeline collapses nesting and leaks the child marker onto the final item.
  const continuation = reference.cases.find((entry) => entry.name === 'list continuation')!;
  expect(
    (continuation.normalized as SlateNode[]).map((entry) => entry.listStyleType ?? null)
  ).toEqual(['decimal', null, 'disc', 'disc']);
  expect(logical(continuation.markdown).blocks.map((block) => block.list?.marker)).toEqual([
    'ordered',
    'ordered',
    'bullet',
    'ordered',
  ]);
});

test('inline and flow containers have equal matching hints but retain exact original forms and origins', () => {
  const inline = logical(reference.cases[0]!.markdown),
    flow = logical(reference.cases[1]!.markdown);
  expect(inline.blocks[0]!.key).toBe(flow.blocks[0]!.key);
  expect(inline.ast).not.toEqual(flow.ast);
  expect(inline.blocks[0]!.path).toEqual([0, 0]);
  expect(flow.blocks[0]!.path).toEqual([0]);
  const child = paragraph(group(inline.blocks[0]!).children[0]!);
  expect(child.synthetic).toBe(true);
  expect(child.inlinePaths).toEqual([[0, 0, 0]]);
  expect(paragraph(group(flow.blocks[0]!).children[0]!).synthetic).toBe(false);
  const annotated = tree([
    {
      ...p(''),
      children: [
        { type: 'mdxJsxTextElement', name: 'lesson_title', attributes: [], children: [t('Title')] },
      ],
      data: { keep: true },
    },
  ]);
  expect(annotated.blocks[0]!.kind).toBe('paragraph');
  expect(paragraph(annotated.blocks[0]!).ast.data).toEqual({ keep: true });
});

test('empty flow source chips and empty paired source retain atom attributes and original AST', () => {
  for (const name of ['standalone source', 'empty paired source']) {
    const entry = reference.cases.find((candidate) => candidate.name === name)!,
      result = logical(entry.markdown),
      block = paragraph(result.blocks[0]!);
    expect(text(block)).toBeNull();
    const atom = (block.ast.children as TargetAstNode[])[0]!;
    expect(atom.name).toBe('source');
    expect(atom.attributes).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ name: 'title', value: 'Book' }),
        expect.objectContaining({ name: 'url', value: 'https://example.org' }),
      ])
    );
    expect(result.ast).toEqual(parsed(entry.markdown).ast);
  }
  const excerpt = group(logical(reference.cases[3]!.markdown).blocks[0]!);
  expect(excerpt.domainType).toBe('source_excerpt');
  expect(excerpt.children.map(text)).toEqual(['Quoted words.']);
});

test('recap list suppression preserves checked state and applies only inside the owning container', () => {
  const result = logical(
    '<recap_goals>\n\n- [x] Done\n- [ ] Todo\n\n<activity>\n\n3. Keep list\n\n</activity>\n\n</recap_goals>\n\n4. Outside'
  );
  const recap = group(result.blocks[0]!);
  expect(recap.children.slice(0, 2).map((block) => block.list)).toMatchObject([
    { suppressed: true, checked: true, marker: 'bullet' },
    { suppressed: true, checked: false, marker: 'bullet' },
  ]);
  expect(group(recap.children[2]!).children[0]!.list).toMatchObject({
    suppressed: false,
    marker: 'ordered',
    start: 3,
  });
  expect(result.blocks[1]!.list).toMatchObject({ suppressed: false, start: 4 });
  expect(
    group(logical('<preparation_recap>\n\n- Keep list\n\n</preparation_recap>').blocks[0]!)
      .children[0]!.list!.suppressed
  ).toBe(false);
  for (const tag of ['recap_curriculum', 'recap_materials', 'recap_adaptations']) {
    const children = group(logical(`<${tag}>\n\n- One\n- Two\n\n</${tag}>`).blocks[0]!).children;
    expect(children.map((block) => block.list?.suppressed)).toEqual([true, true]);
  }
});

test('legacy recap prefixes record removed UTF-16 spans while retaining authored root AST and marks', () => {
  const result = logical('<recap_goals>\n\n**1\\. First**\n\n2\\. Second\n\n</recap_goals>');
  const children = group(result.blocks[0]!).children;
  expect(children.map(text)).toEqual(['First', 'Second']);
  expect(children.map((block) => paragraph(block).prefixRemoved)).toEqual([3, 3]);
  expect(JSON.stringify(result.ast)).toContain('1. First');
  const compiled = compileInlineTarget(paragraph(children[0]!).ast);
  expect(compiled.success).toBe(true);
  if (compiled.success) expect(compiled.target.tokens[0]!.marks).toEqual(['bold']);
  expect(text(logical('1\\. Outside').blocks[0]!)).toBe('1. Outside');
});

test('nested multiparagraph lists retain paths, item versus continuation roles, starts, and nesting', () => {
  const entry = reference.cases.find((candidate) => candidate.name === 'list continuation')!,
    blocks = logical(entry.markdown).blocks;
  expect(blocks.map(text)).toEqual(['First', 'Continuation.', 'Child', 'Second']);
  expect(blocks.map((block) => block.path)).toEqual([
    [0, 0, 0],
    [0, 0, 1],
    [0, 0, 2, 0, 0],
    [0, 1, 0],
  ]);
  expect(blocks.map((block) => block.list)).toEqual([
    {
      marker: 'ordered',
      checked: null,
      level: 0,
      group: [0],
      item: [0, 0],
      role: 'item',
      start: 3,
      suppressed: false,
    },
    {
      marker: 'ordered',
      checked: null,
      level: 0,
      group: [0],
      item: [0, 0],
      role: 'continuation',
      start: 3,
      suppressed: false,
    },
    {
      marker: 'bullet',
      checked: null,
      level: 1,
      group: [0, 0, 2],
      item: [0, 0, 2, 0],
      role: 'item',
      start: null,
      suppressed: false,
    },
    {
      marker: 'ordered',
      checked: null,
      level: 0,
      group: [0],
      item: [0, 1],
      role: 'item',
      start: 3,
      suppressed: false,
    },
  ]);
});

test('empty task items retain markers without inventing NBSP and authored NBSP stays exact', () => {
  const blocks = tree([
    {
      type: 'list',
      ordered: false,
      children: [
        { type: 'listItem', checked: false, children: [] },
        { type: 'listItem', checked: true, children: [] },
      ],
    },
  ]).blocks;
  expect(blocks.map(text)).toEqual(['', '']);
  expect(blocks.map((block) => block.list?.checked)).toEqual([false, true]);
  expect(paragraph(blocks[0]!).synthetic).toBe(true);
  expect(paragraph(blocks[0]!).inlinePaths).toEqual([]);
  const authored = logical('- [ ]   ').blocks[0]!;
  expect(text(authored)).toBe(' ');
  expect(authored.list?.checked).toBe(false);
  // Bare GFM markers without content are literal text. Only source metadata proves a generated NBSP.
  expect(logical('- [ ]\n- [x]').blocks.map(text)).toEqual(['[ ]', '[x]']);
  const nestedOnly = tree([
    {
      type: 'list',
      ordered: false,
      children: [
        {
          type: 'listItem',
          checked: false,
          children: [
            {
              type: 'list',
              ordered: true,
              start: 0,
              children: [{ type: 'listItem', children: [p('Child')] }],
            },
          ],
        },
      ],
    },
  ]).blocks;
  expect(nestedOnly.map(text)).toEqual(['', 'Child']);
  expect(nestedOnly.map((block) => block.list)).toMatchObject([
    { role: 'item', level: 0, checked: false },
    { role: 'item', level: 1, start: 0 },
  ]);
});

test('list content keys do not imply equal membership, numbering, task state, or writable native identities', () => {
  const a = logical('3. Keep').blocks[0]!,
    b = logical('8. Keep').blocks[0]!,
    c = logical('- [x] Keep').blocks[0]!;
  expect(a.key).toBe(b.key);
  expect(a.key).toBe(c.key);
  expect(a.list).not.toEqual(b.list);
  expect(a.list).not.toEqual(c.list);
  for (const block of [a, b, c]) {
    expect(Object.hasOwn(block, 'id')).toBe(false);
    expect(Object.hasOwn(block.list!, 'ordinal')).toBe(false);
    expect(Object.hasOwn(block.list!, 'instanceId')).toBe(false);
  }
});

test('GFM and bare HTML tables expose comparable row and cell streams with exact original ASTs', () => {
  const gfm = logical('| A | B |\n|---|---|\n| One | Two |'),
    html = logical('<table><tr><td>A</td><td>B</td></tr><tr><td>One</td><td>Two</td></tr></table>');
  expect(gfm.blocks.map(shape)).toEqual([
    tableShape([
      ['A', 'B'],
      ['One', 'Two'],
    ]),
  ]);
  expect(html.blocks.map(shape)).toEqual(gfm.blocks.map(shape));
  expect(html.blocks[0]!.key).toBe(gfm.blocks[0]!.key);
  expect(html.ast).not.toEqual(gfm.ast);
  const cell = group(group(group(html.blocks[0]!).children[0]!).children[1]!);
  expect(cell.path).toEqual([0, 0, 0, 1]);
  expect(paragraph(cell.children[0]!).inlinePaths).toEqual([[0, 0, 0, 1, 0]]);
});

test('native HTML row groups, attributes, header cells, merges, and rich nested cell content retain structure', () => {
  const markdown =
    '<table data-table="keep"><thead data-group="keep"><tr><th colspan="2">Head</th></tr></thead><tbody><tr data-row="keep"><td rowspan="2"><p data-p="keep">First<br />line</p><p>Second</p><table><tr><td>Nested</td></tr></table></td></tr></tbody></table>';
  const result = logical(markdown),
    table = group(result.blocks[0]!);
  expect(table.rowGroups.map((entry) => [entry.header.name, entry.start, entry.end])).toEqual([
    ['thead', 0, 1],
    ['tbody', 1, 2],
  ]);
  expect(table.rowGroups[0]!.header.attributes).toMatchObject([
    { name: 'data-group', value: 'keep' },
  ]);
  const head = group(group(table.children[0]!).children[0]!);
  expect(head.header).toMatchObject({ name: 'th', attributes: [{ name: 'colspan', value: '2' }] });
  const body = group(group(table.children[1]!).children[0]!);
  expect(body.header).toMatchObject({ name: 'td', attributes: [{ name: 'rowspan', value: '2' }] });
  expect(body.children.map(shape)).toEqual([
    para('First\nline'),
    para('Second'),
    tableShape([['Nested']]),
  ]);
  expect(paragraph(body.children[0]!).html).toMatchObject({
    name: 'p',
    attributes: [{ name: 'data-p', value: 'keep' }],
  });
  expect(result.ast).toEqual(parsed(markdown).ast);
});

test('HTML paragraph wrapper attributes stay explicit and affect hints without becoming inline text', () => {
  const plain = logical('Keep').blocks[0]!,
    wrapped = logical('<p>Keep</p>').blocks[0]!,
    marked = logical('<p data-extra="keep">Keep</p>').blocks[0]!;
  expect(text(wrapped)).toBe('Keep');
  expect(wrapped.key).toBe(plain.key);
  expect(marked.key).not.toBe(plain.key);
  expect(paragraph(marked).html).toMatchObject({
    attributes: [{ name: 'data-extra', value: 'keep' }],
  });
  expect(logical('<p></p>').blocks.map(text)).toEqual(['']);
});

test('dynamic or unknown HTML table structure stays a complete inert atom without partial row loss', () => {
  for (const markdown of [
    '<table>{danger()}</table>',
    '<table><caption>Keep caption</caption><tr><td>Cell</td></tr></table>',
    '<table><tbody><CustomRow /></tbody></table>',
  ]) {
    const result = logical(markdown),
      block = result.blocks[0]!;
    expect(block.kind).toBe('atomic');
    const original = parsed(markdown).ast.children[0]!;
    if ('ast' in block)
      expect(block.ast).toEqual(
        original.type === 'paragraph' ? (original.children as TargetAstNode[])[0] : original
      );
    expect(JSON.stringify(result.ast)).toContain('table');
  }
});

test('unknown React blocks, spreads, expressions, fallbacks, and opaque metadata retain their entire AST', () => {
  const ast: TargetAstNode = {
    type: 'mdxJsxFlowElement',
    name: 'InteractiveWidget',
    attributes: [
      { type: 'mdxJsxExpressionAttribute', value: '...danger()' },
      {
        type: 'mdxJsxAttribute',
        name: 'value',
        value: { type: 'mdxJsxAttributeValueExpression', value: 'danger()' },
      },
    ],
    children: [
      { type: 'mdxFlowExpression', value: 'globalThis.__logicalExecuted = true' },
      p('Fallback'),
    ],
    data: { future: { type: 'payload', position: { keep: true } } },
  };
  const result = tree([ast]);
  expect(result.blocks[0]).toMatchObject({ kind: 'atomic', ast });
  expect('children' in result.blocks[0]!).toBe(false);
  expect(result.ast.children[0]).toEqual(ast);
  expect((globalThis as Record<string, unknown>).__logicalExecuted).toBeUndefined();
  const inline = logical('Before <Widget value={danger()} /> after').blocks[0]!;
  expect(text(inline)).toBeNull();
  expect(paragraph(inline).ast.children).toMatchObject([
    { type: 'text', value: 'Before ' },
    { type: 'mdxJsxTextElement', name: 'Widget' },
    { type: 'text', value: ' after' },
  ]);
});

test('host registry kinds govern custom and overridden tags without running host attribute callbacks', () => {
  let calls = 0;
  const definitions = new Map([
    ...DOMAIN_BLOCK_TAGS,
    [
      'custom',
      {
        tag: 'Custom',
        kind: 'container' as const,
        attributes: () => {
          calls++;
          return [];
        },
      },
    ],
    ['lesson_activity', { tag: 'activity', kind: 'atomic' as const, attributes: () => [] }],
  ]);
  const result = normalizeLogicalBlocks(
    parsed('<Custom>Text</Custom>\n\n<activity>\n\nFallback\n\n</activity>'),
    definitions
  );
  expect(result.success).toBe(true);
  if (result.success) {
    expect(result.document.blocks.map(shape)).toEqual([
      domain('custom', [para('Text')]),
      ['atomic', 'activity'],
    ]);
    expect(result.document.blocks[1]).toMatchObject({ ast: { children: [{ type: 'paragraph' }] } });
  }
  expect(calls).toBe(0);
  const ambiguous = normalizeLogicalBlocks(
    parsed('<Custom>Text</Custom>'),
    new Map([
      ['a', { tag: 'Custom', kind: 'atomic' as const }],
      ['b', { tag: 'Custom', kind: 'container' as const }],
    ])
  );
  expect(ambiguous).toMatchObject({ success: false, code: 'ambiguous-target' });
  const sameKind = normalizeLogicalBlocks(
    parsed('<Custom>Text</Custom>'),
    new Map([
      ['a', { tag: 'Custom', kind: 'container' as const }],
      ['b', { tag: 'Custom', kind: 'container' as const }],
    ])
  );
  expect(sameKind.success).toBe(true);
  if (sameKind.success) expect(group(sameKind.document.blocks[0]!).domainType).toBeNull();
});

test('list wrappers and row groups retain unknown data in the original AST without treating opaque positions as locations', () => {
  const metadata = { type: 'future', position: { keep: true } };
  const input = raw([
    {
      type: 'list',
      ordered: true,
      start: 3,
      data: metadata,
      children: [
        {
          type: 'listItem',
          checked: false,
          data: { keep: 'item' },
          children: [{ ...p('Keep'), data: { keep: 'paragraph' } }],
        },
      ],
    },
  ]);
  const before = JSON.stringify(input),
    result = normalizeLogicalBlocks(input);
  expect(result.success).toBe(true);
  if (result.success) {
    expect(result.document.ast).toEqual(input.ast);
    expect(paragraph(result.document.blocks[0]!).ast.data).toEqual({ keep: 'paragraph' });
  }
  expect(JSON.stringify(input)).toBe(before);
  const first = tree([{ ...p('Keep'), data: metadata, position: { start: { offset: 1 } } }])
      .blocks[0]!,
    moved = tree([{ ...p('Keep'), data: metadata, position: { start: { offset: 99 } } }])
      .blocks[0]!,
    changed = tree([{ ...p('Keep'), data: { type: 'future', position: { keep: false } } }])
      .blocks[0]!;
  expect(first.key).toBe(moved.key);
  expect(first.key).not.toBe(changed.key);
  const proto = JSON.parse(
    '{"type":"mdxJsxFlowElement","name":"X","children":[],"data":{"__proto__":{"keep":true}}}'
  ) as TargetAstNode;
  expect(tree([proto]).ast.children[0]).toEqual(proto);
  expect(Object.prototype.hasOwnProperty.call(Object.prototype, 'keep')).toBe(false);
});

test('detached results freeze every nested field without freezing or changing caller objects', () => {
  const input = raw([
      {
        type: 'list',
        ordered: false,
        children: [{ type: 'listItem', checked: false, children: [p('Keep')] }],
      },
    ]),
    before = JSON.stringify(input),
    result = normalizeLogicalBlocks(input);
  expect(result.success).toBe(true);
  const frozen = (value: unknown): boolean =>
    value === null ||
    typeof value !== 'object' ||
    (Object.isFrozen(value) && Object.values(value).every(frozen));
  if (result.success) {
    expect(frozen(result.document)).toBe(true);
    expect(result.document.ast).not.toBe(input.ast);
    expect(Object.isFrozen(input.ast)).toBe(false);
    expect(JSON.stringify(input)).toBe(before);
    (input.ast.children as TargetAstNode[])[0] = p('Changed');
    expect(result.document.blocks.map(text)).toEqual(['Keep']);
  }
});

test('malformed roots, list structures, and orphan structural nodes refuse without partial output', () => {
  const invalid = [
    null,
    {},
    { ast: { type: 'root', children: [] }, markdown: 5 },
    raw([{ type: 'list', ordered: 'yes', children: [] } as unknown as TargetAstNode]),
    raw([{ type: 'list', ordered: true, start: -1, children: [] }]),
    raw([{ type: 'list', ordered: true, start: 1.5, children: [] }]),
    raw([{ type: 'list', ordered: false, children: [p('Wrong item')] }]),
    raw([
      {
        type: 'list',
        ordered: false,
        children: [{ type: 'listItem', checked: 'yes', children: [] }],
      },
    ]),
    raw([{ type: 'paragraph' }]),
    raw([{ type: 'table', children: [p('Wrong row')] }]),
    raw([{ type: 'table', children: [{ type: 'tableRow', children: [p('Wrong cell')] }] }]),
    ...['listItem', 'tableRow', 'tableCell'].map((type) => raw([{ type, children: [] }])),
  ];
  for (const input of invalid) {
    const result = normalizeLogicalBlocks(input as Parameters<typeof normalizeLogicalBlocks>[0]);
    expect(result.success).toBe(false);
    expect(Object.hasOwn(result, 'document')).toBe(false);
  }
});

test('getters, toJSON, hidden fields, cycles, sparse arrays, and byte/node/depth excess refuse before evaluation', () => {
  let calls = 0;
  const getter = Object.defineProperty(raw([]), 'ast', {
      enumerable: true,
      get() {
        calls++;
        return { type: 'root', children: [] };
      },
    }),
    nestedGetter = Object.defineProperty(p('Keep'), 'children', {
      enumerable: true,
      get() {
        calls++;
        return [];
      },
    }),
    toJSON = {
      ...p('Keep'),
      toJSON() {
        calls++;
        return p('Changed');
      },
    },
    hidden = Object.defineProperty(p('Keep'), 'hidden', { value: true });
  const cycle: Record<string, unknown> = { type: 'paragraph' };
  cycle.children = [cycle];
  let deep: TargetAstNode = p('Leaf');
  for (let i = 0; i < 34; i++) deep = { type: 'blockquote', children: [deep] };
  const values = [
    getter,
    raw([nestedGetter]),
    raw([toJSON as unknown as TargetAstNode]),
    raw([hidden]),
    raw([{ ...p('Keep'), [Symbol('hidden')]: true }]),
    raw([cycle as TargetAstNode]),
    raw(new Array(2)),
    raw([deep]),
    raw([t('x'.repeat(LOGICAL_BLOCK_LIMITS.maxBytes))]),
    raw([t('🧑'.repeat(LOGICAL_BLOCK_LIMITS.maxBytes / 4))]),
    raw(new Array(LOGICAL_BLOCK_LIMITS.maxNodes + 1).fill(p(''))),
  ];
  for (const input of values) {
    const result = normalizeLogicalBlocks(input);
    expect(result.success).toBe(false);
    expect(Object.hasOwn(result, 'document')).toBe(false);
  }
  expect(calls).toBe(0);
});

test('normalization and refusals preserve native bytes, revision, selection, imported XML, and undo/redo history', () => {
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
    for (const entry of pairs.cases) logical(entry.afterMarkdown);
    expect(normalizeLogicalBlocks(raw([{ type: 'listItem', children: [] }])).success).toBe(false);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(canonicalOoxmlFingerprint(session.part())).toBe(xml);
    expect(session.packageRevision()).toBe(revision);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(session.canRedo()).toBe(true);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
  } finally {
    editor.destroy();
  }
});
