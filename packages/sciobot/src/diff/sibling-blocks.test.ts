import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  findNode,
  insertChildren,
  paragraphTextOf,
  readOoxmlPart,
  removeNode,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import reference from '../__tests__/fixtures/ai-pairs.json';
import { bodyPart, W } from '../../test/support/document';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { decodeEditPlan, serializeEditPlan } from './plan';
import {
  planSiblingBlocks,
  SIBLING_BLOCK_LIMITS,
  type SiblingBlockSource,
  type SiblingBlockResult,
} from './sibling-blocks';

const revision = { documentId: 'sibling-test', epoch: 'mount-1', revision: 5 };
const text = (value: string): TargetAstNode => ({
  type: 'paragraph',
  children: [{ type: 'text', value }],
});
function parsed(markdown: string) {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  return result.target.normalizedAst.children;
}
function children(node: OoxmlNode): readonly OoxmlNode[] {
  return node.kind === 'textValue' ? [] : node.children;
}
function source(
  ast: readonly TargetAstNode[],
  options: { part?: OoxmlPart; parentId?: string; startIndex?: number } = {}
): SiblingBlockSource {
  const part =
    options.part ??
    bodyPart(
      ast
        .map(
          (_, index) =>
            `<w:p><w:pPr><w:customKeep w:val="${index}"/></w:pPr><w:r><w:t>P${index}</w:t></w:r></w:p>`
        )
        .join('')
    );
  const parentId = options.parentId ?? part.root.children[0]!.id;
  const startIndex = options.startIndex ?? 0;
  return {
    revision,
    part,
    parentId,
    startIndex,
    blocks: ast.map((node, index) => ({
      id: children(findNode(part, parentId)!)[startIndex + index]!.id,
      ast: node,
    })),
  };
}
function success(result: SiblingBlockResult) {
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result;
}
const xml = (value: string) =>
  value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');

test('gap ranges identify equal inserted fragments in reversed operation order', () => {
  const before = [text('A'), text('B'), text('C')],
    after = [text('Same'), before[0]!, text('Same'), before[1]!, before[2]!],
    result = success(planSiblingBlocks(source(before), after));
  expect(result.gaps).toEqual([
    { operationIndex: 0, beforeStart: 1, beforeEnd: 1, afterStart: 2, afterEnd: 3 },
    { operationIndex: 1, beforeStart: 0, beforeEnd: 0, afterStart: 0, afterEnd: 1 },
  ]);
  for (const gap of result.gaps) {
    const edit = result.plan.operations[gap.operationIndex]!;
    if (edit.kind !== 'siblings') throw new Error('Expected sibling operation');
    expect(edit.inserted).toEqual(after.slice(gap.afterStart, gap.afterEnd));
    expect(edit.inserted[0]).not.toBe(after[gap.afterStart]);
  }
});
function inline(node: TargetAstNode): string {
  if (node.type === 'text') return String(node.value);
  if (Array.isArray(node.children)) return (node.children as TargetAstNode[]).map(inline).join('');
  throw new Error(`Oracle only authors text fragments: ${node.type}`);
}

/** Pure native oracle for these paragraph/heading fixtures. Production execution belongs to A01/R05. */
function applyStructure(input: SiblingBlockSource, result: ReturnType<typeof success>): OoxmlPart {
  let part = input.part,
    fresh = 0;
  for (const edit of result.plan.operations) {
    if (edit.kind !== 'siblings') throw new Error('Expected sibling operation');
    for (const id of edit.removedIds) {
      const removed = removeNode(part, id);
      if (!removed.ok) throw new Error('Native removal failed');
      part = removed.part;
    }
    const parent = findNode(part, edit.parentId)!;
    const current = children(parent);
    const at =
      edit.afterId !== null
        ? current.findIndex((node) => node.id === edit.afterId)
        : edit.beforeId !== null
          ? current.findIndex((node) => node.id === edit.beforeId) + 1
          : 0;
    const nodes = edit.inserted.map((ast) => {
      const loaded = readOoxmlPart(
        `<w:document xmlns:w="${W}"><w:body><w:p>${ast.type === 'heading' ? `<w:pPr><w:pStyle w:val="Heading${ast.depth}"/></w:pPr>` : ''}<w:r><w:t>${xml(inline(ast))}</w:t></w:r></w:p></w:body></w:document>`,
        { name: `/word/new-${fresh++}.xml`, contentType: 'application/xml' }
      );
      if (!loaded.ok) throw new Error('Oracle native fragment failed');
      return children(loaded.part.root.children[0]!)[0]!;
    });
    if (nodes.length) {
      const inserted = insertChildren(part, edit.parentId, at, nodes);
      if (!inserted.ok) throw new Error('Native insertion failed');
      part = inserted.part;
    }
  }
  return part;
}

for (const name of ['section insertion', 'deletion'])
  test(`captured C01 ${name} retains surrounding native IDs, unknown properties, and parent identity`, () => {
    const fixture = reference.cases.find((item) => item.name === name)!;
    const old = parsed(fixture.beforeMarkdown),
      target = parsed(fixture.afterMarkdown);
    const part = bodyPart(
      old
        .map(
          (node, index) =>
            `<w:p><w:pPr><w:customKeep w:val="${index}"/></w:pPr><w:r><w:t>${xml(inline(node))}</w:t></w:r></w:p>`
        )
        .join('')
    );
    const input = source(old, { part });
    const before = canonicalOoxmlFingerprint(part);
    const result = success(planSiblingBlocks(input, target));
    expect(result.plan.operations).toHaveLength(1);
    expect(result.pairs.every((pair) => !pair.changed)).toBe(true);
    const op = result.plan.operations[0]!;
    if (op.kind !== 'siblings') throw new Error('Expected sibling operation');
    expect(op.parentId).toBe(input.parentId);
    if (name === 'section insertion') {
      expect(op).toMatchObject({
        beforeId: input.blocks[0]!.id,
        afterId: input.blocks[1]!.id,
        removedIds: [],
        inserted: target.slice(1, 4),
      });
      expect(result.pairs.map((pair) => [pair.beforeIndex, pair.afterIndex])).toEqual([
        [0, 0],
        [1, 4],
      ]);
    } else {
      expect(op).toMatchObject({
        beforeId: input.blocks[0]!.id,
        afterId: input.blocks[2]!.id,
        removedIds: [input.blocks[1]!.id],
        inserted: [],
      });
    }
    const applied = applyStructure(input, result);
    expect(
      children(findNode(applied, input.parentId)!).map((node) => paragraphTextOf(applied, node.id))
    ).toEqual(target.map(inline));
    expect(findNode(applied, input.parentId)?.id).toBe(input.parentId);
    for (const pair of result.pairs) {
      expect(findNode(applied, pair.id) === findNode(part, pair.id)).toBe(true);
      expect(canonicalOoxmlFingerprint(findNode(applied, pair.id)!)).toBe(
        canonicalOoxmlFingerprint(findNode(part, pair.id)!)
      );
    }
    expect(canonicalOoxmlFingerprint(part)).toBe(before);
    expect(decodeEditPlan(serializeEditPlan(result.plan))).toEqual({
      success: true,
      plan: result.plan,
    });
    expect(result.plan.operations.some((edit) => edit.kind !== 'siblings')).toBe(false);
  });

test('changed paragraphs and headings remain pairs for content planners, without structural replacement', () => {
  const before = parsed('Before cat after.\n\n## Heading'),
    after = parsed('Before car after.\n\n### Heading');
  const input = source(before);
  const result = success(planSiblingBlocks(input, after));
  expect(result.plan.operations).toEqual([]);
  expect(result.pairs).toEqual(
    input.blocks.map((block, index) => ({
      id: block.id,
      beforeIndex: index,
      afterIndex: index,
      changed: true,
    }))
  );
});

test('stable keys ignore AST property order and parser positions, while preserving opaque metadata', () => {
  const before = parsed('Same paragraph');
  const input = source(before);
  const positioned = {
    ...before[0]!,
    position: { start: { line: 1, column: 1, offset: 0 }, end: { line: 2, column: 1, offset: 15 } },
  };
  const result = success(planSiblingBlocks(input, [positioned]));
  expect(result.plan.operations).toEqual([]);
  expect(result.pairs[0]!.changed).toBe(false);
  const opaque: TargetAstNode = {
    type: 'mdxJsxFlowElement',
    name: 'Unknown',
    attributes: [],
    children: [],
    data: { position: 'retained', value: 1 },
  };
  const next = { ...opaque, data: { value: 1, position: 'retained' } };
  expect(success(planSiblingBlocks(source([opaque]), [next])).plan.operations).toEqual([]);
  expect(
    success(
      planSiblingBlocks(source([opaque]), [{ ...opaque, data: { value: 2, position: 'retained' } }])
    ).plan.operations[0]
  ).toMatchObject({ inserted: [{ data: { value: 2, position: 'retained' } }] });
  const typed = { ...opaque, data: { type: 'opaque-value', position: 'first' } };
  expect(
    success(
      planSiblingBlocks(source([typed]), [
        { ...typed, data: { type: 'opaque-value', position: 'second' } },
      ])
    ).plan.operations[0]
  ).toMatchObject({ inserted: [{ data: { type: 'opaque-value', position: 'second' } }] });
});

test('nested parent scopes retain their native container and outer siblings', () => {
  const part = bodyPart(
    '<w:sdt><w:sdtPr><w:tag w:val="container"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Opening.</w:t></w:r></w:p><w:p><w:r><w:t>Closing.</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>Outer tail.</w:t></w:r></w:p>'
  );
  const root = children(part.root.children[0]!);
  const parent = children(root[0]!).find(
    (node) => node.kind !== 'textValue' && node.localName === 'sdtContent'
  )!;
  const input = source(parsed('Opening.\n\nClosing.'), { part, parentId: parent.id });
  const result = success(planSiblingBlocks(input, parsed('Opening.\n\n## Added\n\nClosing.')));
  expect(result.plan.operations[0]).toMatchObject({ parentId: parent.id, removedIds: [] });
  const applied = applyStructure(input, result);
  expect(findNode(applied, root[0]!.id)?.id).toBe(root[0]!.id);
  expect(findNode(applied, root[1]!.id) === root[1]).toBe(true);
  expect(
    children(findNode(applied, parent.id)!).map((node) => paragraphTextOf(applied, node.id))
  ).toEqual(['Opening.', 'Added', 'Closing.']);
});

test('partial windows anchor against surrounding unprojected nodes and preserve terminal section properties', () => {
  const part = bodyPart(
    '<w:customKeep w:val="opaque"/><w:p><w:r><w:t>Old</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
  );
  const parent = part.root.children[0]!,
    siblings = children(parent);
  const input = source([text('Old')], { part, startIndex: 1 });
  const result = success(planSiblingBlocks(input, [text('Entire replacement')]));
  expect(result.plan.operations[0]).toMatchObject({
    beforeId: siblings[0]!.id,
    afterId: siblings[2]!.id,
    removedIds: [siblings[1]!.id],
  });
  const applied = applyStructure(input, result);
  for (const keep of [siblings[0]!, siblings[2]!])
    expect(findNode(applied, keep.id) === keep).toBe(true);
});

test('edge insertions, empty windows, complete removals, and multiple gaps retain canonical anchors', () => {
  const input = source([text('Keep A'), text('Keep B')]);
  const result = success(
    planSiblingBlocks(input, [text('New head'), text('Keep A'), text('Keep B'), text('New tail')])
  );
  expect(result.plan.operations).toEqual([
    {
      kind: 'siblings',
      partName: input.part.name,
      parentId: input.parentId,
      beforeId: input.blocks[1]!.id,
      afterId: null,
      removedIds: [],
      inserted: [text('New tail')],
    },
    {
      kind: 'siblings',
      partName: input.part.name,
      parentId: input.parentId,
      beforeId: null,
      afterId: input.blocks[0]!.id,
      removedIds: [],
      inserted: [text('New head')],
    },
  ]);
  const empty = { ...input, startIndex: 1, blocks: [] };
  expect(success(planSiblingBlocks(empty, [text('Inserted')])).plan.operations[0]).toMatchObject({
    beforeId: input.blocks[0]!.id,
    afterId: input.blocks[1]!.id,
    removedIds: [],
  });
  expect(success(planSiblingBlocks(input, [])).plan.operations[0]).toMatchObject({
    beforeId: null,
    afterId: null,
    removedIds: input.blocks.map((block) => block.id),
    inserted: [],
  });
  expect(success(planSiblingBlocks(empty, [])).plan.operations).toEqual([]);
});

test('duplicate-looking siblings retain separate IDs and deterministic pair order', () => {
  const input = source([text('Same'), text('Same'), text('Tail')]);
  const result = success(
    planSiblingBlocks(input, [text('Same'), text('Same'), text('Same'), text('Tail')])
  );
  expect(result.pairs.map((pair) => pair.id)).toEqual(input.blocks.map((block) => block.id));
  expect(new Set(result.pairs.map((pair) => pair.afterIndex)).size).toBe(3);
  expect(
    success(planSiblingBlocks(input, [text('Same'), text('Same'), text('Same'), text('Tail')]))
  ).toEqual(result);
});

test('changed native container pairs delegate nested content without replacing their parent', () => {
  for (const type of ['blockquote', 'list', 'listItem', 'table', 'tableRow', 'tableCell']) {
    const before: TargetAstNode = { type, children: [text('Old content')] };
    const result = success(
      planSiblingBlocks(source([before]), [{ type, children: [text('New content')] }])
    );
    expect(result.plan.operations).toEqual([]);
    expect(result.pairs[0]!.changed).toBe(true);
  }
});

test('plan fragments and pair records detach from caller-owned data', () => {
  const input = source([text('Keep')]);
  const target = [text('Keep'), text('Added')];
  const result = success(planSiblingBlocks(input, target));
  const fingerprint = canonicalOoxmlFingerprint(input.part);
  (target[1]!.children as { value: string }[])[0]!.value = 'Changed later';
  expect(result.plan.operations[0]).toMatchObject({ inserted: [text('Added')] });
  expect(Object.isFrozen(target)).toBe(false);
  expect(Object.isFrozen(result.plan.operations)).toBe(true);
  expect(Object.isFrozen(result.pairs)).toBe(true);
  expect(Object.isFrozen(result.pairs[0])).toBe(true);
  expect(canonicalOoxmlFingerprint(input.part)).toBe(fingerprint);
});

test('invalid identities, noncontiguous mappings, native trees, getters, and resource excess refuse without writes', () => {
  const input = source([text('One'), text('Two'), text('Three')]);
  const before = canonicalOoxmlFingerprint(input.part);
  for (const altered of [
    { ...input, startIndex: -1 },
    { ...input, startIndex: 4 },
    { ...input, parentId: 'missing' },
    { ...input, blocks: [input.blocks[0]!, input.blocks[2]!] },
    { ...input, blocks: [input.blocks[0]!, input.blocks[0]!] },
  ])
    expect(planSiblingBlocks(altered, []).success).toBe(false);
  for (const target of [
    [{ type: 'root' }],
    [{ type: 'paragraph', kind: 'paragraph' }],
    [{ type: 'paragraph', children: [{ type: 'text', value: 'A', kind: 'textValue' }] }],
  ])
    expect(planSiblingBlocks(input, target).success).toBe(false);
  const nativeAsAst = findNode(input.part, input.blocks[0]!.id)! as unknown as TargetAstNode;
  expect(planSiblingBlocks(input, [nativeAsAst]).success).toBe(false);
  let read = false;
  const node = Object.defineProperty({ type: 'paragraph' }, 'children', {
    enumerable: true,
    get() {
      read = true;
      return [];
    },
  }) as TargetAstNode;
  expect(planSiblingBlocks(input, [node]).success).toBe(false);
  expect(read).toBe(false);
  const sourceGetter = Object.defineProperty({ ...input }, 'blocks', {
    enumerable: true,
    get() {
      read = true;
      return [];
    },
  });
  expect(planSiblingBlocks(sourceGetter, []).success).toBe(false);
  expect(read).toBe(false);
  expect(
    planSiblingBlocks(
      input,
      Array.from({ length: SIBLING_BLOCK_LIMITS.maxSiblings + 1 }, () => text('New'))
    ).success
  ).toBe(false);
  expect(
    planSiblingBlocks(input, [text('x'.repeat(SIBLING_BLOCK_LIMITS.maxBytes + 1))]).success
  ).toBe(false);
  expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
});

test('unmapped canonical nodes cannot disappear inside a flattened sibling window', () => {
  const part = bodyPart('<w:p/><w:customKeep/><w:p/>');
  const parent = part.root.children[0]!;
  const blocks = children(parent);
  const input = {
    ...source([], { part }),
    blocks: [
      { id: blocks[0]!.id, ast: text('One') },
      { id: blocks[2]!.id, ast: text('Two') },
    ],
  };
  expect(planSiblingBlocks(input, [])).toMatchObject({ success: false, code: 'ambiguous-target' });
  const includesOpaque = {
    ...input,
    blocks: blocks.map((node, index) => ({ id: node.id, ast: text(String(index)) })),
  };
  expect(planSiblingBlocks(includesOpaque, [])).toMatchObject({
    success: false,
    code: 'unsupported-operation',
  });
});

test('generated structural plans and retained content pairs reconstruct target order exactly', () => {
  let seed = 29;
  const random = () => {
    seed = (seed * 1664525 + 1013904223) >>> 0;
    return seed / 2 ** 32;
  };
  for (let round = 0; round < 100; round++) {
    const old = Array.from({ length: 2 + Math.floor(random() * 9) }, () =>
      text(`Block ${Math.floor(random() * 5)}`)
    );
    const next = old.map((node) => JSON.parse(JSON.stringify(node)) as TargetAstNode);
    for (let step = 0; step < 3; step++) {
      const index = Math.floor(random() * (next.length + 1));
      next.splice(index, random() < 0.5 ? 1 : 0, text(`New ${round}-${step}`));
    }
    const input = source(old),
      result = success(planSiblingBlocks(input, next));
    let current = input.blocks.map((block) => ({ id: block.id, ast: block.ast }));
    for (const pair of result.pairs) current[pair.beforeIndex]!.ast = next[pair.afterIndex]!;
    for (const edit of result.plan.operations) {
      if (edit.kind !== 'siblings') throw new Error('Expected structural gap');
      current = current.filter((block) => !edit.removedIds.includes(block.id));
      const index =
        edit.afterId === null
          ? current.length
          : current.findIndex((block) => block.id === edit.afterId);
      expect(index).toBeGreaterThanOrEqual(0);
      current.splice(
        index,
        0,
        ...edit.inserted.map((ast, i) => ({ id: `insert-${round}-${i}`, ast }))
      );
    }
    expect(current.map((block) => block.ast)).toEqual(next);
    for (const pair of result.pairs) expect(current[pair.afterIndex]!.id).toBe(pair.id);
  }
});
