import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  findNode,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import {
  formatCell,
  formatParagraph,
  formatRow,
  formatTable,
} from '../../test/support/table-format-fixture';
import {
  targetSourceFixture,
  projectSourcePackage,
  sourceRevision,
  sourceDefinitions,
  sourceControl,
  sourceParagraph,
} from '../../test/support/target-source-fixture';
import { checklistParagraph } from '../../test/support/list-document';
import { protectedDocument } from '../../test/support/protected-document';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { MDX_BLOCK, MDX_INLINE } from '../projection/mdx';
import { parseTargetMarkdown } from '../projection/parse-target';
import { markdownRangeAt } from '../projection/markdown';
import { nearestCellAncestor } from './cell-scope';
import { bindLogicalSource, type BoundLogicalSource } from './bind-source';
import type { LogicalBlock } from './logical-blocks';
import type { TargetSourceInput } from './target-source';
import { planSiblingBlocks } from './sibling-blocks';

const fixture = targetSourceFixture();
function bound(input: TargetSourceInput, definitions = sourceDefinitions): BoundLogicalSource {
  const result = bindLogicalSource(input, definitions);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.source;
}
async function inputFor(body: string): Promise<TargetSourceInput> {
  const result = readOoxmlPackage(docxFixture(body));
  if (!result.ok) throw new Error(result.reason);
  return {
    revision: sourceRevision,
    pkg: result.package,
    projection: await projectSourcePackage(result.package),
  };
}
const part = (input: TargetSourceInput) => input.pkg.parts.get(input.pkg.mainDocumentPart)!;
const owner = (input: TargetSourceInput, source: BoundLogicalSource, text: string) =>
  source.index.owners.find(
    (entry) => entry.kind === 'paragraph' && paragraphTextOf(part(input), entry.id) === text
  )!;
const binding = (source: BoundLogicalSource, id: string) =>
  source.bindings.find((entry) => entry.contributions.some((item) => item.ownerId === id))!;
function blockAt(source: BoundLogicalSource, path: readonly number[]): LogicalBlock {
  let children = source.logical.blocks,
    current: LogicalBlock | undefined;
  for (const index of path) {
    current = children[index];
    if (!current) throw new Error('No logical block.');
    children = 'children' in current ? current.children : [];
  }
  if (!current) throw new Error('Empty block path.');
  return current;
}
const mutable = (input: TargetSourceInput) => ({
  ...input,
  projection: JSON.parse(JSON.stringify(input.projection)),
});

test('composed MDX retains native table metadata and remaps cell sources through domain wrappers', async () => {
  const input = await fixture;
  expect(input.projection.tables).toHaveLength(2);
  expect(input.projection.cells).toHaveLength(10);
  const nested = input.projection.cells.find((cell) =>
    cell.sources.some((slice) =>
      input.projection.text.slice(slice.markdownStart, slice.markdownEnd).includes('Nested B')
    )
  )!;
  expect(nested).toBeDefined();
  expect(nearestCellAncestor(part(input), nested.paragraphIds[0]!, 'tableCell')?.id).toBe(
    nested.id
  );
  expect(
    input.projection.tables.find((table) => table.id === nested.tableId)?.parentCellId
  ).not.toBeNull();
  for (const cell of input.projection.cells)
    expect(cell.sources).toEqual(
      input.projection.sourceMap.sources.filter(
        (slice) => slice.sourceScope === 'body' && cell.paragraphIds.includes(slice.paragraphId)
      )
    );
  expect(Object.isFrozen(input.projection.cells)).toBe(true);
});

test('domain containers and their child blocks bind to exact native controls and actual content parents', async () => {
  const input = await fixture,
    source = bound(input),
    blocks = buildBlockIndex(input.pkg);
  for (const [id, logicalPath] of [
    ['phase', [1]],
    ['activity', [1, 0]],
  ] as const) {
    const native = blocks.byId.get(id)!;
    const result = binding(source, native.controlId);
    expect(result.logicalPath).toEqual(logicalPath);
    expect(result.structuralOwnerId).toBe(native.controlId);
    expect(result.source).toMatchObject({
      start: input.projection.blocks.find((entry) => entry.id === id)!.start,
      end: input.projection.blocks.find((entry) => entry.id === id)!.end,
    });
  }
  const question = owner(input, source, 'Question'),
    result = binding(source, question.id);
  expect(result.logicalPath).toEqual([1, 0, 0]);
  expect(result.structuralOwnerId).toBe(question.id);
  const window = source.windows.find((entry) => entry.ids.includes(question.id))!;
  expect(window.parentId).toBe(question.parentId);
  expect(window.beforeId).not.toBeNull();
  expect(source.unrepresented.filter((entry) => entry.reason === 'generated')).toHaveLength(2);
});

test('merged grid slots keep physical owners and explicit horizontal/vertical placeholder roles', async () => {
  const input = await fixture,
    source = bound(input),
    table = input.projection.tables.find((entry) => entry.parentCellId === null)!;
  const slots = source.bindings.filter(
    (entry) =>
      blockAt(source, entry.logicalPath).kind === 'cell' && entry.table?.tableId === table.id
  );
  expect(slots.map((entry) => entry.table?.role)).toEqual([
    'cell',
    'horizontal-placeholder',
    'cell',
    'vertical-placeholder',
    'horizontal-placeholder',
    'cell',
    'cell',
    'cell',
    'cell',
  ]);
  expect(slots[0]!.table?.cellId).toBe(slots[1]!.table?.cellId);
  expect(slots[3]!.table?.cellId).toBe(slots[4]!.table?.cellId);
  expect(slots[0]!.table?.cellId).not.toBe(slots[3]!.table?.cellId);
  expect(slots[0]!.table?.gridSpan).toBe(2);
  for (const entry of slots.filter((entry) => entry.table?.role !== 'cell')) {
    expect(entry.contributions).toEqual([]);
    expect(entry.structuralOwnerId).toBeNull();
    expect(
      source.bindings.some(
        (child) =>
          child.logicalPath.length > entry.logicalPath.length &&
          child.logicalPath
            .slice(0, entry.logicalPath.length)
            .every((value, index) => value === entry.logicalPath[index])
      )
    ).toBe(false);
  }
});

test('nested HTML tables retain nearest native cells after break normalization and distinguish AST from logical paths', async () => {
  const input = await fixture,
    source = bound(input),
    native = owner(input, source, 'Nested B'),
    result = binding(source, native.id);
  expect(result.table?.cellId).toBe(native.cell?.cellId);
  expect(result.table?.tableId).toBe(native.cell?.tableId);
  expect(source.parsed.markdown).toContain('<br />');
  expect(source.parsed.sourceMarkdown).toContain('<br>');
  expect(result.source).toEqual({
    start: input.projection.text.indexOf('**Nested B**'),
    end: input.projection.text.indexOf('**Nested B**') + 12,
    exact: true,
    generated: false,
  });
  expect(result.astPath).not.toEqual(result.logicalPath);
  expect(result.structuralOwnerId).toBe(native.id);
  expect(
    markdownRangeAt(input.projection, result.source!.start + 2, result.source!.end - 2)
  ).toMatchObject({
    start: { paragraphId: native.id, offset: 0 },
    end: { paragraphId: native.id, offset: 8 },
  });
});

test('repeated table values bind by native source extents and nearest cells, across distinct tables', async () => {
  const input = await inputFor(
      formatTable(formatRow('Same', 'Same', 'Same')) +
        formatTable(formatRow('Same', 'Same', 'Same'))
    ),
    source = bound(input);
  expect(input.projection.tables).toHaveLength(2);
  const ids: string[] = [];
  for (const cell of input.projection.cells) {
    const id = cell.paragraphIds[0]!,
      result = binding(source, id);
    expect(result.structuralOwnerId).toBe(id);
    expect(result.table).toMatchObject({
      tableId: cell.tableId,
      rowId: cell.rowId,
      cellId: cell.id,
      gridColumn: cell.gridColumn,
      role: 'cell',
    });
    expect(paragraphTextOf(part(input), id)).toBe('Same');
    ids.push(result.structuralOwnerId!);
  }
  expect(new Set(ids).size).toBe(6);
});

test('empty physical cells retain canonical windows and empty paragraphs without inventing AST text', async () => {
  const input = await inputFor(formatTable(formatRow('', 'Text', ''))),
    source = bound(input);
  const empty = input.projection.cells.filter((cell) => !cell.sources.length);
  expect(empty).toHaveLength(2);
  for (const cell of empty) {
    const result = source.bindings.find(
      (entry) =>
        entry.table?.cellId === cell.id && blockAt(source, entry.logicalPath).kind === 'cell'
    )!;
    expect(result.table?.role).toBe('cell');
    expect(result.contributions).toEqual([]);
    expect(source.unrepresented).toContainEqual({
      ownerId: cell.paragraphIds[0]!,
      reason: 'empty-extent',
    });
    const window = source.windows.find((entry) => entry.parentId === cell.id)!;
    expect(window.members).toEqual([{ ownerId: cell.paragraphIds[0]!, logicalPaths: [] }]);
  }
});

test('absent physical grid slots keep null identities instead of inventing empty native cells', async () => {
  const input = await inputFor(
      formatTable('<w:tr>' + formatCell('A') + formatCell('B') + '</w:tr>')
    ),
    source = bound(input),
    missing = source.bindings.filter((entry) => entry.table?.role === 'missing-slot');
  expect(missing).toHaveLength(1);
  expect(missing[0]!.table).toMatchObject({ cellId: null, gridColumn: 2, gridSpan: null });
  expect(missing[0]!.contributions).toEqual([]);
  expect(missing[0]!.structuralOwnerId).toBeNull();
});

test('one logical GFM cell paragraph retains several canonical paragraph contributors without one invented owner', async () => {
  const body = formatTable(formatRow('First', 'B', 'C')).replace(
    formatParagraph('First'),
    formatParagraph('First') + formatParagraph('Second')
  );
  const input = await inputFor(body),
    source = bound(input),
    first = owner(input, source, 'First'),
    second = owner(input, source, 'Second'),
    result = binding(source, first.id);
  expect(result.contributions.map((entry) => entry.ownerId)).toEqual([first.id, second.id]);
  expect(binding(source, second.id)).toBe(result);
  expect(result.structuralOwnerId).toBeNull();
  expect(result.table?.cellId).toBe(first.cell?.cellId);
  expect(first.parentId).toBe(second.parentId);
  const window = source.windows.find((entry) => entry.ids.includes(first.id))!;
  expect(window.members.map((entry) => entry.logicalPaths)).toEqual([
    [result.logicalPath],
    [result.logicalPath],
  ]);
});

test('several parsed blocks from one canonical paragraph keep competing claims explicit', async () => {
  const input = await inputFor(
      '<w:p><w:r><w:t>First</w:t><w:br/><w:br/><w:t>Second</w:t></w:r></w:p>'
    ),
    source = bound(input),
    native = source.index.owners.find((entry) => entry.kind === 'paragraph')!;
  const matches = source.bindings.filter((entry) =>
    entry.contributions.some((item) => item.ownerId === native.id)
  );
  expect(matches).toHaveLength(2);
  expect(matches.map((entry) => entry.structuralOwnerId)).toEqual([null, null]);
  expect(source.windows[0]!.members[0]!.logicalPaths).toEqual(
    matches.map((entry) => entry.logicalPath)
  );
});

test('ordinary Word controls preserve physical parents and do not enter a flattened body sibling window', async () => {
  const input = await fixture,
    source = bound(input),
    native = owner(input, source, 'Word form text'),
    result = binding(source, native.id);
  expect(native.parentId).not.toBe(source.index.bodyId);
  const window = source.windows.find((entry) => entry.ids.includes(native.id))!;
  expect(window.parentId).toBe(native.parentId);
  expect(window.members[0]!.logicalPaths).toEqual([result.logicalPath]);
  expect(
    source.windows
      .filter((entry) => entry.parentId === source.index.bodyId)
      .flatMap((entry) => entry.ids)
  ).not.toContain(native.id);
  expect(source.windows.filter((entry) => entry.parentId === source.index.bodyId)).toHaveLength(3);
});

test('custom empty containers and atomic fallbacks retain native policy and unrepresented owner reasons', async () => {
  const input = await fixture,
    source = bound(input),
    blocks = buildBlockIndex(input.pkg),
    container = blocks.byId.get('custom_container')!,
    atom = blocks.byId.get('custom_atomic')!;
  expect(blockAt(source, binding(source, container.controlId).logicalPath).kind).toBe('container');
  const empty = source.index.owners.find(
    (entry) =>
      entry.kind === 'paragraph' &&
      entry.parentId.includes(container.controlId) &&
      entry.extents.length
  )!;
  expect(source.unrepresented).toContainEqual({ ownerId: empty.id, reason: 'empty-extent' });
  expect(binding(source, atom.controlId).atomicControlId).toBe(atom.controlId);
  expect(source.unrepresented.filter((entry) => entry.reason === 'atomic-fallback')).toHaveLength(
    2
  );
});

test('inline opaque controls retain complete AST and actual model width, independently of surrounding text', async () => {
  const input = await fixture,
    source = bound(input),
    native = owner(input, source, 'Inline before Inline fallback after'),
    result = binding(source, native.id),
    atom = buildBlockIndex(input.pkg).byId.get('inline')!;
  expect(result.inline).toEqual([
    {
      controlId: atom.controlId,
      blockId: 'inline',
      start: input.projection.text.indexOf('<Badge'),
      end: input.projection.text.indexOf('<Badge') + '<Badge label="Keep" />'.length,
      native: { start: 14, end: 29 },
    },
  ]);
  expect(result.structuralOwnerId).toBe(native.id);
  expect(source.unrepresented.some((entry) => entry.ownerId === atom.controlId)).toBe(false);
  expect<unknown>(
    source.index.owners.find((entry) => entry.id === atom.controlId)?.opaque?.ast
  ).toEqual(atom.payload.mdast);
});

async function opaqueInput(inline: boolean): Promise<TargetSourceInput> {
  const loaded = readOoxmlPackage(
    docxFixture(
      inline
        ? `<w:p>${sourceControl('opaque', '<w:r><w:t>Fallback</w:t></w:r>')}</w:p>`
        : sourceControl('opaque', sourceParagraph('Fallback'))
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const stored = withPayloadRecord(
    loaded.package,
    {
      version: 1,
      id: 'opaque',
      type: inline ? MDX_INLINE : MDX_BLOCK,
      attributes: {},
      mdast: {
        type: inline ? 'mdxJsxTextElement' : 'mdxJsxFlowElement',
        name: 'activity',
        attributes: [],
        children: inline
          ? []
          : [{ type: 'paragraph', children: [{ type: 'text', value: 'Fallback' }] }],
        data: { future: { type: 'payload', position: { keep: true } } },
      },
    },
    ''
  );
  if (!stored.ok) throw new Error(stored.reason);
  return {
    revision: sourceRevision,
    pkg: stored.pkg,
    projection: await projectSourcePackage(stored.pkg),
  };
}
test('opaque block tags that resemble domain containers keep atomic native ownership and suppress child binding', async () => {
  const input = await opaqueInput(false),
    source = bound(input),
    atom = buildBlockIndex(input.pkg).byId.get('opaque')!,
    result = binding(source, atom.controlId);
  expect(blockAt(source, result.logicalPath).kind).toBe('container');
  expect(result.atomicControlId).toBe(atom.controlId);
  expect(source.bindings).toHaveLength(1);
  expect(source.unrepresented.filter((entry) => entry.reason === 'atomic-fallback')).toHaveLength(
    1
  );
  expect(
    source.index.owners.find((entry) => entry.id === atom.controlId)?.opaque?.ast.data
  ).toEqual({ future: { type: 'payload', position: { keep: true } } });
});

test('standalone inline atoms that resemble containers bind to their real paragraph and retain separate control ownership', async () => {
  const input = await opaqueInput(true),
    source = bound(input),
    atom = buildBlockIndex(input.pkg).byId.get('opaque')!,
    paragraph = source.index.owners.find((entry) => entry.kind === 'paragraph')!,
    result = binding(source, paragraph.id);
  expect(source.bindings).toHaveLength(1);
  expect(result.atomicControlId).toBe(atom.controlId);
  expect(result.structuralOwnerId).toBe(paragraph.id);
  expect(result.inline[0]).toMatchObject({
    controlId: atom.controlId,
    native: { start: 0, end: 8 },
  });
});

test('host atomic overrides suppress fallback tables and preserve the actual source registry kind', async () => {
  const input = await fixture,
    definitions = new Map(sourceDefinitions);
  definitions.set('lesson_activity', { tag: 'activity', kind: 'atomic', attributes: () => [] });
  const changed = {
      ...input,
      projection: await projectSourcePackage(input.pkg, 'proposed', definitions),
    },
    source = bound(changed, definitions),
    activity = buildBlockIndex(input.pkg).byId.get('activity')!;
  expect(binding(source, activity.controlId).atomicControlId).toBe(activity.controlId);
  expect(source.bindings.some((entry) => entry.table?.role === 'table')).toBe(false);
  expect(source.index.owners.find((entry) => entry.id === activity.controlId)?.block?.kind).toBe(
    'atomic'
  );
});

test('a registry mismatch refuses instead of reinterpreting native container children as sibling paragraphs', async () => {
  const input = await fixture;
  expect(bindLogicalSource(input)).toMatchObject({
    success: false,
    code: 'invalid-input',
    error: 'The source registry differs from its native projection.',
  });
  const definitions = new Map(sourceDefinitions);
  definitions.set('lesson_activity', { tag: 'activity', kind: 'atomic', attributes: () => [] });
  expect(bindLogicalSource(input, definitions).success).toBe(false);
});

test('an empty native source excerpt keeps its canonical wrapper even when its parsed form resembles an inline source', async () => {
  const loaded = readOoxmlPackage(docxFixture(sourceControl('excerpt', sourceParagraph(''))));
  if (!loaded.ok) throw new Error(loaded.reason);
  const stored = withPayloadRecord(
    loaded.package,
    {
      version: 1,
      id: 'excerpt',
      type: 'source_excerpt',
      attributes: { url: 'https://example.org', title: 'Book' },
    },
    ''
  );
  if (!stored.ok) throw new Error(stored.reason);
  const input = {
      revision: sourceRevision,
      pkg: stored.pkg,
      projection: await projectSourcePackage(stored.pkg),
    },
    source = bound(input),
    native = buildBlockIndex(input.pkg).byId.get('excerpt')!,
    result = binding(source, native.controlId);
  expect(result.structuralOwnerId).toBe(native.controlId);
  expect(result.atomicControlId).toBeNull();
  expect(blockAt(source, result.logicalPath).kind).toBe('paragraph');
  expect(source.index.owners.find((entry) => entry.id === native.controlId)?.block?.kind).toBe(
    'container'
  );
  expect(source.unrepresented.map((entry) => entry.reason)).toEqual(['empty-extent']);
});

test('generated task NBSP carries exact source proof while authored NBSP retains ordinary text ownership', async () => {
  const input = await fixture,
    source = bound(input),
    task = source.bindings.find((entry) => entry.generatedTask)!;
  const projected = input.projection.checklists[0]!;
  expect(task.generatedTask).toEqual({
    paragraphId: projected.paragraphId,
    controlId: projected.controlId,
    start: projected.contentStart,
    end: projected.contentStart + 1,
    nativeInsertionOffset: 1,
  });
  expect(
    source.parsed.sourceMarkdown.slice(task.generatedTask!.start, task.generatedTask!.end)
  ).toBe(' ');
  expect(blockAt(source, task.logicalPath).list?.checked).toBe(false);
  const authoredInput = await inputFor(
      sourceParagraph(' ') + checklistParagraph(false).replace('Task', ' ')
    ),
    authored = bound(authoredInput);
  expect(authored.bindings.every((entry) => entry.generatedTask === null)).toBe(true);
  expect(authoredInput.projection.checklists[0]?.empty).toBe(false);
});

test('empty task syntax inside native HTML cells keeps nearest cell and original insertion offset', async () => {
  const input = await inputFor(
      formatTable(formatRow('', 'B', 'C')).replace(
        formatParagraph(''),
        checklistParagraph(true).replace('Task', '')
      )
    ),
    source = bound(input);
  expect(input.projection.text).toContain('<table>');
  const task = source.bindings.find((entry) => entry.generatedTask)!;
  expect(task).toBeDefined();
  expect(task.table?.role).toBe('cell');
  expect(task.table?.cellId).toBe(
    source.index.owners.find((entry) => entry.id === task.generatedTask!.paragraphId)?.cell?.cellId
  );
  expect(blockAt(source, task.logicalPath).list?.checked).toBe(true);
});

test('one-to-one returned bindings drive the existing sibling planner with physical anchors', async () => {
  const input = await fixture,
    source = bound(input),
    native = owner(input, source, 'Tail'),
    result = binding(source, native.id),
    window = source.windows.find((entry) => entry.ids.includes(native.id))!,
    block = blockAt(source, result.logicalPath);
  if (block.kind !== 'paragraph') throw new Error('Expected tail paragraph.');
  const parsed = parseTargetMarkdown('New paragraph');
  if (!parsed.success) throw new Error(parsed.error);
  const plan = planSiblingBlocks(
    {
      revision: source.index.revision,
      part: part(input),
      parentId: window.parentId,
      startIndex: window.startIndex,
      blocks: [{ id: result.structuralOwnerId!, ast: block.ast }],
    },
    [...parsed.target.ast.children, block.ast]
  );
  if (!plan.success) throw new Error(plan.error);
  expect(plan.plan.operations).toEqual([
    {
      kind: 'siblings',
      partName: part(input).name,
      parentId: window.parentId,
      beforeId: window.beforeId,
      afterId: native.id,
      removedIds: [],
      inserted: parsed.target.ast.children,
    },
  ]);
  expect(plan.pairs).toMatchObject([{ id: native.id, changed: false }]);
});

test('all returned windows keep actual canonical children and explicit empty or competing path associations', async () => {
  const input = await fixture,
    source = bound(input);
  for (const window of source.windows) {
    const native = findNode(part(input), window.parentId)!;
    if (native.kind === 'textValue') throw new Error('Invalid native window.');
    expect(window.ids).toEqual(
      native.children
        .slice(window.startIndex, window.startIndex + window.ids.length)
        .map((entry) => entry.id)
    );
    expect(window.beforeId).toBe(native.children[window.startIndex - 1]?.id ?? null);
    expect(window.afterId).toBe(native.children[window.startIndex + window.ids.length]?.id ?? null);
    expect(window.members.map((entry) => entry.ownerId)).toEqual([...window.ids]);
  }
});

test('empty metadata-only native bodies retain physical insertion anchors without fabricated logical content', async () => {
  const input = await inputFor('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'),
    source = bound(input);
  expect(source.bindings).toEqual([]);
  expect(source.logical.blocks).toEqual([]);
  expect(source.windows).toHaveLength(1);
  expect(source.windows[0]).toMatchObject({ ids: [], members: [], beforeId: null });
  expect(source.windows[0]!.afterId).not.toBeNull();
});

test('conflicting table, row, cell, grid, and nearest paragraph metadata refuses without a partial binding', async () => {
  const input = await fixture;
  const changes: ((p: ReturnType<typeof mutable>['projection']) => void)[] = [
    (p) => p.tables.push(p.tables[0]),
    (p) => (p.tables[0].parentCellId = 'missing'),
    (p) => (p.tables[0].rowIds[0] = p.cells[0].paragraphIds[0]),
    (p) => p.cells.push(p.cells[0]),
    (p) => (p.cells[0].rowId = p.tables[1].rowIds[0]),
    (p) => p.cells[0].paragraphIds.push(p.cells.at(-1).paragraphIds[0]),
    (p) => (p.cells[0].sources = []),
    (p) => (p.cells[0].gridColumn = p.cells[1].gridColumn),
    (p) => {
      const column = p.cells[0].gridColumn;
      p.cells[0].gridColumn = p.cells[1].gridColumn;
      p.cells[1].gridColumn = column;
    },
    (p) => (p.cells[0].gridSpan = -1),
    (p) => (p.cells[0].gridColumn = 0.5),
    (p) => (p.checklists[0].checked = true),
    (p) => (p.checklists[0].contentStart -= 1),
  ];
  for (const change of changes) {
    const altered = mutable(input);
    change(altered.projection);
    const result = bindLogicalSource(altered, sourceDefinitions);
    expect(result.success).toBe(false);
    expect(Object.hasOwn(result, 'source')).toBe(false);
  }
});

test('stale projections, malformed syntax, getters, cycles, and missing table metadata refuse without evaluation', async () => {
  const input = await fixture;
  let calls = 0;
  const altered = mutable(input);
  Object.defineProperty(altered.projection, 'tables', {
    enumerable: true,
    get() {
      calls++;
      return [];
    },
  });
  const cyclic = mutable(input);
  cyclic.projection.extra = cyclic.projection;
  const missing = mutable(input);
  delete missing.projection.cells;
  const cases = [
    altered,
    cyclic,
    missing,
    { ...input, revision: { ...input.revision, revision: 99 } },
  ];
  for (const value of cases) {
    const result = bindLogicalSource(value, sourceDefinitions);
    expect(result.success).toBe(false);
    expect(Object.hasOwn(result, 'source')).toBe(false);
  }
  expect(calls).toBe(0);
  const empty = await inputFor('');
  const malformed = bindLogicalSource({
    ...empty,
    projection: { ...empty.projection, text: '<activity>' },
  });
  expect(malformed.success).toBe(false);
});

test('outputs detach and freeze every field while original source metadata stays mutable and unchanged', async () => {
  const input = mutable(await fixture),
    before = JSON.stringify(input.projection),
    source = bound(input);
  const frozen = (value: unknown): boolean =>
    value === null ||
    typeof value !== 'object' ||
    (Object.isFrozen(value) && Object.values(value).every(frozen));
  expect(frozen(source)).toBe(true);
  expect(JSON.stringify(input.projection)).toBe(before);
  expect(Object.isFrozen(input.projection)).toBe(false);
  input.projection.text = 'Changed';
  input.projection.cells[0].id = 'Changed';
  expect(source.parsed.sourceMarkdown).not.toBe('Changed');
  expect(source.bindings.some((entry) => entry.table?.cellId === 'Changed')).toBe(false);
  expect(JSON.parse(JSON.stringify(source))).toEqual(source);
});

test('three native review views bind imported body owners without changing protected parts or package bytes', async () => {
  const loaded = readOoxmlPackage(protectedDocument());
  if (!loaded.ok) throw new Error(loaded.reason);
  const bytes = writeOoxmlPackage(loaded.package),
    xml = canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!);
  for (const mode of ['original', 'proposed', 'all-markup'] as const) {
    const projection = await projectSourcePackage(loaded.package, mode),
      source = bound({ revision: sourceRevision, pkg: loaded.package, projection });
    expect(source.parsed.sourceMarkdown).toBe(projection.text);
    expect(source.index.partName).toBe(loaded.package.mainDocumentPart);
    expect(writeOoxmlPackage(loaded.package)).toEqual(bytes);
    expect(
      canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
    ).toBe(xml);
  }
});

test('binding and refusals preserve live bytes, revision, selection, imported XML, and undo/redo history', async () => {
  const input = await fixture,
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
    const projection = await projectSourcePackage(session.currentPackage()),
      value = { revision: sourceRevision, pkg: session.currentPackage(), projection },
      source = bound(value);
    expect(source.bindings.length).toBeGreaterThan(20);
    expect(
      bindLogicalSource({ ...value, revision: { ...sourceRevision, revision: 99 } }).success
    ).toBe(false);
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

test('HTML source spans bind virtual slots to actual native owners without invented source extents', async () => {
  const bytes = docxFixture(
    '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>' +
      sourceParagraph('Merged') +
      '</w:tc></w:tr><w:tr><w:tc>' +
      sourceParagraph('A') +
      '</w:tc><w:tc>' +
      sourceParagraph('B') +
      '</w:tc></w:tr></w:tbl>'
  );
  const editor = createDocxEditor({ container: document.createElement('div'), document: bytes });
  try {
    const session = editor.surface!.session,
      before = session.currentPackage(),
      revision = readEditorRevision(editor, sourceRevision.documentId, sourceRevision.epoch)!;
    const view = await readWorkingView(session, revision, { tableSyntax: 'auto' });
    if (!view.success) throw new Error(view.error);
    const input = { pkg: before, revision, projection: view.view.projection },
      source = bound(input);
    const cells = source.bindings.filter(
      (entry) => blockAt(source, entry.logicalPath).kind === 'cell'
    );
    expect(cells.map((entry) => entry.table?.role)).toEqual([
      'cell',
      'horizontal-placeholder',
      'cell',
      'cell',
    ]);
    expect(cells[1]!.source).toBeNull();
    expect(cells[1]!.contributions).toEqual([]);
    expect(cells[1]!.structuralOwnerId).toBeNull();
    expect(cells[1]!.table?.cellId).toBe(cells[0]!.table?.cellId);
    expect(findNode(session.part(), cells[1]!.table!.cellId!)).not.toBeNull();
    expect(cells[1]!.logicalPath).not.toEqual(cells[1]!.astPath);
    expect(session.currentPackage()).toBe(before);
    expect(readEditorRevision(editor, sourceRevision.documentId, sourceRevision.epoch)).toEqual(
      revision
    );
    const damaged = mutable(input);
    damaged.projection.cells[0].gridSpan = 1;
    expect(bindLogicalSource(damaged, sourceDefinitions)).toMatchObject({
      success: false,
      code: 'invalid-document',
    });
    expect(session.currentPackage()).toBe(before);
  } finally {
    editor.destroy();
  }
}, 30_000);
