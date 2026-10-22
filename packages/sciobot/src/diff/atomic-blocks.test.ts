import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  applyTreeOp,
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  escapeXml,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import oracle from '../__tests__/fixtures/atomic-assets.json';
import mdxOracle from '../__tests__/fixtures/mdx-tags.json';
import { docxFixture } from '../../test/support/document';
import { buildBlockIndex } from '../blocks/block-index';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import {
  DEFAULT_PAYLOAD_LIMITS,
  type BlockPayload,
  type JsonObject,
} from '../blocks/payload-schema';
import { MDX_BLOCK, MDX_INLINE, projectMdx, serializeOpaqueMdx } from '../projection/mdx';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { planAtomicBlock } from './atomic-blocks';
import { planDomainContainers } from './domain-containers';
import { directParagraphTextOps, planParagraphText } from './paragraph-text';
import {
  decodeEditPlan,
  serializeEditPlan,
  type CanonicalEditPlan,
  type EditPlanResult,
} from './plan';

const revision = { documentId: 'atomic', epoch: 'load-1', revision: 2 };
const p = (text: string) => `<w:p><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
const control = (id: string, body: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${body}</w:sdtContent></w:sdt>`;
const clone = <T>(value: T): T => JSON.parse(JSON.stringify(value));
const visual: BlockPayload = {
  id: 'visual',
  type: 'lesson_visual',
  version: 1,
  attributes: {
    kind: 'explainer',
    phase: 'evocation',
    src: 'vfs://pristine.png',
    fileId: 'pristine-file-1',
    alt: 'A diagram',
    prompt: 'Show the process.',
  },
  extension: { originalAsset: 'keep-this' },
};
function ast(markdown: string): TargetAstNode[] {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  return clone(parsed.target.normalizedAst.children) as TargetAstNode[];
}
const listening: BlockPayload = {
  id: 'audio',
  type: MDX_BLOCK,
  version: 1,
  attributes: { cache: { asset: 'cached-audio-1' } },
  mdast: ast(
    '<listening key="budget-1" title="Market" audio="hash-1" seed="7">\n\n<speaker id="s1" name="Anna" voice="voice-1" portrait="https://example.com/anna.png" />\n\n<line speaker="s1">Good morning!</line>\n\n</listening>'
  )[0]!,
  opaqueExtension: { codecVersion: 3, unknown: [1, null, { keep: true }] },
};
function fixture(payloads: readonly BlockPayload[] = [visual, listening], inline = false) {
  const blocks = payloads
    .map((payload) =>
      control(
        payload.id,
        inline ? '<w:r><w:t>Readable atomic fallback.</w:t></w:r>' : p('Readable atomic fallback.')
      )
    )
    .join('');
  const body = inline
    ? `<w:p><w:r><w:t>Before </w:t></w:r>${blocks}<w:r><w:t> After</w:t></w:r></w:p>`
    : p('Old nearby sentence.') + blocks + p('Keep tail.');
  const loaded = readOoxmlPackage(docxFixture(body));
  if (!loaded.ok) throw new Error(loaded.reason);
  let pkg = loaded.package;
  for (const payload of payloads) {
    const stored = withPayloadRecord(pkg, payload, 'Readable atomic fallback.');
    if (!stored.ok) throw new Error(stored.reason);
    pkg = stored.pkg;
  }
  const part = pkg.parts.get(pkg.mainDocumentPart)!;
  return { pkg, part, source: (blockId: string) => ({ pkg, revision, blockId }) };
}
function success(result: EditPlanResult): CanonicalEditPlan {
  if (!result.success) throw new Error(result.error);
  return result.plan;
}
/** Immutable payload codec oracle; executors also update native fallback and review state. */
function applyPayload(pkg: OoxmlPackage, plan: CanonicalEditPlan, side: 'before' | 'after') {
  let next = pkg;
  for (const edit of plan.operations) {
    if (edit.kind !== 'payload' || edit.path.length || !edit[side].present)
      throw new Error('Expected one root replacement');
    const payload = edit[side].value as BlockPayload;
    const stored = withPayloadRecord(next, payload, 'Readable atomic fallback.');
    if (!stored.ok) throw new Error(stored.reason);
    next = stored.pkg;
  }
  return next;
}
function stored(pkg: OoxmlPackage, id: string) {
  return buildBlockIndex(pkg).byId.get(id)!.payload;
}
function listeningPayload(node: JsonObject): BlockPayload {
  return {
    id: 'audio',
    type: 'listening',
    version: 1,
    attributes: {},
    ...Object.fromEntries(
      Object.entries(node).filter(([key]) => !['type', 'children', 'id'].includes(key))
    ),
  };
}

test('a changed visual produces one whole-payload replacement and retains pristine image metadata', () => {
  const value = fixture(),
    before = stored(value.pkg, 'visual'),
    target = { ...before, attributes: { ...before.attributes, alt: 'An improved diagram' } };
  const plan = success(planAtomicBlock(value.source('visual'), target));
  const block = buildBlockIndex(value.pkg).byId.get('visual')!;
  expect(plan.operations).toEqual([
    {
      kind: 'payload',
      partName: value.part.name,
      blockId: 'visual',
      controlId: block.controlId,
      path: [],
      before: { present: true, value: before },
      after: { present: true, value: target },
    },
  ]);
  const changed = applyPayload(value.pkg, plan, 'after');
  expect(stored(changed, 'visual').attributes.fileId).toBe('pristine-file-1');
  expect(stored(changed, 'visual').attributes.src).toBe('vfs://pristine.png');
  expect(stored(changed, 'visual').extension).toEqual({ originalAsset: 'keep-this' });
  expect(changed.parts.get(value.part.name) === value.part).toBe(true);
  expect(stored(applyPayload(changed, plan, 'before'), 'visual')).toEqual(before);
  const reopened = readOoxmlPackage(writeOoxmlPackage(changed));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(stored(reopened.package, 'visual')).toEqual(target);
});

test('an intentional visual asset replacement carries the requested pristine ID and source together', () => {
  const value = fixture(),
    target = clone(visual);
  const attrs = {
    ...target.attributes,
    src: 'vfs://replacement.png',
    fileId: 'pristine-file-2',
    prompt: 'A new process.',
  };
  const plan = success(planAtomicBlock(value.source('visual'), { ...target, attributes: attrs }));
  expect(plan.operations).toHaveLength(1);
  expect(stored(applyPayload(value.pkg, plan, 'after'), 'visual').attributes).toEqual(attrs);
  expect(stored(value.pkg, 'visual').attributes.fileId).toBe('pristine-file-1');
});

test('unchanged atomic data ignores object key order and retains original payloads without edits', () => {
  const value = fixture(),
    original = stored(value.pkg, 'visual');
  const target = Object.fromEntries(Object.entries(original).reverse()) as BlockPayload;
  expect(success(planAtomicBlock(value.source('visual'), target)).operations).toEqual([]);
  expect(
    success(planAtomicBlock(value.source('audio'), clone(stored(value.pkg, 'audio')))).operations
  ).toEqual([]);
});

test.each(oracle.cases)('atomic listening plans match executed reconciliation: $name', (item) => {
  const before = listeningPayload(item.before as unknown as JsonObject),
    after = listeningPayload(item.after as unknown as JsonObject);
  const value = fixture([before]),
    plan = success(planAtomicBlock(value.source('audio'), after));
  expect(plan.operations).toHaveLength(item.rootChanges.length ? 1 : 0);
  if (item.rootChanges.length) {
    expect(item.rootChanges).toEqual(['delete', 'insert']);
    expect(plan.operations[0]).toMatchObject({
      kind: 'payload',
      path: [],
      before: { present: true, value: before },
      after: { present: true, value: after },
    });
  }
  const proposed = applyPayload(value.pkg, plan, 'after');
  expect(stored(proposed, 'audio')).toEqual(after);
  expect(stored(applyPayload(proposed, plan, 'before'), 'audio')).toEqual(before);
  expect((stored(proposed, 'audio').exercise as JsonObject).tasks).toEqual(
    (before.exercise as JsonObject).tasks
  );
  expect((stored(proposed, 'audio').exercise as JsonObject).speakers).toEqual(
    (after.exercise as JsonObject).speakers
  );
  expect(stored(proposed, 'audio').blockKey).toBe('budget-keep');
});

test('source chips replace atomically in inline controls without changing native surrounding text', () => {
  const source: BlockPayload = {
    id: 'chip',
    type: 'source',
    version: 1,
    attributes: {
      url: 'https://example.com',
      title: 'A book',
      author: 'Author',
      bookId: 42,
      trust: 'verified',
    },
  };
  const value = fixture([source], true),
    target = { ...source, attributes: { ...source.attributes, url: 'https://example.org' } };
  const plan = success(planAtomicBlock(value.source('chip'), target));
  expect(plan.operations).toHaveLength(1);
  const proposed = applyPayload(value.pkg, plan, 'after');
  expect(proposed.parts.get(value.part.name) === value.part).toBe(true);
  expect(stored(proposed, 'chip').attributes).toEqual(target.attributes);
});

test('opaque parser positions create no edits while data positions, literal IDs, and unknown fields remain meaningful', () => {
  const original = clone(listening),
    mdast = original.mdast as JsonObject;
  const value = fixture([original]);
  const relocated = {
    ...original,
    mdast: {
      ...mdast,
      position: { start: { line: 90, column: 1 }, end: { line: 99, column: 2 } },
      attributes: (mdast.attributes as JsonObject[]).map((attr) => ({
        ...attr,
        position: { start: { line: 90, column: 2 }, end: { line: 90, column: 3 } },
      })),
    },
  };
  expect(success(planAtomicBlock(value.source('audio'), relocated)).operations).toEqual([]);
  const expression: BlockPayload = { ...original, mdast: ast('<listening title={"Text"} />')[0]! };
  const expressionValue = fixture([expression]),
    expressionAst = expression.mdast as JsonObject;
  const expressionMoved = {
    ...expression,
    mdast: {
      ...expressionAst,
      attributes: (expressionAst.attributes as JsonObject[]).map((attr) => ({
        ...attr,
        value: {
          ...(attr.value as JsonObject),
          position: { start: { line: 50, column: 2 }, end: { line: 50, column: 3 } },
        },
      })),
    },
  };
  expect(
    success(planAtomicBlock(expressionValue.source('audio'), expressionMoved)).operations
  ).toEqual([]);
  const data: BlockPayload = {
    ...original,
    mdast: { ...mdast, data: { type: 'opaque-metadata', position: { value: 'first' } } },
  };
  const dataValue = fixture([data]);
  expect(
    success(
      planAtomicBlock(dataValue.source('audio'), {
        ...data,
        mdast: {
          ...(data.mdast as JsonObject),
          data: { type: 'opaque-metadata', position: { value: 'second' } },
        },
      })
    ).operations
  ).toHaveLength(1);
  const withId: BlockPayload = {
    ...original,
    mdast: {
      ...mdast,
      attributes: [
        ...(mdast.attributes as JsonObject[]),
        { type: 'mdxJsxAttribute', name: 'id', value: 'literal-a' },
      ],
    },
  };
  const ids = fixture([withId]);
  const changedId: BlockPayload = {
    ...withId,
    mdast: {
      ...(withId.mdast as JsonObject),
      attributes: [
        ...(mdast.attributes as JsonObject[]),
        { type: 'mdxJsxAttribute', name: 'id', value: 'literal-b' },
      ],
    },
  };
  expect(success(planAtomicBlock(ids.source('audio'), changedId)).operations).toHaveLength(1);
  expect(
    success(planAtomicBlock(value.source('audio'), { ...original, opaqueExtension: null }))
      .operations
  ).toHaveLength(1);
});

test('opaque listening child changes produce one root replacement without planning native fallback descendants', () => {
  const value = fixture(),
    before = stored(value.pkg, 'audio');
  const target = {
    ...before,
    mdast: ast(
      '<listening key="budget-1" title="Changed" audio="new-hash" seed="8">\n\n<speaker id="s1" name="Anna" voice="voice-1" />\n\n<line speaker="s1">Good evening!</line>\n\n</listening>'
    )[0]!,
  };
  const plan = success(planAtomicBlock(value.source('audio'), target));
  expect(plan.operations).toHaveLength(1);
  expect(plan.operations[0]).toMatchObject({
    kind: 'payload',
    path: [],
    after: { present: true, value: target },
  });
  expect(stored(applyPayload(value.pkg, plan, 'after'), 'audio').opaqueExtension).toEqual(
    before.opaqueExtension
  );
});

test('complete opaque MDX expressions, spreads, nesting, and prototype-shaped metadata stay inert and detached', () => {
  const captured = mdxOracle.cases.find((item) => item.id === 'nested-listening')!;
  const payload: BlockPayload = {
    id: 'audio',
    type: MDX_BLOCK,
    version: 1,
    attributes: {},
    mdast: captured.ast as JsonObject,
  };
  const value = fixture([payload]);
  expect(success(planAtomicBlock(value.source('audio'), clone(payload))).operations).toEqual([]);
  const target = {
    ...payload,
    mdast: ast(
      '<listening {...globalThis.__atomicSpread} title={globalThis.__atomicCall?.()}>\n\n<line speaker="s1">{globalThis.__atomicCall?.()}</line>\n\n</listening>'
    )[0]!,
    extension: JSON.parse(
      '{"__proto__":{"polluted":true},"constructor":{"prototype":"data"}}'
    ) as JsonObject,
  };
  let calls = 0;
  const globals = globalThis as unknown as Record<string, unknown>;
  globals.__atomicCall = () => {
    calls++;
  };
  Object.defineProperty(globals, '__atomicSpread', {
    configurable: true,
    get() {
      calls++;
      return {};
    },
  });
  try {
    const plan = success(planAtomicBlock(value.source('audio'), target));
    expect(plan.operations).toHaveLength(1);
    expect(calls).toBe(0);
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
    expect(decodeEditPlan(serializeEditPlan(plan))).toEqual({ success: true, plan });
    const serialized = serializeEditPlan(plan);
    (target.extension as { constructor: { prototype: string } }).constructor.prototype = 'changed';
    expect(serializeEditPlan(plan)).toBe(serialized);
    expect(Object.isFrozen(target.mdast)).toBe(false);
    expect(Object.isFrozen(plan.operations[0])).toBe(true);
  } finally {
    delete globals.__atomicCall;
    delete globals.__atomicSpread;
  }
});

test('inline opaque AST mode stays inline and a content edit remains one replacement', () => {
  const parsed = ast('Before <source title="Book" url="https://example.com" /> after.')[0]!;
  const mdast = (parsed.children as TargetAstNode[]).find(
    (child) => child.type === 'mdxJsxTextElement'
  )!;
  const payload: BlockPayload = {
    id: 'inline',
    type: MDX_INLINE,
    version: 1,
    attributes: {},
    mdast,
  };
  const value = fixture([payload], true),
    target = {
      ...payload,
      mdast: {
        ...mdast,
        attributes: [{ type: 'mdxJsxAttribute', name: 'title', value: 'Other book' }],
      },
    };
  expect(success(planAtomicBlock(value.source('inline'), target)).operations).toHaveLength(1);
  expect(planAtomicBlock(value.source('inline'), { ...target, type: MDX_BLOCK }).success).toBe(
    false
  );
});

test('nearby sentence edits preserve untouched opaque audio, pristine image IDs, native controls, and all other parts', async () => {
  const value = fixture(),
    opened = openDocumentForExport(writeOoxmlPackage(value.pkg));
  if (!opened.ok) throw new Error(opened.reason);
  let markdown: string;
  try {
    const projected = projectMdx(value.pkg, await opened.session.layoutFor('proposed'), revision);
    if (!projected.success) throw new Error(projected.error);
    markdown = projected.projection.text;
  } finally {
    opened.session.dispose();
  }
  const before = ast(markdown),
    target = ast(markdown.replace('Old nearby sentence', 'New nearby sentence')),
    root = bodyStoryRoot(value.part)!;
  if (root.kind === 'textValue') throw new Error('Missing body');
  const result = planDomainContainers(
    {
      pkg: value.pkg,
      revision,
      partName: value.part.name,
      parentId: root.id,
      startIndex: 0,
      blocks: before.map((ast, at) => ({ id: root.children[at]!.id, ast })),
    },
    target
  );
  if (!result.success) throw new Error(result.error);
  expect(result.plan.operations).toEqual([]);
  expect(
    result.pairs
      .filter((pair) => pair.kind === 'atomic')
      .map((pair) => [pair.blockId, pair.changed])
  ).toEqual([
    ['visual', false],
    ['audio', false],
  ]);
  for (const blockId of ['visual', 'audio'])
    expect(
      success(planAtomicBlock(value.source(blockId), clone(stored(value.pkg, blockId)))).operations
    ).toEqual([]);
  const paragraphId = root.children[0]!.id;
  const text = success(
    planParagraphText(
      {
        revision,
        partName: value.part.name,
        paragraphId,
        text: paragraphTextOf(value.part, paragraphId)!,
      },
      'New nearby sentence.'
    )
  );
  let part = value.part;
  for (const edit of text.operations) {
    if (edit.kind !== 'text') throw new Error('Unexpected text plan');
    for (const op of directParagraphTextOps(edit)) {
      const changed = applyTreeOp(part, op);
      if (!changed.ok) throw new Error(changed.reason);
      part = changed.part;
    }
  }
  const proposed = withPart(value.pkg, part);
  for (const id of ['visual', 'audio']) {
    const block = buildBlockIndex(value.pkg).byId.get(id)!;
    expect(findNode(part, block.controlId) === findNode(value.part, block.controlId)).toBe(true);
    expect(stored(proposed, id)).toEqual(stored(value.pkg, id));
  }
  for (const [name, native] of value.pkg.parts)
    if (name !== value.part.name) expect(proposed.parts.get(name) === native).toBe(true);
  const reopened = readOoxmlPackage(writeOoxmlPackage(proposed));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(stored(reopened.package, 'visual').attributes.fileId).toBe('pristine-file-1');
  expect(serializeOpaqueMdx(stored(reopened.package, 'audio')).text).toContain(
    'audio="hash-1" seed="7"'
  );
  expect(stored(reopened.package, 'audio').opaqueExtension).toEqual(listening.opaqueExtension);
});

test('invalid target identity, syntax, versions, data, and resource excess refuse without native writes', () => {
  const value = fixture(),
    bytes = writeOoxmlPackage(value.pkg);
  let calls = 0;
  const getter = Object.defineProperty({ ...visual }, 'attributes', {
    enumerable: true,
    get() {
      calls++;
      return {};
    },
  });
  const cycle: Record<string, unknown> = {};
  cycle.self = cycle;
  let deep: unknown = 'leaf';
  for (let at = 0; at < 70; at++) deep = { child: deep };
  const cases: unknown[] = [
    null,
    {},
    { ...visual, id: 'other' },
    { ...visual, type: 'source' },
    { ...visual, version: 2 },
    getter,
    { ...visual, attributes: { value: undefined } },
    { ...visual, attributes: cycle },
    { ...visual, extension: deep },
    { ...visual, attributes: { alt: 'x'.repeat(DEFAULT_PAYLOAD_LIMITS.maxBytes) } },
  ];
  for (const target of cases)
    expect(planAtomicBlock(value.source('visual'), target as BlockPayload).success).toBe(false);
  expect(calls).toBe(0);
  const sourceGetter = Object.defineProperty(value.source('visual'), 'blockId', {
    enumerable: true,
    get() {
      calls++;
      return 'visual';
    },
  });
  expect(planAtomicBlock(sourceGetter, visual).success).toBe(false);
  expect(calls).toBe(0);
  expect(planAtomicBlock(value.source('missing'), visual)).toMatchObject({
    success: false,
    code: 'invalid-document',
  });
  for (const type of ['lesson_activity', 'source_excerpt', 'unknown']) {
    const payload = { ...visual, type },
      container = fixture([payload]);
    expect(planAtomicBlock(container.source('visual'), payload)).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
  }
  const invalid = {
    ...listening,
    mdast: { ...(listening.mdast as JsonObject), name: 'invalid tag' },
  };
  expect(planAtomicBlock(value.source('audio'), invalid)).toMatchObject({
    success: false,
    code: 'invalid-input',
  });
  const badSource = fixture([invalid]);
  expect(planAtomicBlock(badSource.source('audio'), listening)).toMatchObject({
    success: false,
    code: 'invalid-document',
  });
  expect(writeOoxmlPackage(value.pkg)).toEqual(bytes);
});

test('planning leaves the live package, revision, selection, and existing undo/redo state intact', () => {
  const value = fixture(),
    editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(value.pkg),
    });
  try {
    const surface = editor.surface!,
      session = surface.session,
      paragraphId = session.paragraphIds()[0]!;
    surface.applyAutomationOps(() => [{ op: 'insertText', paragraphId, offset: 0, text: 'User ' }]);
    const pkg = session.currentPackage(),
      bytes = writeOoxmlPackage(pkg),
      current = session.packageRevision(),
      selection = surface.state().selection,
      history = [session.canUndo(), session.canRedo()],
      fingerprint = canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
    const result = planAtomicBlock(
      { pkg, revision: { ...revision, revision: current }, blockId: 'visual' },
      { ...visual, attributes: { ...visual.attributes, alt: 'New description' } }
    );
    expect(result.success).toBe(true);
    expect(session.packageRevision()).toBe(current);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
    expect(
      canonicalOoxmlFingerprint(session.currentPackage().parts.get(pkg.mainDocumentPart)!)
    ).toBe(fingerprint);
  } finally {
    editor.destroy();
  }
});
