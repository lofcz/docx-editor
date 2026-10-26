import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  applyTreeOp,
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  contentControlPropertiesNodeOf,
  contentControlPropertiesOf,
  escapeXml,
  findNode,
  insertChildren,
  paragraphTextOf,
  readOoxmlPackage,
  readOoxmlPart,
  withPart,
  writeOoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import pairingFixture from '../__tests__/fixtures/domain-pairing.json';
import { docxFixture, W } from '../../test/support/document';
import { buildBlockIndex } from '../blocks/block-index';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import type { BlockPayload } from '../blocks/payload-schema';
import { projectMdx, MDX_BLOCK } from '../projection/mdx';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { GENERATED_LABEL_TAG } from '../projection/block-tags';
import { directParagraphTextOps, planParagraphText } from './paragraph-text';
import {
  planDomainContainers,
  domainNodeFromAst,
  type DomainContainerSource,
  type MappedDomainBlock,
} from './domain-containers';
import { sameContainerIdentity } from './domain-strategies';
import { decodeEditPlan, serializeEditPlan } from './plan';

const revision = { documentId: 'domain', epoch: 'load-1', revision: 4 };
const p = (text: string) =>
  `<w:p><w:pPr><w:customKeep w:val="yes"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
const control = (id: string, body: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${body}</w:sdtContent></w:sdt>`;
const label = `<w:sdt><w:sdtPr><w:tag w:val="${GENERATED_LABEL_TAG}"/></w:sdtPr><w:sdtContent>${p('Generated label')}</w:sdtContent></w:sdt>`;
const children = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;
const payload = (
  id: string,
  type: string,
  attributes: BlockPayload['attributes'] = {}
): BlockPayload => ({ id, type, attributes, version: 1 });
function parsed(markdown: string): TargetAstNode[] {
  const result = parseTargetMarkdown(markdown);
  if (!result.success) throw new Error(result.error);
  return JSON.parse(JSON.stringify(result.target.normalizedAst.children));
}
async function fixture(options: { audio?: boolean } = {}) {
  const body =
    control(
      'phase',
      label +
        control(
          'activity',
          control('excerpt', p('Old question.') + p('Keep quotation.')) + p('Keep activity.')
        )
    ) +
    (options.audio ? control('audio', p('Native audio fallback')) : '') +
    p('Outside.');
  const loaded = readOoxmlPackage(docxFixture(body));
  if (!loaded.ok) throw new Error(loaded.reason);
  let pkg = loaded.package;
  const payloads = [
    payload('phase', 'lesson_phase', { phaseType: 'evocation' }),
    payload('activity', 'lesson_activity', { name: 'Warmup', duration: 10, resource: 'A' }),
    payload('excerpt', 'source_excerpt', { url: 'https://example.com', trust: 'verified' }),
  ];
  if (options.audio)
    payloads.push({
      ...payload('audio', MDX_BLOCK),
      mdast: parsed(
        '<listening id="first" title="Audio">\n\n<line>Stored audio.</line>\n\n</listening>'
      )[0]!,
    });
  for (const value of payloads) {
    const written = withPayloadRecord(pkg, value, '');
    if (!written.ok) throw new Error(written.reason);
    pkg = written.pkg;
  }
  const opened = openDocumentForExport(writeOoxmlPackage(pkg));
  if (!opened.ok) throw new Error(opened.reason);
  let markdown: string;
  try {
    const result = projectMdx(pkg, await opened.session.layoutFor('proposed'), revision);
    if (!result.success) throw new Error(result.error);
    markdown = result.projection.text;
  } finally {
    opened.session.dispose();
  }
  const ast = parsed(markdown),
    part = pkg.parts.get(pkg.mainDocumentPart)!,
    root = bodyStoryRoot(part)!;
  const indexed = buildBlockIndex(pkg),
    controls = new Map([...indexed.byId.values()].map((block) => [block.controlId, block]));
  const window = (parent: OoxmlNode, nodes: readonly TargetAstNode[]) => {
    const all = children(parent),
      startIndex =
        all[0]?.kind === 'contentControl' &&
        contentControlPropertiesOf(all[0]!).tag === GENERATED_LABEL_TAG
          ? 1
          : 0;
    const blocks: MappedDomainBlock[] = nodes.map((ast, at) => {
      const native = all[startIndex + at]!;
      if (!native) throw new Error('Fixture native mapping is incomplete');
      const owner = controls.get(native.id),
        content = owner && contentControlContentNodeOf(native);
      return {
        id: native.id,
        ast,
        ...(content && owner?.payload.type !== MDX_BLOCK
          ? { children: window(content, (ast.children as TargetAstNode[]) ?? []) }
          : {}),
      };
    });
    return { parentId: parent.id, startIndex, blocks };
  };
  const source: DomainContainerSource = {
    revision,
    pkg,
    partName: part.name,
    ...window(root, ast),
  };
  return { pkg, part, source, markdown, ast, indexed };
}
function success(result: ReturnType<typeof planDomainContainers>) {
  if (!result.success) throw new Error(result.error);
  return result;
}
function atPath(nodes: readonly TargetAstNode[], path: readonly number[]): TargetAstNode {
  let node = nodes[path[0]!]!;
  for (const at of path.slice(1)) node = (node.children as TargetAstNode[])[at]!;
  return node;
}
function text(node: TargetAstNode): string {
  return typeof node.value === 'string'
    ? node.value
    : Array.isArray(node.children)
      ? (node.children as TargetAstNode[]).map(text).join('')
      : '';
}

test.each(pairingFixture.cases)(
  'container identity matches the executed application: $name',
  (value) => {
    expect(sameContainerIdentity(value.before, value.after)).toBe(value.retainsContainer);
  }
);

test('one nested question retains phase, activity, citation, labels, and neighboring native identities', async () => {
  const value = await fixture(),
    target = parsed(value.markdown.replace('Old question', 'Edited question'));
  const before = canonicalOoxmlFingerprint(value.part),
    result = success(planDomainContainers(value.source, target));
  expect(result.plan.operations).toEqual([]);
  expect(
    result.pairs.filter((pair) => pair.kind === 'container').map((pair) => pair.blockId)
  ).toEqual(['phase', 'activity', 'excerpt']);
  const leaves = result.pairs.filter((pair) => pair.kind === 'content' && pair.changed);
  expect(leaves).toHaveLength(1);
  const leaf = leaves[0]!;
  expect(text(atPath(target, leaf.afterPath))).toBe('Edited question.');
  const oldText = paragraphTextOf(value.part, leaf.id)!;
  const planned = planParagraphText(
    { revision, partName: value.part.name, paragraphId: leaf.id, text: oldText },
    text(atPath(target, leaf.afterPath))
  );
  if (!planned.success) throw new Error(planned.error);
  let part = value.part;
  for (const op of planned.plan.operations.flatMap((edit) =>
    edit.kind === 'text' ? directParagraphTextOps(edit) : []
  )) {
    const changed = applyTreeOp(part, op);
    if (!changed.ok) throw new Error(changed.reason);
    part = changed.part;
  }
  for (const pair of result.pairs) {
    if (pair.kind === 'container') {
      const old = findNode(value.part, pair.id)!,
        next = findNode(part, pair.id)!;
      expect(next.id).toBe(old.id);
      expect(contentControlPropertiesNodeOf(next) === contentControlPropertiesNodeOf(old)).toBe(
        true
      );
      expect(contentControlContentNodeOf(next)!.id).toBe(contentControlContentNodeOf(old)!.id);
    } else if (!pair.changed)
      expect(findNode(part, pair.id) === findNode(value.part, pair.id)).toBe(true);
  }
  expect(paragraphTextOf(part, leaf.id)).toBe('Edited question.');
  expect(canonicalOoxmlFingerprint(value.part)).toBe(before);
  expect(
    [...withPart(value.pkg, part).parts]
      .filter(([name]) => name !== value.part.name)
      .map(([name, p]) => [name, canonicalOoxmlFingerprint(p)])
  ).toEqual(
    [...value.pkg.parts]
      .filter(([name]) => name !== value.part.name)
      .map(([name, p]) => [name, canonicalOoxmlFingerprint(p)])
  );
  expect(decodeEditPlan(serializeEditPlan(result.plan))).toEqual({
    success: true,
    plan: result.plan,
  });
});

test('phaseType changes produce one structural replacement without edits inside the replaced wrapper', async () => {
  const value = await fixture(),
    target = parsed(value.markdown.replace('type="evocation"', 'type="reflection"'));
  const result = success(planDomainContainers(value.source, target)),
    phase = value.indexed.byId.get('phase')!;
  expect(result.plan.operations).toHaveLength(1);
  expect(result.plan.operations[0]).toMatchObject({
    kind: 'siblings',
    parentId: value.source.parentId,
    removedIds: [phase.controlId],
  });
  const edit = result.plan.operations[0]!;
  if (edit.kind !== 'siblings') throw new Error('Expected structural change');
  expect(domainNodeFromAst(edit.inserted[0]!)?.phaseType).toBe('reflection');
  expect(
    result.pairs.some((pair) => ['phase', 'activity', 'excerpt'].includes(pair.blockId ?? ''))
  ).toBe(false);
  expect(JSON.parse(serializeEditPlan(result.plan)).operations[0].inserted[0].type).toBe(
    'mdxJsxFlowElement'
  );
});

test('nested insertion anchors inside citation content and preserves surrounding native objects', async () => {
  const value = await fixture(),
    target = parsed(value.markdown);
  (atPath(target, [0, 0, 0]).children as TargetAstNode[]).splice(1, 0, {
    type: 'paragraph',
    children: [{ type: 'text', value: 'Added question.' }],
  });
  const result = success(planDomainContainers(value.source, target));
  expect(result.plan.operations).toHaveLength(1);
  const op = result.plan.operations[0]!;
  if (op.kind !== 'siblings') throw new Error('Expected child insertion');
  const excerpt = findNode(value.part, value.indexed.byId.get('excerpt')!.controlId)!,
    content = contentControlContentNodeOf(excerpt)!;
  expect(op).toMatchObject({
    parentId: content.id,
    removedIds: [],
    beforeId: children(content)[0]!.id,
    afterId: children(content)[1]!.id,
  });
  expect(op.inserted.map(text)).toEqual(['Added question.']);
  const loaded = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>${escapeXml(text(op.inserted[0]!))}</w:t></w:r></w:p></w:body></w:document>`,
    { name: '/test/new-domain.xml', contentType: 'application/xml' }
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const inserted = insertChildren(value.part, content.id, 1, children(bodyStoryRoot(loaded.part)!));
  if (!inserted.ok) throw new Error('Native insertion failed');
  for (const node of children(content))
    expect(findNode(inserted.part, node.id) === node).toBe(true);
  expect(
    result.pairs.filter((pair) => pair.kind === 'container').map((pair) => pair.blockId)
  ).toEqual(['phase', 'activity', 'excerpt']);
});

test('activity updates remain explicit container pairs while a resource-only change replaces the activity', async () => {
  const value = await fixture();
  const updates = success(
    planDomainContainers(
      value.source,
      parsed(
        value.markdown
          .replace('name="Warmup"', 'name="Opening"')
          .replace('duration="10"', 'duration="15"')
      )
    )
  );
  expect(updates.plan.operations).toEqual([]);
  expect(updates.pairs.find((pair) => pair.blockId === 'activity')).toMatchObject({
    kind: 'container',
    changed: true,
  });
  expect(
    updates.pairs.filter((pair) => pair.kind === 'content').every((pair) => !pair.changed)
  ).toBe(true);
  const resource = success(
    planDomainContainers(
      value.source,
      parsed(value.markdown.replace('resource="A"', 'resource="B"'))
    )
  );
  expect(resource.plan.operations).toHaveLength(1);
  expect(resource.plan.operations[0]).toMatchObject({
    kind: 'siblings',
    removedIds: [value.indexed.byId.get('activity')!.controlId],
  });
  expect(resource.pairs.some((pair) => pair.blockId === 'activity')).toBe(false);
});

test('citation identity changes replace only the excerpt and preserve its outer domain owners', async () => {
  const value = await fixture();
  const result = success(
    planDomainContainers(
      value.source,
      parsed(value.markdown.replace('https://example.com', 'https://other.example'))
    )
  );
  expect(result.plan.operations).toHaveLength(1);
  const owner = findNode(value.part, value.indexed.byId.get('activity')!.controlId)!;
  expect(result.plan.operations[0]).toMatchObject({
    kind: 'siblings',
    parentId: contentControlContentNodeOf(owner)!.id,
    removedIds: [value.indexed.byId.get('excerpt')!.controlId],
  });
  expect(
    result.pairs.filter((pair) => pair.kind === 'container').map((pair) => pair.blockId)
  ).toEqual(['phase', 'activity']);
});

test('stored opaque listening remains atomic and never pairs its native fallback children', async () => {
  const value = await fixture({ audio: true });
  const result = success(
    planDomainContainers(
      value.source,
      parsed(value.markdown.replace('title="Audio"', 'title="Changed audio"'))
    )
  );
  const audio = value.indexed.byId.get('audio')!;
  expect(result.pairs.find((pair) => pair.id === audio.controlId)).toMatchObject({
    kind: 'atomic',
    changed: true,
  });
  expect(result.plan.operations).toEqual([]);
  const content = contentControlContentNodeOf(findNode(value.part, audio.controlId)!)!;
  expect(result.pairs.some((pair) => children(content).some((node) => node.id === pair.id))).toBe(
    false
  );
  const attribute = success(
    planDomainContainers(value.source, parsed(value.markdown.replace('id="first"', 'id="second"')))
  );
  expect(attribute.pairs.find((pair) => pair.id === audio.controlId)).toMatchObject({
    kind: 'atomic',
    changed: true,
  });
});

test('attribute-only edits can retain unmapped unchanged child subtrees', async () => {
  const value = await fixture();
  const blocks = JSON.parse(JSON.stringify(value.source.blocks)) as MappedDomainBlock[];
  const activity = blocks[0]!.children!.blocks[0]!;
  delete (activity as { children?: unknown }).children;
  const result = success(
    planDomainContainers(
      { ...value.source, blocks },
      parsed(value.markdown.replace('name="Warmup"', 'name="Opening"'))
    )
  );
  expect(result.plan.operations).toEqual([]);
  expect(result.pairs.find((pair) => pair.blockId === 'activity')).toMatchObject({
    kind: 'container',
    changed: true,
  });
  expect(result.pairs.some((pair) => pair.blockId === 'excerpt')).toBe(false);
});

test('domain attribute decoding preserves inert expressions and native property aliases', () => {
  const phase = parsed(
    '<phase type="evocation"><activity name="A" duration="10">Text.</activity></phase>'
  )[0]!;
  const phaseNode = phase.type === 'paragraph' ? (phase.children as TargetAstNode[])[0]! : phase;
  expect(domainNodeFromAst(phaseNode)).toMatchObject({
    type: 'lesson_phase',
    phaseType: 'evocation',
  });
  const citation = parsed(
    '<source title="1984" bookid="7" type="web" url="https://example.com?a=1&amp;b=2">\n\nQuote.\n\n</source>'
  )[0]!;
  expect(domainNodeFromAst(citation)).toEqual({
    type: 'source_excerpt',
    bookId: 7,
    trust: 'verified',
    sourceType: 'web',
    url: 'https://example.com?a=1&b=2',
  });
  const expression = parsed(
    '<activity duration={globalThis.__domainProbe()}>\n\nText.\n\n</activity>'
  )[0]!;
  expect(domainNodeFromAst(expression)).toMatchObject({
    type: 'lesson_activity',
    duration: { type: 'mdxJsxAttributeValueExpression', value: 'globalThis.__domainProbe()' },
  });
});

test('exact containers can retain their whole subtree without a recursive map', async () => {
  const value = await fixture();
  const shallow = {
    ...value.source,
    blocks: value.source.blocks.map(({ id, ast }) => ({ id, ast })),
  };
  const result = success(planDomainContainers(shallow, value.ast));
  expect(result.plan.operations).toEqual([]);
  expect(
    result.pairs.filter((pair) => pair.kind === 'container').map((pair) => pair.blockId)
  ).toEqual(['phase']);
});

test('invalid child parents, omitted child maps, forged native types, getters, and excessive input fail atomically', async () => {
  const value = await fixture(),
    target = parsed(value.markdown.replace('Old question', 'New question'));
  const clone = () => JSON.parse(JSON.stringify(value.source.blocks)) as MappedDomainBlock[];
  const missing = {
    ...value.source,
    blocks: value.source.blocks.map(({ id, ast }) => ({ id, ast })),
  };
  expect(planDomainContainers(missing, target)).toMatchObject({
    success: false,
    code: 'ambiguous-target',
  });
  const wrong = clone();
  (wrong[0]!.children as { parentId: string }).parentId = value.source.parentId;
  expect(planDomainContainers({ ...value.source, blocks: wrong }, target)).toMatchObject({
    success: false,
    code: 'ambiguous-target',
  });
  const forged = clone();
  (forged[0]!.ast as unknown as { name: string }).name = 'activity';
  expect(planDomainContainers({ ...value.source, blocks: forged }, target)).toMatchObject({
    success: false,
    code: 'invalid-document',
  });
  let called = false;
  const getter = Object.defineProperty({}, 'type', {
    enumerable: true,
    get() {
      called = true;
      return 'paragraph';
    },
  }) as TargetAstNode;
  expect(planDomainContainers(value.source, [getter])).toMatchObject({
    success: false,
    code: 'invalid-input',
  });
  expect(called).toBe(false);
  expect(
    planDomainContainers(
      { ...value.source, blocks: [null] as unknown as MappedDomainBlock[] },
      target
    )
  ).toMatchObject({ success: false, code: 'invalid-input' });
  const broken = clone();
  (broken[0] as unknown as { children: unknown }).children = {
    parentId: 'x',
    startIndex: 0,
    blocks: null,
  };
  expect(planDomainContainers({ ...value.source, blocks: broken }, target)).toMatchObject({
    success: false,
    code: 'invalid-input',
  });
  const cyclic: Record<string, unknown> = { type: 'paragraph' };
  cyclic.children = [cyclic];
  expect(planDomainContainers(value.source, [cyclic as TargetAstNode])).toMatchObject({
    success: false,
    code: 'invalid-input',
  });
  let deep: TargetAstNode = { type: 'paragraph', children: [] };
  for (let at = 0; at < 70; at++) deep = { type: 'blockquote', children: [deep] };
  expect(planDomainContainers(value.source, [deep])).toMatchObject({
    success: false,
    code: 'invalid-input',
    error: 'depth-limit',
  });
});

test('plans detach inserted fragments and pair paths without retaining or freezing caller data', async () => {
  const value = await fixture(),
    target = parsed(value.markdown.replace('type="evocation"', 'type="reflection"'));
  const result = success(planDomainContainers(value.source, target));
  (target[0]!.attributes as unknown[]).length = 0;
  expect(Object.isFrozen(target[0])).toBe(false);
  const edit = result.plan.operations[0]!;
  if (edit.kind !== 'siblings') throw new Error('Missing replacement');
  expect(domainNodeFromAst(edit.inserted[0]!)?.phaseType).toBe('reflection');
  expect(
    result.pairs.every(
      (pair) =>
        Object.isFrozen(pair) && Object.isFrozen(pair.beforePath) && Object.isFrozen(pair.afterPath)
    )
  ).toBe(true);
});

test('planning preserves the live document, payloads, selection, revision, and existing history', async () => {
  const value = await fixture(),
    editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(value.pkg),
    });
  try {
    const surface = editor.surface!,
      session = surface.session,
      id = session.paragraphIds().at(-1)!;
    surface.applyAutomationOps(() => [
      { op: 'insertText', paragraphId: id, offset: 0, text: 'User ' },
    ]);
    const before = writeOoxmlPackage(session.currentPackage()),
      selection = surface.state().selection,
      history = [session.canUndo(), session.canRedo()],
      current = session.packageRevision();
    const result = planDomainContainers(
      {
        ...value.source,
        pkg: session.currentPackage(),
        revision: { ...revision, revision: current },
      },
      parsed(value.markdown.replace('Old question', 'Edited question'))
    );
    expect(result.success).toBe(true);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
    expect(surface.state().selection).toEqual(selection);
    expect(session.packageRevision()).toBe(current);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
  } finally {
    editor.destroy();
  }
});

for (const [attributes, media, trust] of [
  ['type="generated" url="https://youtu.be/video"', 'youtube', 'generated'],
  ['type="verified" url="https://www.youtube.com/watch?v=id"', 'youtube', 'verified'],
  ['type="web" url="https://youtu.be/video"', 'web', 'verified'],
  ['type="youtube" url="https://example.test/page"', 'youtube', 'verified'],
  ['title="Book"', 'web', undefined],
] as const)
  test(`source media follows the captured application grammar independently of trust: ${attributes}`, () => {
    const root = parsed(`<source ${attributes}>\n\nQuote.\n\n</source>`)[0]!;
    expect(domainNodeFromAst(root)).toMatchObject({ type: 'source_excerpt', sourceType: media });
    expect(domainNodeFromAst(root)?.trust).toBe(trust);
  });
