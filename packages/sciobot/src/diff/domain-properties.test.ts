import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import {
  applyTreeOp,
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  escapeXml,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import oracle from '../__tests__/fixtures/domain-properties.json';
import pairingOracle from '../__tests__/fixtures/domain-pairing.json';
import { docxFixture } from '../../test/support/document';
import { buildBlockIndex } from '../blocks/block-index';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import { DEFAULT_PAYLOAD_LIMITS, type JsonObject } from '../blocks/payload-schema';
import { GENERATED_LABEL_TAG } from '../projection/block-tags';
import { projectMdx } from '../projection/mdx';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { planDomainContainers } from './domain-containers';
import { planDomainProperties, type DomainPropertyTarget } from './domain-properties';
import { directParagraphTextOps, planParagraphText } from './paragraph-text';
import {
  decodeEditPlan,
  serializeEditPlan,
  type CanonicalEditPlan,
  type EditPlanResult,
} from './plan';

const revision = { documentId: 'properties', epoch: 'load-1', revision: 3 };
const p = (text: string) => `<w:p><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
function fixture(
  type = 'lesson_activity',
  attributes: JsonObject = { name: 'Old', duration: 10, resource: 'A', cache: { audio: 'keep' } }
) {
  const body =
    `<w:sdt><w:sdtPr><w:tag w:val="${blockTag('block')}"/></w:sdtPr><w:sdtContent>` +
    `<w:sdt><w:sdtPr><w:tag w:val="${GENERATED_LABEL_TAG}"/></w:sdtPr><w:sdtContent>${p('Label')}</w:sdtContent></w:sdt>` +
    p('Old sentence.') +
    p('Keep sentence.') +
    '</w:sdtContent></w:sdt>' +
    p('Outside.');
  const loaded = readOoxmlPackage(docxFixture(body));
  if (!loaded.ok) throw new Error(loaded.reason);
  const stored = withPayloadRecord(
    loaded.package,
    {
      id: 'block',
      type,
      version: 1,
      attributes,
      extension: { nativeAsset: 'pristine-1' },
    },
    'Label'
  );
  if (!stored.ok) throw new Error(stored.reason);
  const pkg = stored.pkg,
    part = pkg.parts.get(pkg.mainDocumentPart)!,
    block = buildBlockIndex(pkg).byId.get('block')!,
    content = contentControlContentNodeOf(findNode(part, block.controlId)!)!;
  return { pkg, part, block, content, source: { pkg, revision, blockId: block.id } };
}
function success(result: EditPlanResult): CanonicalEditPlan {
  if (!result.success) throw new Error(result.error);
  return result.plan;
}
/** Immutable native codec oracle only. Production execution remains A01/R03. */
function applyMetadata(pkg: OoxmlPackage, plan: CanonicalEditPlan, side: 'before' | 'after') {
  const block = buildBlockIndex(pkg).byId.get('block')!;
  let attrs = block.payload.attributes;
  for (const edit of plan.operations) {
    if (edit.kind !== 'payload' || edit.path.length !== 2 || edit.path[0] !== 'attributes')
      throw new Error('Unexpected property oracle edit');
    const name = edit.path[1] as string,
      state = edit[side];
    attrs = Object.fromEntries([
      ...Object.entries(attrs).filter(([key]) => key !== name),
      ...(state.present ? [[name, state.value] as const] : []),
    ]);
  }
  const stored = withPayloadRecord(pkg, { ...block.payload, attributes: attrs }, 'Label');
  if (!stored.ok) throw new Error(stored.reason);
  return stored.pkg;
}
function attrs(pkg: OoxmlPackage) {
  return buildBlockIndex(pkg).byId.get('block')!.payload.attributes;
}

test('renaming an activity produces one exact native payload delta without a wrapper or text edit', () => {
  const value = fixture(),
    plan = success(planDomainProperties(value.source, { set: { name: 'New' } }));
  expect(plan.operations).toEqual([
    {
      kind: 'payload',
      partName: value.part.name,
      blockId: 'block',
      controlId: value.block.controlId,
      path: ['attributes', 'name'],
      before: { present: true, value: 'Old' },
      after: { present: true, value: 'New' },
    },
  ]);
  const proposed = applyMetadata(value.pkg, plan, 'after');
  expect(proposed.parts.get(value.part.name) === value.part).toBe(true);
  expect(attrs(proposed)).toEqual({ ...value.block.payload.attributes, name: 'New' });
  expect(buildBlockIndex(proposed).byId.get('block')!.payload.extension).toEqual({
    nativeAsset: 'pristine-1',
  });
  expect(canonicalOoxmlFingerprint(proposed.parts.get(value.part.name)!)).toBe(
    canonicalOoxmlFingerprint(value.part)
  );
});

test.each(oracle.cases)('property presence matches executed application views: $name', (item) => {
  const before = item.before as JsonObject,
    after = item.after as JsonObject;
  const value = fixture('lesson_activity', before);
  const target: DomainPropertyTarget = {
    set: Object.fromEntries(
      item.keys.filter((key) => Object.hasOwn(after, key)).map((key) => [key, after[key]!])
    ),
    remove: item.keys.filter((key) => !Object.hasOwn(after, key)),
  };
  const plan = success(planDomainProperties(value.source, target));
  expect(plan.operations).toHaveLength(item.keys.length);
  for (const edit of plan.operations) {
    if (edit.kind !== 'payload') throw new Error('Expected only property deltas');
    const key = edit.path[1] as string;
    expect(edit.before).toEqual(
      Object.hasOwn(before, key) ? { present: true, value: before[key] } : { present: false }
    );
    expect(edit.after).toEqual(
      Object.hasOwn(after, key) ? { present: true, value: after[key] } : { present: false }
    );
  }
  const proposed = applyMetadata(value.pkg, plan, 'after');
  expect(attrs(proposed)).toEqual(item.working as JsonObject);
  expect(attrs(applyMetadata(proposed, plan, 'before'))).toEqual(item.committed as JsonObject);
});

test('captured activity updates retain both declared update properties and the requested resource', () => {
  const captured = pairingOracle.cases.find((item) => item.name === 'activity name and duration')!;
  const value = fixture(),
    plan = success(planDomainProperties(value.source, { set: { name: 'Opening', duration: 15 } }));
  const states = Object.fromEntries(
    plan.operations.map((edit) => {
      if (edit.kind !== 'payload' || !edit.after.present) throw new Error('Missing property');
      return [edit.path[1], edit.after.value];
    })
  );
  expect(states).toEqual(captured.update?.next as JsonObject);
  const resource = success(
    planDomainProperties(value.source, { set: { duration: 15, resource: 'B' } })
  );
  expect(resource.operations.map((edit) => edit.kind === 'payload' && edit.path[1])).toEqual([
    'duration',
    'resource',
  ]);
  expect(attrs(applyMetadata(value.pkg, resource, 'after'))).toEqual({
    ...value.block.payload.attributes,
    duration: 15,
    resource: 'B',
  });
  expect(planDomainProperties(value.source, { set: { resource: 'B' } })).toMatchObject({
    success: false,
    code: 'unsupported-operation',
  });
});

test('duration and one sentence compose as separate plans under the retained activity', async () => {
  const value = fixture(),
    opened = openDocumentForExport(writeOoxmlPackage(value.pkg));
  if (!opened.ok) throw new Error(opened.reason);
  let markdown: string;
  try {
    const result = projectMdx(value.pkg, await opened.session.layoutFor('proposed'), revision);
    if (!result.success) throw new Error(result.error);
    markdown = result.projection.text;
  } finally {
    opened.session.dispose();
  }
  const parse = (text: string): readonly TargetAstNode[] => {
    const result = parseTargetMarkdown(text);
    if (!result.success) throw new Error(result.error);
    return result.target.normalizedAst.children;
  };
  const before = parse(markdown),
    after = parse(
      markdown.replace('duration="10"', 'duration="15"').replace('Old sentence', 'New sentence')
    );
  const root = bodyStoryRoot(value.part)!;
  if (root.kind === 'textValue') throw new Error('Missing native body children');
  const paired = planDomainContainers(
    {
      ...value.source,
      partName: value.part.name,
      parentId: root.id,
      startIndex: 0,
      blocks: [
        {
          id: value.block.controlId,
          ast: before[0]!,
          children: {
            parentId: value.content.id,
            startIndex: 1,
            blocks: (before[0]!.children as TargetAstNode[]).map((ast, at) => ({
              id: value.content.children[at + 1]!.id,
              ast,
            })),
          },
        },
        { id: root.children[1]!.id, ast: before[1]! },
      ],
    },
    after
  );
  if (!paired.success) throw new Error(paired.error);
  expect(paired.plan.operations).toHaveLength(0);
  expect(paired.pairs.filter((pair) => pair.kind === 'container').map((pair) => pair.id)).toEqual([
    value.block.controlId,
  ]);
  const props = success(planDomainProperties(value.source, { set: { duration: 15 } }));
  const paragraph = value.content.children[1]!;
  const text = success(
    planParagraphText(
      {
        revision,
        partName: value.part.name,
        paragraphId: paragraph.id,
        text: paragraphTextOf(value.part, paragraph.id)!,
      },
      'New sentence.'
    )
  );
  let part = value.part;
  for (const edit of text.operations) {
    if (edit.kind !== 'text') throw new Error('Unexpected sentence edit');
    for (const op of directParagraphTextOps(edit)) {
      const changed = applyTreeOp(part, op);
      if (!changed.ok) throw new Error(changed.reason);
      part = changed.part;
    }
  }
  const proposed = applyMetadata(withPart(value.pkg, part), props, 'after');
  const reopened = readOoxmlPackage(writeOoxmlPackage(proposed));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(attrs(reopened.package).duration).toBe(15);
  expect(paragraphTextOf(part, paragraph.id)).toBe('New sentence.');
  expect(findNode(part, value.content.children[2]!.id) === value.content.children[2]).toBe(true);
  expect(contentControlContentNodeOf(findNode(part, value.block.controlId)!)!.id).toBe(
    value.content.id
  );
  expect(props.operations).toHaveLength(1);
  expect(text.operations.every((edit) => edit.kind === 'text')).toBe(true);
});

test('title minutes and phase duration use canonical metadata names while identity changes stay structural', () => {
  const title = fixture('lesson_title', {
    grade: '7',
    organizationalMinutes: 5,
    lessonMinutes: 40,
  });
  const plan = success(
    planDomainProperties(title.source, {
      set: { organizationalMinutes: 10 },
      remove: ['lessonMinutes'],
    })
  );
  expect(plan.operations.map((edit) => edit.kind === 'payload' && edit.path)).toEqual([
    ['attributes', 'organizationalMinutes'],
    ['attributes', 'lessonMinutes'],
  ]);
  expect(planDomainProperties(title.source, { set: { grade: '8' } }).success).toBe(false);
  const phase = fixture('lesson_phase', { phaseType: 'evocation', duration: 10 });
  expect(
    success(planDomainProperties(phase.source, { set: { duration: 15 } })).operations
  ).toHaveLength(1);
  expect(planDomainProperties(phase.source, { set: { phaseType: 'reflection' } }).success).toBe(
    false
  );
  const recap = fixture('recap_phase', { phaseType: 'evocation', duration: 10 });
  expect(planDomainProperties(recap.source, { set: { duration: 15 } }).success).toBe(false);
  const source = fixture('source_excerpt', {
    url: 'https://example.com',
    trust: 'verified',
    note: 'old',
  });
  expect(
    success(planDomainProperties(source.source, { set: { note: 'new' } })).operations
  ).toHaveLength(1);
  expect(
    planDomainProperties(source.source, { set: { url: 'https://other.example.com' } }).success
  ).toBe(false);
});

test('semantic no-ops and absent removals preserve hidden payload values and produce no work', () => {
  const value = fixture();
  expect(success(planDomainProperties(value.source, {})).operations).toEqual([]);
  expect(
    success(planDomainProperties(value.source, { set: { name: 'Old', duration: 10 } })).operations
  ).toEqual([]);
  const missing = fixture('lesson_activity', { name: 'Old' });
  expect(
    success(planDomainProperties(missing.source, { remove: ['duration'] })).operations
  ).toEqual([]);
  const title = fixture('lesson_title', { grade: '7', extra: { a: 1, b: [2] } });
  expect(
    success(planDomainProperties(title.source, { set: { extra: { b: [2], a: 1 } } })).operations
  ).toEqual([]);
});

test('data paths and nested values stay detached and prototype-shaped keys remain inert', () => {
  const value = fixture('lesson_title', { grade: '7' });
  const set = JSON.parse(
    '{"constructor":{"prototype":"data","__proto__":{"polluted":true}},"extra":{"nested":[1]}}'
  ) as JsonObject;
  const plan = success(planDomainProperties(value.source, { set }));
  expect(plan.operations.map((edit) => edit.kind === 'payload' && edit.path[1])).toEqual([
    'constructor',
    'extra',
  ]);
  (set.extra as { nested: number[] }).nested.push(2);
  expect(Object.isFrozen(set.extra)).toBe(false);
  expect(JSON.stringify(plan)).not.toContain('[1,2]');
  expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  const decoded = decodeEditPlan(serializeEditPlan(plan));
  expect(decoded).toEqual({ success: true, plan });
  expect(Object.isFrozen(plan.operations[0])).toBe(true);
  expect(
    Object.entries(attrs(applyMetadata(value.pkg, plan, 'after'))).find(
      ([key]) => key === 'constructor'
    )?.[1]
  ).toEqual(JSON.parse('{"prototype":"data","__proto__":{"polluted":true}}') as JsonObject);
  expect(
    planDomainProperties(value.source, { set: JSON.parse('{"__proto__":{"polluted":true}}') })
      .success
  ).toBe(false);
});

test('invalid and unsupported metadata patches refuse atomically before getters or native writes', () => {
  const value = fixture(),
    bytes = writeOoxmlPackage(value.pkg);
  let calls = 0;
  const getter = Object.defineProperty({}, 'name', {
    enumerable: true,
    get() {
      calls++;
      return 'Bad';
    },
  });
  const cycle: Record<string, unknown> = {};
  cycle.self = cycle;
  let deep: unknown = 'leaf';
  for (let at = 0; at < 70; at++) deep = { child: deep };
  const cases: unknown[] = [
    null,
    [],
    { unknown: true },
    { set: [] },
    { remove: 'name' },
    { set: { name: 'New' }, remove: ['name'] },
    { remove: ['name', 'name'] },
    { remove: [1] },
    { set: getter },
    { set: { name: undefined } },
    { set: { duration: Infinity } },
    { set: cycle },
    { set: { extra: deep } },
    ...[
      'type',
      'id',
      'children',
      'listStart',
      'suggestion',
      '_private',
      'suggestion_update',
      'cache',
    ].map((key) => ({ set: { [key]: 'Bad' } })),
  ];
  for (const target of cases)
    expect(planDomainProperties(value.source, target as DomainPropertyTarget).success).toBe(false);
  expect(
    planDomainProperties(value.source, {
      set: { name: 'x'.repeat(DEFAULT_PAYLOAD_LIMITS.maxBytes) },
    })
  ).toMatchObject({ success: false, code: 'invalid-input', error: 'byte-limit' });
  expect(calls).toBe(0);
  expect(
    planDomainProperties({ ...value.source, blockId: 'missing' }, { set: { name: 'New' } })
  ).toMatchObject({ success: false, code: 'invalid-document' });
  const forged = Object.defineProperty({ ...value.source }, 'blockId', {
    enumerable: true,
    get() {
      calls++;
      return 'block';
    },
  });
  expect(planDomainProperties(forged, { set: { name: 'New' } }).success).toBe(false);
  expect(calls).toBe(0);
  for (const type of ['lesson_visual', 'source', 'listening', 'unknown'])
    expect(planDomainProperties(fixture(type).source, { set: { name: 'New' } })).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
  expect(writeOoxmlPackage(value.pkg)).toEqual(bytes);
});

test('planning preserves live selection, revision, package bytes, and existing undo/redo', () => {
  const value = fixture(),
    editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(value.pkg),
    });
  try {
    const surface = editor.surface!,
      session = surface.session,
      paragraphId = session.paragraphIds().at(-1)!;
    surface.applyAutomationOps(() => [{ op: 'insertText', paragraphId, offset: 0, text: 'User ' }]);
    const bytes = writeOoxmlPackage(session.currentPackage()),
      current = session.packageRevision(),
      selection = surface.state().selection,
      history = [session.canUndo(), session.canRedo()];
    const result = planDomainProperties(
      {
        ...value.source,
        pkg: session.currentPackage(),
        revision: { ...revision, revision: current },
      },
      { set: { name: 'New' } }
    );
    expect(result.success).toBe(true);
    expect(session.packageRevision()).toBe(current);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(bytes);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
  } finally {
    editor.destroy();
  }
});
