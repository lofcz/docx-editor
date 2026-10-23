import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import type { JsonValue } from '../blocks/payload-schema';
import { parseTargetMarkdown } from '../projection/parse-target';
import {
  createEditPlan,
  decodeEditPlan,
  serializeEditPlan,
  EDIT_PLAN_LIMITS,
  type CanonicalEdit,
  type CanonicalEditPlan,
  type PlannedValue,
} from './plan';

const revision = { documentId: 'lesson', epoch: 'load-1', revision: 7 };
const partName = '/word/document.xml';
const absent = { present: false } as const;
const present = (value: JsonValue): PlannedValue => ({ present: true, value });
const text = {
  kind: 'text',
  partName,
  paragraphId: 'native-paragraph',
  start: 3,
  end: 6,
  before: 'old',
  after: 'new',
} as const;
const payload = {
  kind: 'payload',
  partName,
  blockId: 'activity',
  controlId: 'native-control',
  path: ['attributes', 'duration'],
  before: present('10'),
  after: present('15'),
} as const;
function plan(operations: readonly CanonicalEdit[]): CanonicalEditPlan {
  const result = createEditPlan(revision, operations);
  if (!result.success) throw new Error(result.error);
  return result.plan;
}
function fragment() {
  const result = parseTargetMarkdown('## Inserted\n\nA new paragraph.');
  if (!result.success) throw new Error(result.error);
  return result.target.normalizedAst.children;
}

test('empty diffs and explicit no-op edits emit no operations', () => {
  expect(plan([])).toEqual({ schemaVersion: 1, revision, operations: [] });
  expect(
    plan([
      { ...text, after: 'old' },
      {
        kind: 'format',
        partName,
        target: { kind: 'paragraph', paragraphId: 'p' },
        changes: [
          {
            name: 'spacing',
            before: present({ after: 2, before: 1 }),
            after: present({ before: 1, after: 2 }),
          },
        ],
      },
      { ...payload, after: payload.before },
      {
        kind: 'list',
        partName,
        paragraphId: 'p',
        before: { numbering: null, checkbox: null, fingerprint: 'sha256:' + 'a'.repeat(64) },
        after: { kind: 'none' },
      },
      {
        kind: 'siblings',
        partName,
        parentId: 'body',
        beforeId: 'previous',
        afterId: 'next',
        removedIds: [],
        inserted: [],
      },
    ]).operations
  ).toEqual([]);
});

test('mode-independent plans roundtrip text, format, sibling, table, list, and payload addresses', () => {
  const operations: CanonicalEdit[] = [
    text,
    {
      kind: 'format',
      partName,
      target: { kind: 'run', paragraphId: 'native-paragraph', start: 3, end: 6 },
      changes: [{ name: 'b', before: absent, after: present({}) }],
    },
    {
      kind: 'siblings',
      partName,
      parentId: 'native-body',
      beforeId: 'previous',
      afterId: 'next',
      removedIds: ['removed'],
      inserted: fragment(),
    },
    {
      kind: 'table',
      partName,
      operation: {
        op: 'insertTableRow',
        tableId: 'native-table',
        rowId: 'native-row',
        where: 'below',
      },
    },
    payload,
    {
      kind: 'list',
      partName,
      paragraphId: 'native-list-item',
      before: {
        numbering: { numId: '1', level: 0 },
        checkbox: null,
        fingerprint: 'sha256:' + 'a'.repeat(64),
      },
      after: { kind: 'ordered', level: 1, instance: { kind: 'existing', numId: '1' } },
    },
    {
      kind: 'table',
      partName,
      operation: {
        op: 'authorTable',
        action: {
          kind: 'existing',
          tableId: 'native-table',
          mutation: {
            kind: 'insertRows',
            rowId: 'native-row',
            location: 'after',
            count: 1,
            values: [['One', 'Two']],
          },
        },
      },
    },
  ];
  const result = plan(operations);
  expect(result.operations).toEqual(operations);
  const serialized = serializeEditPlan(result);
  expect(decodeEditPlan(serialized)).toEqual({ success: true, plan: result });
  expect(JSON.parse(serialized)).toEqual({ schemaVersion: 1, revision, operations });
  expect(Object.keys(JSON.parse(serialized))).toEqual(['schemaVersion', 'revision', 'operations']);
  expect(serialized).not.toContain('storyRootId');
  expect(serialized).not.toContain('replaceStoryBlocks');
  expect(serialized).not.toContain('"mode"');
});

test('plan data is detached and deeply frozen without freezing caller data', () => {
  const source = { ...text, after: 'first' };
  const sourceValue = { minutes: 15 };
  const result = plan([source, { ...payload, before: absent, after: present(sourceValue) }]);
  source.after = 'later';
  sourceValue.minutes = 20;
  expect(result.operations[0]).toMatchObject({ after: 'first' });
  expect(Object.isFrozen(source)).toBe(false);
  expect(Object.isFrozen(result)).toBe(true);
  expect(Object.isFrozen(result.revision)).toBe(true);
  expect(Object.isFrozen(result.operations)).toBe(true);
  expect(Object.isFrozen(result.operations[0])).toBe(true);
  expect(result.revision === revision).toBe(false);
  expect(result.operations[1]).toMatchObject({ after: { present: true, value: { minutes: 15 } } });
  expect(Object.isFrozen(sourceValue)).toBe(false);
});

test('missing properties and explicit null remain separate review states', () => {
  const changes: CanonicalEdit[] = [
    { ...payload, before: absent, after: present(null) },
    { ...payload, before: present(null), after: absent },
  ];
  const result = plan(changes);
  expect(result.operations).toEqual(changes);
  expect(decodeEditPlan(serializeEditPlan(result))).toEqual({ success: true, plan: result });
});

test('format plans omit unchanged properties and retain changed old-value presence', () => {
  const result = plan([
    {
      kind: 'format',
      partName,
      target: { kind: 'hyperlink', hyperlinkId: 'link-1' },
      changes: [
        {
          name: 'url',
          before: present('https://example.com'),
          after: present('https://example.org'),
        },
        { name: 'title', before: absent, after: absent },
      ],
    },
  ]);
  expect(result.operations).toEqual([
    {
      kind: 'format',
      partName,
      target: { kind: 'hyperlink', hyperlinkId: 'link-1' },
      changes: [
        {
          name: 'url',
          before: present('https://example.com'),
          after: present('https://example.org'),
        },
      ],
    },
  ]);
});

test('prototype-shaped payload keys remain inert data', () => {
  const value = JSON.parse(
    '{"__proto__":{"polluted":true},"constructor":{"prototype":{"polluted":true}}}'
  );
  const result = plan([{ ...payload, before: absent, after: present(value) }]);
  expect(result.operations[0]).toMatchObject({ after: { present: true, value } });
  expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  expect(decodeEditPlan(serializeEditPlan(result))).toEqual({ success: true, plan: result });
});

test('getters, toJSON, cycles, and non-JSON values fail before code executes', () => {
  let calls = 0;
  const getter = { ...text };
  Object.defineProperty(getter, 'after', {
    enumerable: true,
    get() {
      calls += 1;
      return 'bad';
    },
  });
  const customized = {
    ...text,
    toJSON() {
      calls += 1;
      return text;
    },
  };
  const cyclic: Record<string, unknown> = { ...text };
  cyclic.self = cyclic;
  for (const value of [
    getter,
    customized,
    cyclic,
    { ...text, after: undefined },
    { ...text, after: Infinity },
    { ...text, after: new Date() },
  ]) {
    const result = createEditPlan(revision, [value as CanonicalEdit]);
    expect(result).toMatchObject({ success: false, code: 'invalid-input' });
    expect(Object.hasOwn(result, 'plan')).toBe(false);
  }
  expect(calls).toBe(0);
});

test('live editors and canonical trees cannot become inserted sibling fragments', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture('<w:p><w:r><w:t>Keep</w:t></w:r></w:p>'),
  });
  try {
    const session = editor.surface!.session;
    const before = canonicalOoxmlFingerprint(session.part());
    for (const inserted of [
      [session.part().root],
      [{ type: 'root', children: fragment() }],
      [{ type: 'paragraph', kind: 'paragraph', children: [] }],
      [{ type: 'paragraph', children: [], editor }],
    ]) {
      const result = createEditPlan(revision, [
        {
          kind: 'siblings',
          partName,
          parentId: 'body',
          beforeId: null,
          afterId: null,
          removedIds: [],
          inserted,
        } as unknown as CanonicalEdit,
      ]);
      expect(result).toMatchObject({ success: false, code: 'invalid-input' });
      expect(Object.hasOwn(result, 'plan')).toBe(false);
    }
    expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
    expect(session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

test('invalid addresses and structural anchors fail without partial plans', () => {
  const invalid = [
    { ...text, start: -1 },
    { ...text, end: 5 },
    { ...text, bias: 'unknown' },
    { ...text, insideHyperlink: '' },
    { ...text, insideHyperlink: 'link', after: 'line\nbreak' },
    { ...text, insideHyperlink: 'link', after: '' },
    { ...text, insideHyperlink: 'link', before: 'x', start: 0, end: 1 },
    { ...text, partName: '/word/../document.xml' },
    { ...payload, path: [-1] },
    {
      kind: 'table',
      partName,
      operation: {
        op: 'insertTableRow',
        tableId: 'table',
        rowId: 'row',
        where: 'below',
        document: {},
      },
    },
    { ...payload, before: { present: false, value: null } },
    {
      kind: 'format',
      partName,
      target: { kind: 'paragraph', paragraphId: 'p' },
      changes: [
        { name: 'b', before: absent, after: present({}) },
        { name: 'b', before: absent, after: present({}) },
      ],
    },
    {
      kind: 'siblings',
      partName,
      parentId: 'body',
      beforeId: 'removed',
      afterId: 'next',
      removedIds: ['removed'],
      inserted: [],
    },
    {
      kind: 'siblings',
      partName,
      parentId: 'body',
      beforeId: null,
      afterId: null,
      removedIds: ['duplicate', 'duplicate'],
      inserted: [],
    },
    {
      kind: 'table',
      partName,
      operation: { op: 'replaceStoryBlocks', storyRootId: 'body', paragraphs: ['Destroyed'] },
    },
    {
      kind: 'table',
      partName,
      operation: {
        op: 'insertTableRow',
        tableId: 'table',
        rowId: 'row',
        where: 'below',
        revision: { author: 'AI' },
      },
    },
  ];
  for (const value of invalid) {
    const result = createEditPlan(revision, [text, value as CanonicalEdit]);
    expect(result).toMatchObject({ success: false, code: 'invalid-input' });
    expect(Object.hasOwn(result, 'plan')).toBe(false);
  }
});

test('serialized plans reject malformed JSON, unknown versions, and hidden root state', () => {
  const valid = plan([text]);
  for (const raw of [
    '{',
    'null',
    '[]',
    JSON.stringify({ ...valid, schemaVersion: 2 }),
    JSON.stringify({ ...valid, document: {} }),
    JSON.stringify({ ...valid, revision: { ...revision, revision: -1 } }),
  ])
    expect(decodeEditPlan(raw)).toMatchObject({ success: false, code: 'invalid-input' });
  expect(() => serializeEditPlan({ ...valid, document: {} } as CanonicalEditPlan)).toThrow(
    'invalid-plan-schema'
  );
});

test('resource limits reject excessive bytes, nodes, and depth', () => {
  let deep: JsonValue = null;
  for (let index = 0; index < EDIT_PLAN_LIMITS.maxDepth + 1; index += 1) deep = { nested: deep };
  const attempts: [CanonicalEdit, string][] = [
    [{ ...text, after: 'a'.repeat(EDIT_PLAN_LIMITS.maxBytes) }, 'byte-limit'],
    [{ ...payload, after: present(Array(EDIT_PLAN_LIMITS.maxNodes + 1).fill(1)) }, 'node-limit'],
    [{ ...payload, after: present(deep) }, 'depth-limit'],
  ];
  for (const [edit, error] of attempts)
    expect(createEditPlan(revision, [edit])).toMatchObject({
      success: false,
      code: 'invalid-input',
      error,
    });
});
