import { expect, test } from 'bun:test';
import { createEditPlan, decodeEditPlan, serializeEditPlan, type CanonicalEdit } from './plan';
import type { InsertedListBinding } from './inserted-list-types';

const revision = { documentId: 'inserted-lists', epoch: 'load-1', revision: 1 };
const fragment = [
  { type: 'paragraph', children: [{ type: 'text', value: 'New' }] },
  {
    type: 'blockquote',
    children: [{ type: 'heading', depth: 2, children: [{ type: 'text', value: 'Nested' }] }],
  },
];
const edit = (lists: readonly InsertedListBinding[]): CanonicalEdit => ({
  kind: 'siblings',
  partName: '/word/document.xml',
  parentId: 'body',
  beforeId: 'before',
  afterId: 'after',
  removedIds: [],
  inserted: fragment,
  lists,
});

test('list bindings retain existing numbering, new instance keys, checklist state, and plain paragraph targets', () => {
  for (const target of [
    { kind: 'none' },
    { kind: 'ordered', level: 0, instance: { kind: 'existing', numId: '7' } },
    { kind: 'bullet', level: 2, instance: { kind: 'new', key: 'new-group' } },
    { kind: 'ordered', level: 1, instance: { kind: 'new', key: 'new-group', start: 0 } },
    { kind: 'checklist', level: 0, checked: false },
  ] as const) {
    const result = createEditPlan(revision, [
      edit([
        { path: [0], target },
        { path: [1, 0], target },
      ]),
    ]);
    expect(result.success).toBe(true);
    if (result.success) {
      expect(decodeEditPlan(serializeEditPlan(result.plan))).toEqual(result);
      expect(Object.isFrozen(result.plan.operations[0])).toBe(true);
      expect(result.plan.operations[0]).toMatchObject({
        lists: [
          { path: [0], target },
          { path: [1, 0], target },
        ],
      });
    }
  }
});

test('invalid fragment paths, duplicate bindings, native IDs, limits, and extra fields refuse before a plan returns', () => {
  const base = { path: [0], target: { kind: 'none' } };
  const invalid = [
    [{ ...base, path: [] }],
    [{ ...base, path: [-1] }],
    [{ ...base, path: [0.5] }],
    [{ ...base, path: ['0'] }],
    [{ ...base, path: [2] }],
    [{ ...base, path: [0, 0] }],
    [base, base],
    [{ ...base, paragraphId: 'preview-id' }],
    [
      {
        ...base,
        target: { kind: 'ordered', level: 9, instance: { kind: 'existing', numId: '1' } },
      },
    ],
    [
      {
        ...base,
        target: {
          kind: 'ordered',
          level: 0,
          instance: { kind: 'new', key: 'g', start: 2147483648 },
        },
      },
    ],
    [{ ...base, target: { kind: 'bullet', level: 0, instance: { kind: 'existing', numId: '0' } } }],
    [{ ...base, target: { kind: 'checklist', level: 0, checked: 'yes' } }],
  ];
  for (const lists of invalid) {
    const result = createEditPlan(revision, [edit(lists as unknown as InsertedListBinding[])]);
    expect(result.success).toBe(false);
    expect(Object.hasOwn(result, 'plan')).toBe(false);
  }
});

test('new list instance definitions remain consistent across inserted fragments and retained paragraph edits', () => {
  const ordered = {
    kind: 'ordered',
    level: 0,
    instance: { kind: 'new', key: 'one', start: 7 },
  } as const;
  const before = { numbering: null, checkbox: null, fingerprint: 'sha256:' + 'a'.repeat(64) };
  const first = edit([{ path: [0], target: ordered }]);
  for (const after of [
    ordered,
    { ...ordered, instance: { ...ordered.instance, start: 8 } },
    { kind: 'bullet', level: 0, instance: { kind: 'new', key: 'one' } },
  ] as const) {
    const result = createEditPlan(revision, [
      first,
      { kind: 'list', partName: '/word/document.xml', paragraphId: 'retained', before, after },
    ]);
    expect(result.success).toBe(after === ordered);
  }
});

test('binding outputs detach while malformed getters, sparse paths, and absent paragraphs cannot lose intentions', () => {
  const input = [{ path: [0], target: { kind: 'checklist' as const, level: 0, checked: false } }],
    before = JSON.stringify(input),
    result = createEditPlan(revision, [edit(input)]);
  expect(result.success).toBe(true);
  expect(JSON.stringify(input)).toBe(before);
  input[0]!.target.checked = true;
  if (result.success)
    expect(result.plan.operations[0]).toMatchObject({ lists: [{ target: { checked: false } }] });
  let calls = 0;
  const getter = Object.defineProperty({ path: [0] }, 'target', {
    enumerable: true,
    get() {
      calls++;
      return { kind: 'none' };
    },
  });
  expect(createEditPlan(revision, [edit([getter as unknown as InsertedListBinding])]).success).toBe(
    false
  );
  expect(
    createEditPlan(revision, [edit([{ path: new Array(1), target: { kind: 'none' } }])]).success
  ).toBe(false);
  expect(
    createEditPlan(revision, [
      { ...edit([{ path: [0], target: { kind: 'none' } }]), inserted: [] } as CanonicalEdit,
    ]).success
  ).toBe(false);
  expect(calls).toBe(0);
});

test('one new multilevel instance permits independent level formats and rejects conflicts at the same level', () => {
  const root = { kind: 'bullet', level: 0, instance: { kind: 'new', key: 'mixed' } } as const,
    child = {
      kind: 'ordered',
      level: 1,
      instance: { kind: 'new', key: 'mixed', start: 7 },
    } as const;
  expect(
    createEditPlan(revision, [
      edit([
        { path: [0], target: root },
        { path: [1, 0], target: child },
      ]),
    ]).success
  ).toBe(true);
  expect(
    createEditPlan(revision, [
      edit([
        { path: [0], target: root },
        { path: [1, 0], target: { ...child, level: 0 } },
      ]),
    ]).success
  ).toBe(false);
  const implicit = { ...child, instance: { kind: 'new', key: 'mixed' } } as const,
    explicit = { ...child, instance: { kind: 'new', key: 'mixed', start: 1 } } as const;
  expect(
    createEditPlan(revision, [
      edit([
        { path: [0], target: implicit },
        { path: [1, 0], target: explicit },
      ]),
    ]).success
  ).toBe(true);
});
