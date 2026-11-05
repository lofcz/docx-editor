import { expect, test } from 'bun:test';
import { blankDocumentBytes } from '@docx-editor.dev/core';
import {
  customXmlNodes,
  readOoxmlPackage,
  replaceChildren,
  withPart,
  writeOoxmlPackage,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import type { BlockPayload, JsonObject } from './payload-schema';
import { withPreparedPayloadFields } from './payload-preparation';
import { readFields } from './payload-fields';

const CREATED: BlockPayload = {
  id: 'block-1',
  type: 'lesson_title',
  version: 1,
  attributes: { grade: '5', lessonMinutes: 45, nested: { a: [1, 2] } },
  mdast: { type: 'root', children: [] },
};

function created(): OoxmlPackage {
  const opened = readOoxmlPackage(blankDocumentBytes());
  if (!opened.ok) throw new Error(opened.reason);
  const written = withPayloadRecord(opened.package, CREATED, 'Title');
  if (!written.ok) throw new Error(written.reason);
  return written.pkg;
}

function store(pkg: OoxmlPackage) {
  const read = readPayloadStore(pkg);
  if (read.status !== 'ready')
    throw new Error(read.status === 'invalid' ? read.reason : read.status);
  return read;
}

function current(pkg: OoxmlPackage, id = 'block-1') {
  const entry = store(pkg).entries.get(id)!.decoded;
  if (entry.status !== 'supported') throw new Error(entry.status);
  return entry.value;
}

function record(pkg: OoxmlPackage, id = 'block-1'): OoxmlElement {
  const part = pkg.parts.get(store(pkg).part.partName)!;
  return part.root.children.find(
    (child) =>
      child.kind !== 'textValue' &&
      child.attributes.some((a) => a.localName === 'id' && a.value === `scio_${id}`)
  ) as OoxmlElement;
}

function update(pkg: OoxmlPackage, payload: BlockPayload, label = 'Title') {
  const written = withPayloadRecord(pkg, payload, label);
  if (!written.ok) throw new Error(written.reason);
  return written.pkg;
}

function withAuthoredChildren(pkg: OoxmlPackage, extra: string): OoxmlPackage {
  const node = record(pkg);
  const ns = node.namespaceUri;
  const parsed = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!parsed.ok) throw new Error(parsed.reason);
  const reread = record(parsed.package);
  const part = parsed.package.parts.get(store(parsed.package).part.partName)!;
  const doc = new DOMParser().parseFromString(`<x xmlns="${ns}">${extra}</x>`, 'application/xml');
  const authored = [...doc.documentElement.children].map((element, index) => ({
    id: `${reread.id}-authored-${index}`,
    kind: 'generic',
    namespaceUri: ns,
    localName: element.localName,
    namespaceBindings: [],
    attributes: [...element.attributes].map((attribute) => ({
      kind: 'genericExtension',
      namespaceUri: '',
      localName: attribute.localName,
      value: attribute.value,
    })),
    children: [],
  })) as unknown as OoxmlElement[];
  const replaced = replaceChildren(part, reread.id, [...reread.children, ...authored]);
  if (!replaced.ok) throw new Error('replace refused');
  return withPart(parsed.package, replaced.part);
}

test('an update leaves the created data text alone and records changed fields only', () => {
  const pkg = created();
  const before = customXmlNodes(pkg, store(pkg).part.partName)[0]!.data;
  const next = update(pkg, {
    ...CREATED,
    attributes: { ...CREATED.attributes, grade: '6', added: 'x' },
  });
  expect(customXmlNodes(next, store(next).part.partName)[0]!.data).toBe(before);
  expect(current(next)).toEqual({
    ...CREATED,
    attributes: { ...CREATED.attributes, grade: '6', added: 'x' },
  });
  const registers = record(next).children.filter(
    (child) => child.kind !== 'textValue' && child.localName === 'attr'
  );
  expect(registers).toHaveLength(2);
});

test('a replacement has a fresh identity; removal leaves a valueless register', () => {
  const first = update(created(), {
    ...CREATED,
    attributes: { ...CREATED.attributes, grade: '6' },
  });
  const register = () =>
    record(first).children.find(
      (child) => child.kind !== 'textValue' && child.localName === 'attr'
    )!;
  const second = update(first, {
    ...CREATED,
    attributes: { ...CREATED.attributes, grade: '7' },
  });
  const after = record(second)
    .children.filter((child) => child.kind !== 'textValue' && child.localName === 'attr')
    .at(-1)!;
  expect(after.id).not.toBe(register().id);
  const { grade: _grade, ...withoutGrade } = CREATED.attributes as JsonObject;
  const removed = update(second, { ...CREATED, attributes: withoutGrade });
  expect(current(removed).attributes).toEqual(withoutGrade);
  const kept = record(removed).children.filter(
    (child) => child.kind !== 'textValue' && child.localName === 'attr'
  ) as OoxmlElement[];
  expect(kept).toHaveLength(3);
  expect(kept.at(-1)!.attributes.some((a) => a.localName === 'value')).toBe(false);
  expect(kept.at(-1)!.id).not.toBe(after.id);
  // Re-adding the key creates another immutable replacement.
  const readded = update(removed, CREATED);
  expect(current(readded)).toEqual(CREATED);
});

test('field replacement preserves unknown metadata and unrelated XML identities', () => {
  const pkg = withAuthoredChildren(
    created(),
    '<attr name="grade" value="&quot;A&quot;" opaque="keep"/><future data="untouched"/>'
  );
  const before = record(pkg);
  const oldField = before.children.find(
    (node) => node.kind !== 'textValue' && node.localName === 'attr'
  ) as OoxmlElement;
  const opaque = before.children.find(
    (node) => node.kind !== 'textValue' && node.localName === 'future'
  );
  const next = update(pkg, { ...CREATED, attributes: { ...CREATED.attributes, grade: 'B' } });
  const after = record(next);
  const field = after.children
    .filter((node) => node.kind !== 'textValue' && node.localName === 'attr')
    .at(-1) as OoxmlElement;
  expect(field.id).not.toBe(oldField.id);
  expect(after.children.find((node) => node.id === oldField.id)).toBe(oldField);
  expect(oldField.attributes.find((attribute) => attribute.localName === 'opaque')?.value).toBe(
    'keep'
  );
  expect(after.children.find((node) => node.id === opaque!.id)).toBe(opaque);
  expect(customXmlNodes(next, store(next).part.partName)[0]!.data).toBe(
    customXmlNodes(pkg, store(pkg).part.partName)[0]!.data
  );
});

test('top-level fields are registers too; identity and version stay fixed', () => {
  const next = update(created(), {
    ...CREATED,
    type: 'lesson_title_v2',
    mdast: { type: 'root', children: [{ type: 'text', value: 'x' }] },
  });
  expect(current(next).type).toBe('lesson_title_v2');
  expect(current(next).mdast).toEqual({ type: 'root', children: [{ type: 'text', value: 'x' }] });
  const { mdast: _mdast, ...withoutMdast } = CREATED;
  const removed = update(next, withoutMdast as BlockPayload);
  expect('mdast' in current(removed)).toBe(false);
  const fields = record(removed).children.filter(
    (child) => child.kind !== 'textValue' && child.localName === 'field'
  ) as OoxmlElement[];
  expect(fields.map((f) => f.attributes.find((a) => a.localName === 'name')!.value).sort()).toEqual(
    ['mdast', 'mdast', 'type', 'type']
  );
});

test('an unchanged payload writes nothing; save/reopen keeps the effective payload', () => {
  const pkg = update(created(), {
    ...CREATED,
    attributes: { ...CREATED.attributes, grade: '6' },
  });
  // Same content, different key order.
  const reordered = update(pkg, {
    version: 1,
    attributes: { nested: { a: [1, 2] }, lessonMinutes: 45, grade: '6' },
    type: 'lesson_title',
    id: 'block-1',
    mdast: { children: [], type: 'root' },
  });
  expect(reordered).toBe(pkg);
  const reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(current(reopened.package)).toEqual(current(pkg));
});

test('duplicate registers resolve to the last one and retain their history anchors', () => {
  const pkg = withAuthoredChildren(
    created(),
    `<attr name="grade" value="&quot;A&quot;"/><attr name="grade" value="&quot;B&quot;"/>`
  );
  expect(current(pkg).attributes.grade).toBe('B');
  const next = update(pkg, { ...CREATED, attributes: { ...CREATED.attributes, grade: 'C' } });
  expect(current(next).attributes.grade).toBe('C');
  const grades = record(next).children.filter(
    (child) =>
      child.kind !== 'textValue' &&
      child.localName === 'attr' &&
      child.attributes.some((a) => a.localName === 'name' && a.value === 'grade')
  );
  expect(grades).toHaveLength(3);
});

test('malformed registers make the store invalid rather than silently dropped', () => {
  for (const [extra, reason] of [
    ['<attr name="grade" value="{"/>', 'block-1:invalid-field-value'],
    ['<attr value="1"/>', 'block-1:invalid-field-name'],
    ['<field name="id" value="&quot;other&quot;"/>', 'block-1:invalid-field-name'],
    ['<field name="version" value="2"/>', 'block-1:invalid-field-name'],
    ['<field name="attributes" value="{}"/>', 'block-1:invalid-field-name'],
    ['<field name="type" value="7"/>', 'block-1:invalid-type'],
  ] as const) {
    const read = readPayloadStore(withAuthoredChildren(created(), extra));
    expect(read).toEqual({ status: 'invalid', reason });
  }
});

test('records without registers keep their exact created JSON', () => {
  const pkg = created();
  const entry = store(pkg).entries.get('block-1')!;
  expect(entry.raw).toBe(customXmlNodes(pkg, store(pkg).part.partName)[0]!.data);
});

test('field capacity refuses writes and preparation recovers capacity without changing payload data', () => {
  const pkg = withAuthoredChildren(created(), '<attr name="grade" value="&quot;A&quot;"/>');
  const original = record(pkg);
  const field = original.children.at(-1)!;
  const part = pkg.parts.get(store(pkg).part.partName)!;
  const children = [
    ...original.children.slice(0, -1),
    ...Array.from({ length: 10_000 }, (_, index) => ({ ...field, id: `history-${index}` })),
  ];
  const replaced = replaceChildren(part, original.id, children);
  if (!replaced.ok) throw Error('Cannot construct field capacity fixture');
  const full = withPart(pkg, replaced.part);
  const changed = { ...CREATED, attributes: { ...CREATED.attributes, grade: 'B' } };
  expect(withPayloadRecord(full, changed, 'Title')).toMatchObject({
    ok: false,
    pkg: full,
    reason: 'block-1:field-limit',
  });
  const compact = withPreparedPayloadFields(full);
  expect(current(compact)).toEqual(current(full));
  expect(customXmlNodes(compact, store(compact).part.partName)[0]!.data).toBe(
    customXmlNodes(full, store(full).part.partName)[0]!.data
  );
  expect(
    record(compact).children.filter(
      (node) => node.kind !== 'textValue' && node.localName === 'attr'
    )
  ).toHaveLength(1);
  expect(withPayloadRecord(compact, changed, 'Title').ok).toBe(true);
  expect(withPreparedPayloadFields(compact)).toBe(compact);
});

test('preparation keeps opaque field metadata and unrelated XML', () => {
  const pkg = withAuthoredChildren(
    created(),
    '<attr name="grade" value="&quot;A&quot;" opaque="keep"/><future opaque="keep"/><attr name="grade" value="&quot;B&quot;"/><attr name="grade" value="&quot;C&quot;"/>'
  );
  const before = record(pkg);
  const compact = withPreparedPayloadFields(pkg);
  const after = record(compact);
  expect(current(compact)).toEqual(current(pkg));
  expect(after.children).toEqual(before.children.filter((node) => node !== before.children.at(-2)));
});

test('field history bounds each value and the combined retained value bytes', () => {
  const node = record(created());
  const field = (value: string): OoxmlElement => ({
    id: 'field',
    kind: 'generic',
    namespaceUri: node.namespaceUri,
    localName: 'attr',
    namespaceBindings: [],
    children: [],
    attributes: [
      { kind: 'genericExtension', namespaceUri: '', localName: 'name', value: 'grade' },
      { kind: 'genericExtension', namespaceUri: '', localName: 'value', value },
    ],
  });
  expect(
    readFields({
      ...node,
      children: [field(JSON.stringify('x'.repeat(2 * 1024 * 1024)))],
    } as OoxmlElement)
  ).toEqual({ ok: false, reason: 'field-byte-limit' });
  const entry = field(JSON.stringify('x'.repeat(2 * 1024 * 1024 - 2)));
  expect(
    readFields({ ...node, children: Array.from({ length: 5 }, () => entry) } as OoxmlElement)
  ).toEqual({ ok: false, reason: 'field-byte-limit' });
});

test('preparation leaves unregistered payload types unchanged', () => {
  const pkg = update(created(), { ...CREATED, type: 'future_unregistered_type' });
  const repeated = update(pkg, {
    ...current(pkg),
    attributes: { ...CREATED.attributes, grade: '6' },
  });
  const final = update(repeated, {
    ...current(repeated),
    attributes: { ...CREATED.attributes, grade: '7' },
  });
  expect(withPreparedPayloadFields(final)).toBe(final);
});

test('field names remain own data keys without changing object prototypes', () => {
  const pkg = withAuthoredChildren(
    created(),
    '<attr name="__proto__" value="{&quot;polluted&quot;:true}"/><field name="constructor" value="{&quot;kept&quot;:true}"/>'
  );
  const payload = current(pkg);
  expect(Object.getPrototypeOf(payload.attributes)).toBe(Object.prototype);
  expect(Object.hasOwn(payload.attributes, '__proto__')).toBe(true);
  expect(payload.attributes.__proto__).toEqual({ polluted: true });
  expect(Object.hasOwn(payload, 'constructor')).toBe(true);
  expect(({} as Record<string, unknown>).polluted).toBeUndefined();
});
