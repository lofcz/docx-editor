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

test('an existing register keeps its identity; removal leaves a valueless register', () => {
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
  const after = record(second).children.find(
    (child) => child.kind !== 'textValue' && child.localName === 'attr'
  )!;
  expect(after.id).toBe(register().id);
  const { grade: _grade, ...withoutGrade } = CREATED.attributes as JsonObject;
  const removed = update(second, { ...CREATED, attributes: withoutGrade });
  expect(current(removed).attributes).toEqual(withoutGrade);
  const kept = record(removed).children.filter(
    (child) => child.kind !== 'textValue' && child.localName === 'attr'
  ) as OoxmlElement[];
  expect(kept).toHaveLength(1);
  expect(kept[0]!.attributes.some((a) => a.localName === 'value')).toBe(false);
  // Re-adding the key reuses that register.
  const readded = update(removed, CREATED);
  expect(current(readded)).toEqual(CREATED);
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
    ['mdast', 'type']
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

test('duplicate registers resolve to the last one, and the next write keeps one', () => {
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
  expect(grades).toHaveLength(1);
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
