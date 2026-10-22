import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  customXmlDataParts,
  customXmlNodes,
  readOoxmlPackage,
  readOoxmlPart,
  serializeOoxmlPart,
  replaceNode,
  withCustomXmlDataPart,
  withCustomXmlNode,
  withPart,
  writeOoxmlPackage,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { propertyRecordFixture } from '../../test/support/property-record';
import {
  labelEntry,
  labelPayload,
  labelStore,
  legacyLabelNode,
} from '../../test/support/payload-label';
import { withPayloadRecord, readPayloadStore } from './payload-opc-codec';
import {
  isPayloadLabel,
  PAYLOAD_LABEL_ELEMENT,
  PAYLOAD_LABEL_LIMIT,
  PAYLOAD_LABEL_RECORD_LIMITS,
  readPayloadLabel,
  withPayloadLabel,
} from './payload-label';

function record(pkg: OoxmlPackage): OoxmlElement {
  return pkg.parts
    .get(labelStore(pkg).part.partName)!
    .root.children.find(
      (node) =>
        node.kind !== 'textValue' &&
        node.attributes.some((a) => a.localName === 'id' && a.value === 'scio_activity')
    ) as OoxmlElement;
}
function update(pkg: OoxmlPackage, label: string) {
  const result = withPayloadRecord(pkg, labelPayload(pkg), label);
  if (!result.ok) throw new Error(result.reason);
  return result.pkg;
}
function saved(pkg: OoxmlPackage) {
  const result = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!result.ok) throw new Error(result.reason);
  return result.package;
}

test('a label-only update keeps created label/data text, payloads, native parts, and no-op identity intact', () => {
  const pkg = propertyRecordFixture(),
    old = legacyLabelNode(pkg),
    payload = labelPayload(pkg),
    next = update(pkg, 'New activity'),
    node = record(next);
  expect(labelEntry(next).label).toBe('New activity');
  expect(labelPayload(next)).toEqual(payload);
  expect(legacyLabelNode(next).data).toBe(old.data);
  expect(legacyLabelNode(next).label).toBe(old.label);
  expect(next.parts.get(pkg.mainDocumentPart)).toBe(pkg.parts.get(pkg.mainDocumentPart));
  expect(node.id).toBe(record(pkg).id);
  expect(node.children.slice(0, -1)).toEqual([...record(pkg).children]);
  expect(
    (node.children.at(-1) as OoxmlElement).attributes.find((a) => a.localName === 'value')!.value
  ).toBe('"New activity"');
  expect(update(next, 'New activity')).toBe(next);
  expect(labelEntry(saved(next)).label).toBe('New activity');
  expect(labelEntry(pkg).label).toBe(old.label);
});

test('legacy and scalar labels retain empty text, XML whitespace, quotes, multilingual text, and bounded lengths', () => {
  for (const label of [
    '',
    '  Žák\t\r\n尾 😀  ',
    '"quoted" & <tags> \\ path',
    'x'.repeat(PAYLOAD_LABEL_LIMIT),
  ]) {
    const pkg = update(propertyRecordFixture(), label);
    expect(labelEntry(pkg).label).toBe(label);
    expect(labelEntry(saved(pkg)).label).toBe(label);
    expect(isPayloadLabel(label)).toBe(true);
  }
  for (const label of [
    'x'.repeat(PAYLOAD_LABEL_LIMIT + 1),
    'bad\u0000label',
    'bad\uFFFElabel',
    'bad\uD800label',
  ]) {
    expect(isPayloadLabel(label)).toBe(false);
    const pkg = propertyRecordFixture(),
      before = writeOoxmlPackage(pkg);
    expect(withPayloadRecord(pkg, labelPayload(pkg), label)).toMatchObject({
      ok: false,
      pkg,
      reason: 'invalid-label',
    });
    expect(writeOoxmlPackage(pkg)).toEqual(before);
  }
});

test('malformed, oversized, duplicate, and non-string scalar labels refuse without changing source packages', () => {
  const pkg = propertyRecordFixture(),
    node = record(pkg),
    part = pkg.parts.get(labelStore(pkg).part.partName)!;
  for (const raw of [
    'not-json',
    'null',
    '{}',
    '[]',
    '3',
    'true',
    '"bad\\ud800label"',
    '"bad\\u0000label"',
    'x'.repeat(PAYLOAD_LABEL_LIMIT * 2 + 3),
  ]) {
    const malformed = {
      ...node,
      children: [
        ...node.children,
        {
          id: 'label_malformed',
          kind: 'generic',
          namespaceUri: node.namespaceUri,
          localName: PAYLOAD_LABEL_ELEMENT,
          namespaceBindings: [],
          attributes: [
            { kind: 'genericExtension', namespaceUri: '', localName: 'version', value: '1' },
            { kind: 'genericExtension', namespaceUri: '', localName: 'value', value: raw },
          ],
          children: [],
        },
      ],
    } as OoxmlElement;
    expect(readPayloadLabel(malformed, 'legacy').ok).toBe(false);
    const replaced = replaceNode(part, node.id, malformed);
    if (!replaced.ok) throw new Error('Malformed fixture refused');
    const candidate = withPart(pkg, replaced.part),
      before = writeOoxmlPackage(candidate);
    expect(readPayloadStore(candidate).status).toBe('invalid');
    expect(withPayloadRecord(candidate, labelPayload(pkg), 'New')).toMatchObject({
      ok: false,
      pkg: candidate,
    });
    expect(writeOoxmlPackage(candidate)).toEqual(before);
  }
  const invalid = {
    id: 'label_duplicate',
    kind: 'generic',
    namespaceUri: node.namespaceUri,
    localName: PAYLOAD_LABEL_ELEMENT,
    namespaceBindings: [],
    attributes: [
      { kind: 'genericExtension', namespaceUri: '', localName: 'version', value: '1' },
      { kind: 'genericExtension', namespaceUri: '', localName: 'value', value: '"new"' },
      { kind: 'genericExtension', namespaceUri: '', localName: 'value', value: '"other"' },
    ],
    children: [],
  } as OoxmlElement;
  expect(
    readPayloadLabel({ ...node, children: [...node.children, invalid] } as OoxmlElement, 'old')
  ).toEqual({ ok: false, reason: 'invalid-label-register' });
});

test('label writes preserve foreign stores, binding XML, opaque record attributes, and future payload semantics', () => {
  const pkg = propertyRecordFixture(),
    foreign = withCustomXmlDataPart(pkg, pkg.mainDocumentPart, 'urn:foreign', 'foreign');
  if (!foreign.part) throw new Error('Missing foreign store');
  let candidate = withCustomXmlNode(foreign.pkg, foreign.part.partName, {
    id: 'foreign',
    label: 'Word label',
    data: ' opaque ',
  });
  const body = candidate.parts.get(candidate.mainDocumentPart)!,
    bound = readOoxmlPart(
      serializeOoxmlPart(body).replace(
        '<w:sdtPr>',
        `<w:sdtPr><w:dataBinding w:storeItemID="${foreign.part.itemId}" w:prefixMappings="xmlns:f='urn:foreign'" w:xpath="/f:foreign/f:node[@id='foreign']/f:label"/>`
      ),
      { name: body.name, contentType: body.contentType }
    );
  if (!bound.ok) throw new Error(bound.reason);
  candidate = withPart(candidate, bound.part);
  const descriptor = labelStore(candidate).part,
    node = record(candidate),
    part = candidate.parts.get(descriptor.partName)!;
  const extended = {
    ...node,
    namespaceBindings: [...node.namespaceBindings, { prefix: 'f', namespaceUri: 'urn:foreign' }],
    attributes: [
      ...node.attributes,
      {
        kind: 'genericExtension' as const,
        prefix: 'f',
        namespaceUri: 'urn:foreign',
        localName: 'opaque',
        value: ' keep ',
      },
    ],
  } as OoxmlElement;
  const replaced = replaceNode(part, node.id, extended);
  if (!replaced.ok) throw new Error('Cannot extend record');
  candidate = withPart(candidate, replaced.part);
  const before = canonicalOoxmlFingerprint(candidate.parts.get(candidate.mainDocumentPart)!),
    next = update(candidate, 'New activity');
  expect(canonicalOoxmlFingerprint(next.parts.get(next.mainDocumentPart)!)).toBe(before);
  expect(serializeOoxmlPart(next.parts.get(next.mainDocumentPart)!)).toContain(
    `w:xpath="/f:foreign/f:node[@id=&apos;foreign&apos;]/f:label"`
  );
  expect(next.parts.get(foreign.part.partName)).toBe(candidate.parts.get(foreign.part.partName));
  expect(record(next).attributes.find((a) => a.namespaceUri === 'urn:foreign')!.value).toBe(
    ' keep '
  );
  const futureRaw = ' {"version":2,"id":"activity","type":"future","unknown":" keep "} ',
    future = withCustomXmlNode(candidate, descriptor.partName, {
      id: 'scio_activity',
      label: 'Future Word label',
      data: futureRaw,
    });
  const futureNode = record(future),
    futurePart = future.parts.get(descriptor.partName)!,
    changed = replaceNode(futurePart, futureNode.id, {
      ...futureNode,
      attributes: [
        ...futureNode.attributes,
        {
          kind: 'genericExtension',
          namespaceUri: '',
          localName: PAYLOAD_LABEL_ELEMENT,
          value: 'future non-json semantics',
        },
      ],
    } as OoxmlElement);
  if (!changed.ok) throw new Error('Cannot add opaque future label');
  const opaque = withPart(future, changed.part),
    read = labelEntry(opaque);
  expect(read.decoded.status).toBe('unsupported-version');
  expect(read.label).toBe('Future Word label');
  expect(read.raw).toBe(futureRaw);
  expect(withPayloadRecord(opaque, labelPayload(pkg), 'New')).toMatchObject({
    ok: false,
    pkg: opaque,
    reason: 'unsupported-version',
  });
  expect(
    customXmlNodes(saved(opaque), descriptor.partName).find((n) => n.id === 'scio_activity')!.data
  ).toBe(futureRaw);
  expect(customXmlDataParts(next, next.mainDocumentPart)).toHaveLength(
    customXmlDataParts(candidate, candidate.mainDocumentPart).length
  );
});

test('the scalar writer preserves unrelated record fields and does not change a known value twice', () => {
  const original = record(propertyRecordFixture()),
    next = withPayloadLabel(original, 'New');
  expect(next.children.slice(0, -1)).toEqual([...original.children]);
  expect(next.id).toBe(original.id);
  expect(withPayloadLabel(next, 'New')).toBe(next);
  expect(readPayloadLabel(undefined, 'Legacy')).toEqual({ ok: true, label: 'Legacy' });
  expect(() => withPayloadLabel(original, 'bad\u0000')).toThrow('invalid-label');
});

test('immutable label rows retain earlier values, choose the last authored row, and skip empty undo wrappers', () => {
  const original = record(propertyRecordFixture()),
    first = withPayloadLabel(original, 'First'),
    second = withPayloadLabel(first, 'Second');
  expect(second.children.slice(0, first.children.length)).toEqual([...first.children]);
  expect(readPayloadLabel(second, 'Legacy')).toEqual({ ok: true, label: 'Second' });
  const reversed = {
    ...second,
    children: [...original.children, second.children.at(-1)!, first.children.at(-1)!],
  } as OoxmlElement;
  expect(readPayloadLabel(reversed, 'Legacy')).toEqual({ ok: true, label: 'First' });
  const empty = {
    ...(second.children.at(-1) as OoxmlElement),
    id: 'label_empty',
    attributes: [],
    children: [],
  } as OoxmlElement;
  expect(
    readPayloadLabel({ ...first, children: [...first.children, empty] } as OoxmlElement, 'Legacy')
  ).toEqual({ ok: true, label: 'First' });
  const future = {
    ...empty,
    attributes: [
      { kind: 'genericExtension', namespaceUri: '', localName: 'version', value: '2' },
      { kind: 'genericExtension', namespaceUri: '', localName: 'value', value: 'opaque future' },
    ],
  } as OoxmlElement;
  expect(
    readPayloadLabel({ ...first, children: [...first.children, future] } as OoxmlElement, 'Legacy')
  ).toEqual({ ok: false, reason: 'unsupported-label-register-version' });
  expect(() =>
    withPayloadLabel({ ...first, children: [...first.children, future] } as OoxmlElement, 'New')
  ).toThrow('unsupported-label-register-version');
});

test('label history bounds count all authored updates, native slots, and aggregate bytes', () => {
  const node = record(propertyRecordFixture()),
    template = withPayloadLabel(node, 'Value').children.at(-1) as OoxmlElement;
  const count = {
    ...node,
    children: Array.from({ length: PAYLOAD_LABEL_RECORD_LIMITS.maxRecords + 1 }, (_, i) => ({
      ...template,
      id: 'label_' + i,
    })),
  } as OoxmlElement;
  expect(readPayloadLabel(count, 'Legacy')).toEqual({
    ok: false,
    reason: 'label-register-count-limit',
  });
  const slots = {
    ...node,
    children: Array.from({ length: PAYLOAD_LABEL_RECORD_LIMITS.maxSlots + 1 }, (_, i) => ({
      ...template,
      id: 'label_' + i,
      attributes: [],
      children: [],
    })),
  } as OoxmlElement;
  expect(readPayloadLabel(slots, 'Legacy')).toEqual({
    ok: false,
    reason: 'label-register-slot-limit',
  });
  const value = JSON.stringify('x'.repeat(1_850_000)),
    bytes = {
      ...node,
      children: Array.from({ length: 5 }, (_, i) => ({
        ...template,
        id: 'label_' + i,
        attributes: template.attributes.map((a) => (a.localName === 'value' ? { ...a, value } : a)),
      })),
    } as OoxmlElement;
  expect(readPayloadLabel(bytes, 'Legacy')).toEqual({
    ok: false,
    reason: 'label-register-byte-limit',
  });
});
