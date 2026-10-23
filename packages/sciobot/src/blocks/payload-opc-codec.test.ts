import { expect, test } from 'bun:test';
import { blankDocumentBytes } from '@docx-editor.dev/core';
import {
  readOoxmlPackage,
  readOoxmlPart,
  writeOoxmlPackage,
  withPart,
  withCustomXmlDataPart,
  withCustomXmlNode,
  customXmlNodes,
  canonicalOoxmlFingerprint,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import fixture from '../__tests__/fixtures/block-payloads.json';
import { decodeBlockPayload, MAX_BLOCK_ID_LENGTH } from './payload-schema';
import {
  PAYLOAD_NAMESPACE,
  PAYLOAD_ROOT,
  blockTag,
  blockIdFromTag,
  payloadReferences,
  readPayloadStore,
  withPayloadRecord,
  withInitializedPayloadStore,
} from './payload-opc-codec';

function emptyPackage() {
  const opened = readOoxmlPackage(blankDocumentBytes());
  if (!opened.ok) throw new Error(opened.reason);
  return opened.package;
}
function payload() {
  const result = decodeBlockPayload(JSON.stringify(fixture.nested));
  if (result.status !== 'supported') throw new Error('Invalid reference payload');
  return result.value;
}
function write(pkg = emptyPackage()) {
  const result = withPayloadRecord(pkg, payload(), 'Warmup <teacher> & Žák 😀');
  if (!result.ok) throw new Error(result.reason);
  return result;
}
function replaceStore(pkg: OoxmlPackage, partName: string, body: string) {
  const parsed = readOoxmlPart(`<blocks xmlns="${PAYLOAD_NAMESPACE}">${body}</blocks>`, {
    name: partName,
    contentType: pkg.parts.get(partName)!.contentType,
  });
  if (!parsed.ok) throw new Error(parsed.reason);
  return withPart(pkg, parsed.part);
}

test('nested payloads use a custom XML part and survive DOCX save/reopen', () => {
  const original = emptyPackage();
  const fingerprint = canonicalOoxmlFingerprint(original.parts.get(original.mainDocumentPart)!);
  const written = write(original);
  expect(written.pkg.parts.get(original.mainDocumentPart)).toBe(
    original.parts.get(original.mainDocumentPart)
  );
  const opened = readOoxmlPackage(writeOoxmlPackage(written.pkg));
  if (!opened.ok) throw new Error(opened.reason);
  const read = readPayloadStore(opened.package);
  expect(read.status).toBe('ready');
  if (read.status !== 'ready') throw new Error('Missing payload store');
  expect(read.entries.get('activity-1')?.decoded).toEqual({
    status: 'supported',
    value: payload(),
  });
  expect(read.entries.get('activity-1')?.label).toBe('Warmup <teacher> & Žák 😀');
  expect(canonicalOoxmlFingerprint(opened.package.parts.get(original.mainDocumentPart)!)).toBe(
    fingerprint
  );
  expect(read.part.propsPartName).toContain('itemProps');
});

test('short stable tags cover the full permitted ID length and reject malformed references', () => {
  const id = 'a'.repeat(MAX_BLOCK_ID_LENGTH);
  expect(blockTag(id)).toHaveLength(64);
  expect(blockIdFromTag(blockTag(id))).toBe(id);
  expect(() => blockTag(id + 'a')).toThrow();
  expect(blockIdFromTag('foreign:tag')).toBeNull();
  expect(payloadReferences([blockTag('same'), blockTag('same')])).toEqual({
    ok: false,
    reason: 'duplicate-reference:same',
  });
  expect(payloadReferences(['scio:1:'])).toEqual({ ok: false, reason: 'malformed-reference' });
});

test('duplicate and malformed XML records fail without changing the package', () => {
  const written = write();
  const duplicate = replaceStore(
    written.pkg,
    written.part.partName,
    '<node id="scio_same"><data>{}</data></node><node id="scio_same"><data>{}</data></node>'
  );
  const record = JSON.stringify({ ...payload(), id: 'same' })
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;');
  const duplicated = replaceStore(
    written.pkg,
    written.part.partName,
    `<node id="scio_same"><data>${record}</data></node><node id="scio_same"><data>${record}</data></node>`
  );
  expect(readPayloadStore(duplicated)).toEqual({
    status: 'invalid',
    reason: 'duplicate-node:same',
  });
  expect(readPayloadStore(duplicate).status).toBe('invalid');
  const refused = withPayloadRecord(duplicated, payload(), 'Label');
  expect(refused.ok).toBe(false);
  expect(refused.pkg).toBe(duplicated);
  const missing = replaceStore(written.pkg, written.part.partName, '<node><data>{}</data></node>');
  expect(readPayloadStore(missing)).toEqual({ status: 'invalid', reason: 'missing-node-id' });
  const mismatched = replaceStore(
    written.pkg,
    written.part.partName,
    `<node id="scio_other"><data>${record}</data></node>`
  );
  expect(readPayloadStore(mismatched)).toEqual({
    status: 'invalid',
    reason: 'mismatched-node-id:other',
  });
});

test('foreign custom parts and opaque future payload records remain intact', () => {
  const foreign = withCustomXmlDataPart(
    emptyPackage(),
    '/word/document.xml',
    'urn:foreign:store',
    'foreign'
  );
  if (!foreign.part) throw new Error('No foreign store');
  const foreignData = withCustomXmlNode(foreign.pkg, foreign.part.partName, {
    id: 'foreign',
    label: 'Foreign',
    data: '{"keep":true}',
  });
  const first = write(foreignData);
  const future = JSON.stringify(fixture.future, null, 2);
  const withFuture = withCustomXmlNode(first.pkg, first.part.partName, {
    id: 'scio_activity-2',
    label: 'Future',
    data: future,
  });
  const second = write(withFuture);
  expect(customXmlNodes(second.pkg, foreign.part.partName)).toEqual(
    customXmlNodes(foreignData, foreign.part.partName)
  );
  const read = readPayloadStore(second.pkg);
  if (read.status !== 'ready') throw new Error('No store');
  expect(read.entries.get('activity-2')?.raw).toBe(future);
  expect(read.entries.get('activity-2')?.decoded.status).toBe('unsupported-version');
  expect(PAYLOAD_ROOT).toBe('blocks');
});

test('a label that cannot survive XML serialization refuses without changing the package', () => {
  const pkg = emptyPackage();
  const refused = withPayloadRecord(pkg, payload(), 'Bad\u0000label');
  expect(refused.ok).toBe(false);
  expect(refused.pkg).toBe(pkg);
});

test('pre-bootstrap payload storage is empty, stable, and preserves populated and invalid imports', () => {
  const pkg = emptyPackage();
  const before = canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
  const initialized = withInitializedPayloadStore(pkg);
  if (!initialized.ok) throw Error(initialized.reason);
  expect(readPayloadStore(initialized.pkg)).toMatchObject({ status: 'ready', entries: new Map() });
  expect(canonicalOoxmlFingerprint(initialized.pkg.parts.get(pkg.mainDocumentPart)!)).toBe(before);
  expect(withInitializedPayloadStore(initialized.pkg)).toEqual({ ok: true, pkg: initialized.pkg });
  const stored = write(initialized.pkg);
  expect(withInitializedPayloadStore(stored.pkg)).toEqual({ ok: true, pkg: stored.pkg });
  const root = stored.pkg.parts.get(stored.part.partName)!;
  if (root.root.kind !== 'generic') throw Error('Expected generic custom XML root');
  const malformed = withPart(stored.pkg, { ...root, root: { ...root.root, localName: 'wrong' } });
  expect(withInitializedPayloadStore(malformed)).toEqual({
    ok: false,
    pkg: malformed,
    reason: 'invalid-store-root',
  });
});
