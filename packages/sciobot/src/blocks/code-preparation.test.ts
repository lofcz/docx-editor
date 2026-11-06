import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  contentControlPropertiesOf,
  findNode,
  serializeOoxmlPart,
  withCustomXmlNode,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { decodeBlockPayload as oldDecode } from '../../test/compat/payload-reader-gallery-v2';
import { codeDocument, openCode } from '../../test/support/code';
import { buildBlockIndex } from './block-index';
import { withNativeCodeSources } from './code-preparation';
import { nativeCodeSource } from './code-native-source';
import { readPayloadStore } from './payload-opc-codec';
import { encodeBlockPayload } from './payload-schema';
import { prepareNativeDocument } from '../native/prepare-document';

test('native code preparation retains source identities, unknown metadata, and exact imported language', async () => {
  const base = openCode(await codeDocument('const x = "<>&";\n\t  x;\n', false));
  const original = buildBlockIndex(base).byId.get('code')!;
  const store = readPayloadStore(base);
  if (store.status !== 'ready') throw Error('Missing store');
  const pkg = withCustomXmlNode(base, store.part.partName, {
    id: 'scio_code',
    label: 'Exact label',
    data: encodeBlockPayload({
      ...original.payload,
      unknown: { retained: ['opaque'] },
      attributes: {
        ...original.payload.attributes,
        language: 'unknown-imported-Language',
        custom: { value: 17 },
      },
    }),
  });
  const before = buildBlockIndex(pkg).byId.get('code')!;
  const part = pkg.parts.get(before.partName)!;
  const control = findNode(part, before.controlId)!;
  const ids = nativeCodeSource(part, control)!.paragraphs.map((paragraph) => paragraph.id);
  const prepared = withNativeCodeSources(pkg);
  const after = buildBlockIndex(prepared).byId.get('code')!;
  expect(after.payload).toEqual({ ...before.payload, version: 2 });
  expect(after.controlId).toBe(before.controlId);
  const nativePart = prepared.parts.get(after.partName)!;
  const nativeControl = findNode(nativePart, after.controlId)!;
  expect(
    nativeCodeSource(nativePart, nativeControl)!.paragraphs.map((paragraph) => paragraph.id)
  ).toEqual(ids);
  expect(contentControlPropertiesOf(nativeControl).lock).toBe('contentLocked');
  expect(withNativeCodeSources(prepared)).toBe(prepared);
  expect(buildBlockIndex(openCode(writeOoxmlPackage(prepared))).byId.get('code')!.payload).toEqual(
    after.payload
  );
});

test('the pinned pre-change reader keeps code version 2 opaque and generic source edits refuse', async () => {
  const pkg = openCode(await codeDocument());
  const store = readPayloadStore(pkg);
  if (store.status !== 'ready') throw Error('Missing store');
  const record = store.entries.get('code')!;
  const legacy = oldDecode(record.raw);
  expect(legacy.status).toBe('unsupported-version');
  expect('raw' in legacy && legacy.raw).toBe(record.raw);
  const block = buildBlockIndex(pkg).byId.get('code')!;
  const part = pkg.parts.get(block.partName)!;
  const source = nativeCodeSource(part, findNode(part, block.controlId)!)!;
  expect(
    applyTreeOp(part, {
      op: 'insertText',
      paragraphId: source.paragraphs[0]!.id,
      offset: 0,
      text: 'legacy',
    })
  ).toMatchObject({ ok: false, reason: 'locked' });
  const reopened = openCode(writeOoxmlPackage(pkg));
  expect(serializeOoxmlPart(reopened.parts.get(block.partName)!)).toBe(serializeOoxmlPart(part));
  const saved = readPayloadStore(reopened);
  expect(saved.status === 'ready' && oldDecode(saved.entries.get('code')!.raw).status).toBe(
    'unsupported-version'
  );
});

test('ordinary preparation preserves version 1 unless fresh-document code preparation is explicit', async () => {
  const original = await codeDocument('one', false);
  const prepared = prepareNativeDocument(original);
  if (!prepared.success) throw Error(prepared.error);
  expect(buildBlockIndex(openCode(prepared.document)).byId.get('code')!.payload.version).toBe(1);
  const native = prepareNativeDocument(original, { nativeCodeSources: true });
  if (!native.success) throw Error(native.error);
  expect(buildBlockIndex(openCode(native.document)).byId.get('code')!.payload.version).toBe(2);
  expect(
    buildBlockIndex(openCode(native.document)).byId.get('code')!.payload.attributes.value
  ).toBe('one');
  const repeated = prepareNativeDocument(native.document, { nativeCodeSources: true });
  expect(repeated.success && repeated.changed).toBe(false);
});

for (const replacement of [
  '<w:fldSimple w:instr="DATE"><w:r><w:t>one</w:t></w:r></w:fldSimple>',
  '<w:r><w:t>changed in Word</w:t></w:r>',
])
  test(`unsupported imported code keeps its authored source and fallback: ${replacement}`, async () => {
    const entries = unzipSync(await codeDocument('one\nmiddle\nlast', false));
    entries['word/document.xml'] = strToU8(
      strFromU8(entries['word/document.xml']!).replace(
        /<w:r><w:rPr>[\s\S]*?<\/w:rPr><w:t[^>]*>one<\/w:t><\/w:r>/,
        replacement
      )
    );
    const pkg = openCode(zipSync(entries));
    const before = serializeOoxmlPart(pkg.parts.get(pkg.mainDocumentPart)!);
    const prepared = withNativeCodeSources(pkg);
    expect(prepared).toBe(pkg);
    expect(buildBlockIndex(prepared).byId.get('code')!.payload.version).toBe(1);
    expect(serializeOoxmlPart(prepared.parts.get(prepared.mainDocumentPart)!)).toBe(before);
  });
