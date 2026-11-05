import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  contentControlPropertiesOf,
  findNode,
  serializeOoxmlPart,
  withPart,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { codeDocument, openCode } from '../../test/support/code';
import { buildBlockIndex } from './block-index';
import { updateCodeLanguage, updateCodeSource } from './code-edit';
import { withNativeCodeSources } from './code-preparation';
import { ownedCodeSourceLock } from './code-source-lock';
import { updateAtomicBlock } from './update-atomic';

const ORIGINAL = 'one\nmiddle\nlast';

test('code preparation preserves every authored content or wrapper lock', async () => {
  const original = openCode(await codeDocument(ORIGINAL, false));
  const block = buildBlockIndex(original).byId.get('code')!;
  for (const lock of ['sdtLocked', 'contentLocked', 'sdtContentLocked'] as const) {
    const result = applyTreeOp(original.parts.get(block.partName)!, {
      op: 'setContentControlProperties',
      controlId: block.controlId,
      lock,
    });
    if (!result.ok) throw Error(result.reason);
    const pkg = withPart(original, result.part);
    expect(withNativeCodeSources(pkg)).toBe(pkg);
    expect(buildBlockIndex(pkg).byId.get('code')!.payload.version).toBe(1);
    expect(contentControlPropertiesOf(findNode(result.part, block.controlId)!).lock).toBe(lock);
  }
});

test('code preparation retains inherited namespace bindings and marks only its generated lock', async () => {
  const entries = unzipSync(await codeDocument(ORIGINAL, false));
  entries['word/document.xml'] = strToU8(
    strFromU8(entries['word/document.xml']!).replace(
      '<w:sdt>',
      '<w:sdt xmlns:scioCode="urn:imported" scioCode:retained="yes">'
    )
  );
  const pkg = withNativeCodeSources(openCode(zipSync(entries)));
  for (const document of [pkg, openCode(writeOoxmlPackage(pkg))]) {
    const block = buildBlockIndex(document).byId.get('code')!;
    const control = findNode(document.parts.get(block.partName)!, block.controlId)!;
    expect(ownedCodeSourceLock(control)).not.toBeNull();
    if (control.kind !== 'contentControl') throw Error('Missing control');
    expect(
      control.attributes.find((attribute) => attribute.namespaceUri === 'urn:imported')?.value
    ).toBe('yes');
  }
});

test('source and metadata commands refuse changed, absent, or unowned code locks before writes', async () => {
  const bytes = await codeDocument();
  const variants = [
    (xml: string) => xml.replace('w:val="contentLocked"', 'w:val="sdtLocked"'),
    (xml: string) => xml.replace('w:val="contentLocked"', 'w:val="sdtContentLocked"'),
    (xml: string) => xml.replace('w:val="contentLocked"', 'w:val="unlocked"'),
    (xml: string) => xml.replace('scioCode:sourceOwner="1"', 'scioCode:sourceOwner="2"'),
    (xml: string) => xml.replace(' scioCode:sourceOwner="1"', ''),
    (xml: string) => xml.replace(/<w:lock[^>]*\/>/, ''),
    (xml: string) =>
      xml.replace(
        '<w:sdtPr>',
        '<w:sdtPr><w:dataBinding w:xpath="/value" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/>'
      ),
  ];
  for (const mutate of variants) {
    const entries = unzipSync(bytes);
    const beforeXml = strFromU8(entries['word/document.xml']!);
    const changed = mutate(beforeXml);
    expect(changed).not.toBe(beforeXml);
    entries['word/document.xml'] = strToU8(changed);
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({ container, document: zipSync(entries) });
    try {
      const before = editor.surface!.session.currentPackage();
      const attributes = buildBlockIndex(before).byId.get('code')!.payload.attributes;
      const calls = [
        () =>
          updateCodeSource(editor, { blockId: 'code', expectedSource: ORIGINAL, value: 'changed' }),
        () => updateCodeLanguage(editor, { blockId: 'code', language: 'javascript' }),
        () =>
          updateAtomicBlock(editor, {
            blockId: 'code',
            attributes: { ...attributes, language: 'javascript' },
          }),
        () =>
          updateAtomicBlock(editor, {
            blockId: 'code',
            attributes: { ...attributes, value: 'changed' },
          }),
      ];
      for (const call of calls) {
        expect(call()).toMatchObject({ success: false, code: 'unsupported-operation' });
        expect(editor.surface!.session.currentPackage()).toBe(before);
      }
    } finally {
      editor.destroy();
      container.remove();
    }
  }
});

test('a peer wrapper-lock change withdraws code ownership until that exact change is undone', async () => {
  const pair = await createTwoEditorFixture({ document: await codeDocument() });
  const [alice, bob] = pair.editors;
  try {
    const pkg = bob!.surface!.session.currentPackage();
    const block = buildBlockIndex(pkg).byId.get('code')!;
    const original = serializeOoxmlPart(pkg.parts.get(block.partName)!);
    expect(
      bob!.surface!.applyAutomationOps(
        () => [
          {
            op: 'setContentControlProperties',
            controlId: block.controlId,
            lock: 'sdtContentLocked',
          },
        ],
        { kind: 'body' },
        undefined,
        false,
        true
      ).committed
    ).toBe(true);
    await pair.settle();
    expect(
      updateCodeSource(alice!, { blockId: 'code', expectedSource: ORIGINAL, value: 'changed' })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
    expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(
      serializeOoxmlPart(alice!.surface!.session.currentPackage().parts.get(block.partName)!)
    ).toBe(original);
    expect(
      updateCodeSource(alice!, {
        blockId: 'code',
        expectedSource: ORIGINAL,
        value: 'one Alice\nmiddle\nlast',
      })
    ).toMatchObject({ success: true });
    await pair.settle();
    expect(
      buildBlockIndex(bob!.surface!.session.currentPackage()).byId.get('code')!.payload.attributes
        .value
    ).toBe('one Alice\nmiddle\nlast');
    expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
  } finally {
    pair.destroy();
  }
});
