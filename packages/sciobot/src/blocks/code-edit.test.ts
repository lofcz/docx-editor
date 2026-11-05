import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { findNode, paragraphTextOf, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { codeDocument, openCode } from '../../test/support/code';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { buildBlockIndex } from './block-index';
import { updateCodeLanguage, updateCodePayload, updateCodeSource } from './code-edit';
import { updateAtomicBlock } from './update-atomic';
import { nativeCodeSource } from './code-native-source';

const ORIGINAL = 'one\nmiddle\nlast';
const read = (editor: DocxEditorInstance) => {
  const pkg = editor.surface!.session.currentPackage();
  const block = buildBlockIndex(pkg).byId.get('code');
  if (!block) return null;
  const part = pkg.parts.get(block.partName)!;
  const native = nativeCodeSource(part, findNode(part, block.controlId)!);
  expect(typeof block.payload.attributes.value).toBe('string');
  expect(native?.value).toBe(block.payload.attributes.value as string);
  return { block, source: native!.value, ids: native!.paragraphs.map((paragraph) => paragraph.id) };
};

test('the existing atomic API and code payload helper retain v2 metadata-only writes without source history', async () => {
  const pair = await createTwoEditorFixture({ document: await codeDocument() });
  const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
  try {
    const initial = read(alice)!;
    expect(
      updateAtomicBlock(alice, {
        blockId: 'code',
        attributes: {
          ...initial.block.payload.attributes,
          language: 'typescript',
          opaque: { retained: ['value'] },
        },
      })
    ).toMatchObject({ success: true });
    await pair.settle();
    for (const editor of pair.editors) {
      expect(read(editor)!.ids).toEqual(initial.ids);
      expect(read(editor)!.source).toBe(ORIGINAL);
      expect(read(editor)!.block.payload.attributes).toEqual({
        value: ORIGINAL,
        language: 'typescript',
        opaque: { retained: ['value'] },
      });
    }
    const payload = read(bob)!.block.payload;
    expect(
      updateCodePayload(bob, payload, {
        ...payload,
        attributes: {
          ...payload.attributes,
          language: 'javascript',
          opaque: { retained: ['changed'] },
        },
      })
    ).toMatchObject({ success: true });
    await pair.settle();
    expect(read(alice)!.block.payload.attributes.opaque).toEqual({ retained: ['changed'] });
    expect(read(alice)!.block.payload.attributes.language).toBe('javascript');
    expect(read(alice)!.ids).toEqual(initial.ids);
    expect(bob.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(read(alice)!.block.payload.attributes.language).toBe('typescript');
    expect(alice.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(read(alice)!.block.payload.attributes).toEqual({ value: ORIGINAL, language: 'python' });
    expect(read(alice)!.ids).toEqual(initial.ids);
  } finally {
    pair.destroy();
  }
});

test('code source refuses malicious values and external protection before writes', async () => {
  for (const protection of [
    '',
    '<w:lock w:val="contentLocked"/>',
    '<w:dataBinding w:xpath="/value" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/>',
  ]) {
    const entries = unzipSync(await codeDocument());
    if (protection)
      entries['word/document.xml'] = strToU8(
        strFromU8(entries['word/document.xml']!)
          .replace('<w:body>', `<w:body><w:sdt><w:sdtPr>${protection}</w:sdtPr><w:sdtContent>`)
          .replace('</w:body>', '</w:sdtContent></w:sdt></w:body>')
      );
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({ container, document: zipSync(entries) });
    try {
      const before = editor.surface!.session.currentPackage();
      const invalid = protection
        ? ['changed']
        : [null, 7, {}, 'x'.repeat(2 * 1024 * 1024 + 1), 'first\r\nsecond'];
      for (const value of invalid) {
        expect(
          updateCodeSource(editor, {
            blockId: 'code',
            expectedSource: ORIGINAL,
            value: value as string,
          }).success
        ).toBe(false);
        expect(editor.surface!.session.currentPackage()).toBe(before);
      }
      if (protection)
        expect(
          updateCodeLanguage(editor, { blockId: 'code', language: 'javascript' }).success
        ).toBe(false);
    } finally {
      editor.destroy();
      container.remove();
    }
  }
});

test('native code edits preserve whitespace, untouched lines, independent language, and one-step history', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: await codeDocument(), author: 'Alice' });
  try {
    const ids = read(editor)!.ids;
    const value = 'one 😀\n\tmiddle  x\n\nlast';
    expect(
      updateCodeSource(editor, { blockId: 'code', expectedSource: ORIGINAL, value })
    ).toMatchObject({ success: true });
    expect(read(editor)!.source).toBe(value);
    expect(read(editor)!.ids[0]).toBe(ids[0]);
    expect(read(editor)!.ids.at(-1)).toBe(ids.at(-1));
    expect(updateCodeLanguage(editor, { blockId: 'code', language: 'typescript' })).toMatchObject({
      success: true,
    });
    expect(read(editor)!.source).toBe(value);
    const reopened = openCode(new Uint8Array(await editor.save()));
    expect(buildBlockIndex(reopened).byId.get('code')!.payload.attributes).toEqual({
      language: 'typescript',
      value,
    });
    expect((await projectSourcePackage(reopened)).text).toContain(
      `\`\`\`typescript\n${value}\n\`\`\``
    );
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(read(editor)!.block.payload.attributes.language).toBe('python');
    expect(read(editor)!.source).toBe(value);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(read(editor)!.source).toBe(ORIGINAL);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(read(editor)!.source).toBe(value);
    expect(
      updateCodeSource(editor, { blockId: 'code', expectedSource: ORIGINAL, value: 'stale' })
    ).toMatchObject({ success: false, code: 'stale-revision' });
    expect(read(editor)!.source).toBe(value);
  } finally {
    editor.destroy();
    container.remove();
  }
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`native code source merges peer typing, line splits, and language with exact histories (${clientIds.join('/')})`, async () => {
    const pair = await createTwoEditorFixture({ document: await codeDocument(), clientIds });
    const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
    try {
      pair.pause();
      expect(
        updateCodeSource(alice, {
          blockId: 'code',
          expectedSource: ORIGINAL,
          value: 'one Alice\nmiddle\n\tnew\nlast',
        })
      ).toMatchObject({ success: true });
      expect(read(alice)!.source).toBe('one Alice\nmiddle\n\tnew\nlast');
      expect(
        updateCodeSource(bob, {
          blockId: 'code',
          expectedSource: ORIGINAL,
          value: 'Bob one\nmiddle\nlast Bob',
        })
      ).toMatchObject({ success: true });
      expect(read(bob)!.source).toBe('Bob one\nmiddle\nlast Bob');
      expect(updateCodeLanguage(bob, { blockId: 'code', language: 'javascript' })).toMatchObject({
        success: true,
      });
      expect(read(bob)!.source).toBe('Bob one\nmiddle\nlast Bob');
      pair.reconnect();
      await pair.settle();
      expect(pair.rooms.map((room) => room.session.statusSnapshot().status)).toEqual([
        'ready',
        'ready',
      ]);
      for (const editor of pair.editors) {
        expect(read(editor)!.source).toBe('Bob one Alice\nmiddle\n\tnew\nlast Bob');
        expect(read(editor)!.block.payload.attributes.language).toBe('javascript');
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(read(editor)!.source).toBe('Bob one\nmiddle\nlast Bob');
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(read(alice)!.source).toBe('Bob one\nmiddle\nlast Bob');
      expect(read(alice)!.block.payload.attributes.language).toBe('python');
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(read(editor)!.source).toBe(ORIGINAL);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      expect(bob.exec({ type: 'redo' }).ok).toBe(true);
      expect(bob.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(read(editor)!.source).toBe('Bob one Alice\nmiddle\n\tnew\nlast Bob');
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
      const reopened = openCode(writeOoxmlPackage(alice.surface!.session.currentPackage()));
      expect(buildBlockIndex(reopened).byId.get('code')!.payload.attributes.value).toBe(
        'Bob one Alice\nmiddle\n\tnew\nlast Bob'
      );
      const tail = alice.surface!.session.paragraphIds().at(-1)!;
      expect(paragraphTextOf(alice.surface!.session.part(), tail)).toBe('Tail');
    } finally {
      pair.destroy();
    }
  });

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const) {
  test(`code line joins preserve peer text through both histories (${clientIds.join('/')})`, async () => {
    const pair = await createTwoEditorFixture({ document: await codeDocument(), clientIds });
    const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
    try {
      pair.pause();
      expect(
        updateCodeSource(alice, {
          blockId: 'code',
          expectedSource: ORIGINAL,
          value: 'onemiddle\nlast',
        })
      ).toMatchObject({ success: true });
      expect(
        updateCodeSource(bob, {
          blockId: 'code',
          expectedSource: ORIGINAL,
          value: 'one\nmiddle Bob\nlast',
        })
      ).toMatchObject({ success: true });
      pair.reconnect(1);
      await pair.settle();
      for (const editor of pair.editors) expect(read(editor)!.source).toBe('onemiddle Bob\nlast');
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(read(editor)!.source).toBe('one\nmiddle Bob\nlast');
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(read(editor)!.source).toBe(ORIGINAL);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      expect(bob.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(read(editor)!.source).toBe('onemiddle Bob\nlast');
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      pair.destroy();
    }
  });

  test(`code deletion refuses stale drafts and undo retains peer source (${clientIds.join('/')})`, async () => {
    const pair = await createTwoEditorFixture({ document: await codeDocument(), clientIds });
    const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
    try {
      pair.pause();
      expect(
        updateCodeSource(alice, {
          blockId: 'code',
          expectedSource: ORIGINAL,
          value: 'one Alice\nmiddle\nlast',
        })
      ).toMatchObject({ success: true });
      const controlId = read(bob)!.block.controlId;
      expect(
        bob.surface!.applyAutomationOps(
          () => [{ op: 'removeContentControl', controlId, keepContent: false }],
          { kind: 'body' },
          undefined,
          false,
          true
        ).committed
      ).toBe(true);
      pair.reconnect();
      await pair.settle();
      for (const editor of pair.editors) expect(read(editor)).toBe(null);
      expect(
        updateCodeSource(alice, { blockId: 'code', expectedSource: ORIGINAL, value: 'stale' })
      ).toMatchObject({ success: false, code: 'stale-revision' });
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(read(editor)!.source).toBe('one Alice\nmiddle\nlast');
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(read(editor)!.source).toBe(ORIGINAL);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(bob.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(read(editor)).toBe(null);
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
    } finally {
      pair.destroy();
    }
  });
}
