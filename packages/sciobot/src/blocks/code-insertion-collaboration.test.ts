import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { findNode, readOoxmlPackage } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { codeDocument } from '../../test/support/code';
import { buildBlockIndex } from './block-index';
import { insertBlock } from './insert';
import { nativeCodeSource } from './code-native-source';
import { ownedCodeSourceLock } from './code-source-lock';
import { updateCodeSource } from './code-edit';

test('the first text in a newly created empty code block retains native code formatting', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: await codeDocument('legacy', false) });
  try {
    expect(
      insertBlock(editor, {
        parentControlId: null,
        beforeId: null,
        payload: {
          version: 1,
          id: 'empty-code',
          type: 'code_block',
          attributes: { language: '', value: '' },
        },
        label: 'Code',
      }).success
    ).toBe(true);
    expect(
      updateCodeSource(editor, { blockId: 'empty-code', expectedSource: '', value: 'typed' })
        .success
    ).toBe(true);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw Error(reopened.reason);
    const block = buildBlockIndex(reopened.package).byId.get('empty-code')!;
    const part = reopened.package.parts.get(block.partName)!;
    const source = nativeCodeSource(part, findNode(part, block.controlId)!)!;
    expect(source.value).toBe('typed');
    const run = source.paragraphs[0]!.children.find(
      (node) =>
        node.kind === 'run' &&
        node.children.some(
          (child) =>
            child.kind === 'text' &&
            child.children.some((text) => text.kind === 'textValue' && text.value === 'typed')
        )
    );
    if (!run || run.kind !== 'run') throw Error('Typed native run is missing');
    const properties = run.children.find((node) => node.kind === 'runProperties');
    const style = properties?.children.find((node) => node.localName === 'rStyle');
    expect(style?.attributes.find((attr) => attr.localName === 'val')?.value).toBe('SciobotCode');
    for (const value of ['one\rtwo', 'one\r\ntwo']) {
      const before = editor.surface!.session.currentPackage();
      expect(
        insertBlock(editor, {
          parentControlId: null,
          beforeId: null,
          payload: {
            version: 1,
            id: 'unsupported-new-code',
            type: 'code_block',
            attributes: { language: '', value },
          },
          label: 'Code',
        })
      ).toMatchObject({ success: false, code: 'invalid-input' });
      expect(editor.surface!.session.currentPackage()).toBe(before);
    }
  } finally {
    editor.destroy();
    container.remove();
  }
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`new code owns native source in its insertion transaction and retains peer edits through history ${clientIds}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await codeDocument('legacy', false),
      clientIds,
    });
    try {
      const [alice, bob] = pair.editors;
      const inserted = insertBlock(alice!, {
        parentControlId: null,
        beforeId: null,
        payload: {
          version: 1,
          id: 'new-code',
          type: 'code_block',
          attributes: { language: 'python', value: 'alpha\nbeta' },
        },
        label: 'Code',
      });
      expect(inserted.success).toBe(true);
      await pair.settle();
      const verify = (value: string) => {
        for (const editor of pair.editors) {
          const pkg = editor!.surface!.session.currentPackage();
          const index = buildBlockIndex(pkg);
          expect(index.byId.get('code')!.payload.version).toBe(1);
          const block = index.byId.get('new-code')!;
          expect(block.payload.version).toBe(2);
          expect(block.payload.attributes.value).toBe(value);
          const part = pkg.parts.get(block.partName)!;
          const control = findNode(part, block.controlId)!;
          expect(ownedCodeSourceLock(control)).not.toBeNull();
          expect(nativeCodeSource(part, control)!.value).toBe(value);
        }
      };
      verify('alpha\nbeta');
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(
          buildBlockIndex(editor!.surface!.session.currentPackage()).byId.has('new-code')
        ).toBe(false);
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      verify('alpha\nbeta');
      expect(
        updateCodeSource(bob!, {
          blockId: 'new-code',
          expectedSource: 'alpha\nbeta',
          value: 'alpha\nbeta peer',
        }).success
      ).toBe(true);
      await pair.settle();
      verify('alpha\nbeta peer');
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      verify('alpha\nbeta peer');
      expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      verify('alpha\nbeta');
      expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      verify('alpha\nbeta peer');
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      const reopened = readOoxmlPackage(new Uint8Array(await alice!.save()));
      if (!reopened.ok) throw Error(reopened.reason);
      const block = buildBlockIndex(reopened.package).byId.get('new-code')!;
      expect(block.payload.version).toBe(2);
      expect(block.payload.attributes.value).toBe('alpha\nbeta peer');
    } finally {
      pair.destroy();
    }
  });
