import { expect, test } from 'bun:test';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  paragraphTextOf,
  readOoxmlPackage,
  serializeOoxmlPart,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { blockClipboardDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { beginMarkdownInsertion } from './markdown-stream';

const attr = (node: OoxmlNode, name: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find((attribute) => attribute.localName === name)?.value;
function resolvable(pkg: OoxmlPackage) {
  const numbering = pkg.parts.get('/word/numbering.xml');
  expect(numbering).toBeDefined();
  const abstracts = numbering!.root.children.filter(
    (node) => node.kind !== 'textValue' && node.localName === 'abstractNum'
  );
  const nums = numbering!.root.children.filter(
    (node) => node.kind !== 'textValue' && node.localName === 'num'
  );
  expect(abstracts).toHaveLength(1);
  expect(nums).toHaveLength(1);
  const abstractIds = abstracts.map((node) => attr(node, 'abstractNumId'));
  expect(abstractIds).not.toContain(undefined);
  for (const num of nums) {
    if (num.kind === 'textValue') throw new Error('Expected num element.');
    const reference = num.children.find(
      (node) => node.kind !== 'textValue' && node.localName === 'abstractNumId'
    );
    expect(reference).toBeDefined();
    expect(abstractIds).toContain(attr(reference!, 'val'));
  }
  const numIds = nums.map((node) => attr(node, 'numId'));
  expect(numIds).not.toContain(undefined);
  const xml = serializeOoxmlPart(pkg.parts.get('/word/document.xml')!);
  const references = [...xml.matchAll(/<w:numId w:val="([0-9]+)"/g)].map((match) => match[1]);
  expect(references).toHaveLength(2);
  for (const reference of references) expect(numIds).toContain(reference);
}

for (const reuse of [false, true])
  for (const receiver of [0, 1] as const)
    test(`peer numbering survives author stream undo; reuse=${reuse}, receiver=${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: blockClipboardDocument(),
        clientIds: [73, 41],
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const block = buildBlockIndex(alice.surface!.session.currentPackage()).byId.get(
          'activity-1'
        )!;
        const a = beginMarkdownInsertion(
          alice,
          { documentId: 'peer-lists', epoch: 'mount' },
          { selectedIds: [block.controlId] }
        );
        if (!a.success) throw new Error(a.error);
        if (!reuse) pair.pause();
        expect(a.stream.push('- Alice one\n- Alice two\n\nFinal').success).toBe(true);
        if (reuse) {
          await pair.settle();
          pair.pause();
        }
        const tail = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Tail')!;
        const b = beginMarkdownInsertion(
          bob,
          { documentId: 'peer-lists', epoch: 'mount' },
          { selectedIds: [tail] }
        );
        if (!b.success) throw new Error(b.error);
        expect(b.stream.push('- Bob one\n- Bob two').success).toBe(true);
        expect(b.stream.finish().success).toBe(true);
        pair.reconnect(receiver);
        await pair.settle();
        expect(a.stream.finish().success).toBe(true);
        await pair.settle();
        const landed = await pair.savedDigests();
        expect(landed[0]).toEqual(landed[1]);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!opened.ok) throw new Error(opened.reason);
          resolvable(opened.package);
          const text = editor
            .surface!.session.paragraphIds()
            .map((id) => paragraphTextOf(editor.surface!.session.part(), id));
          expect(text).toContain('Bob one');
          expect(text).toContain('Bob two');
          expect(text).not.toContain('Alice one');
          expect(text).not.toContain('Final');
        }
        const undone = await pair.savedDigests();
        expect(undone[0]).toEqual(undone[1]);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(await pair.savedDigests()).toEqual(landed);
      } finally {
        pair.destroy();
      }
    });

for (const receiver of [0, 1] as const)
  test(`a delayed peer reference restores an undone allocation; receiver=${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: blockClipboardDocument(),
      clientIds: [41, 73],
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const block = buildBlockIndex(alice.surface!.session.currentPackage()).byId.get(
        'activity-1'
      )!;
      const a = beginMarkdownInsertion(
        alice,
        { documentId: 'late-list', epoch: 'mount' },
        { selectedIds: [block.controlId] }
      );
      if (!a.success) throw new Error(a.error);
      expect(a.stream.push('- Alice one\n- Alice two').success).toBe(true);
      expect(a.stream.finish().success).toBe(true);
      await pair.settle();
      pair.pause();
      const tail = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Tail')!;
      const b = beginMarkdownInsertion(
        bob,
        { documentId: 'late-list', epoch: 'mount' },
        { selectedIds: [tail] }
      );
      if (!b.success) throw new Error(b.error);
      expect(b.stream.push('- Bob one\n- Bob two').success).toBe(true);
      expect(b.stream.finish().success).toBe(true);
      // Alice has not received Bob's reference when she undoes the allocation.
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      const alone = readOoxmlPackage(new Uint8Array(await alice.save()));
      if (!alone.ok) throw new Error(alone.reason);
      expect(alone.package.parts.get('/word/numbering.xml')!.root.children).toHaveLength(0);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
        if (!opened.ok) throw new Error(opened.reason);
        resolvable(opened.package);
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const restored = await pair.savedDigests();
      expect(restored[0]).toEqual(restored[1]);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(await pair.savedDigests()).toEqual(saved);
    } finally {
      pair.destroy();
    }
  });
