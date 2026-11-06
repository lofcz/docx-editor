import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { sha256Bytes, type ImageDecodePort, type OoxmlNode } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { galleryPng } from '../../test/support/native-gallery';
import { zipDocument } from '../../../pro/src/collaboration/__tests__/document-peer-support';

const documentBytes = zipDocument('<w:p/><w:p><w:r><w:t>Peer text</w:t></w:r></w:p>');
const replacement = Uint8Array.from(
  atob('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'),
  (c) => c.charCodeAt(0)
);

function drawings(editor: DocxEditorInstance): string[] {
  const pending: OoxmlNode[] = [editor.surface!.session.part().root];
  const found: string[] = [];
  while (pending.length) {
    const node = pending.pop()!;
    if (node.kind === 'drawing') found.push(node.id);
    if (node.kind !== 'textValue') pending.push(...node.children);
  }
  return found;
}
function deferredDecoder() {
  let pending: Promise<void> | null = null;
  let notify = () => {};
  const port: ImageDecodePort = {
    async decode() {
      if (pending) {
        notify();
        await pending;
      }
      return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
    },
  };
  return {
    port,
    hold() {
      let release = () => {};
      const started = new Promise<void>((resolve) => {
        notify = resolve;
      });
      pending = new Promise<void>((resolve) => {
        release = () => {
          pending = null;
          resolve();
        };
      });
      return { started, release };
    },
  };
}
function media(editor: DocxEditorInstance) {
  return [...editor.surface!.session.currentPackage().partBytes]
    .filter(([name]) => name.includes('/media/'))
    .map(([name, bytes]) => [name, sha256Bytes(bytes)])
    .sort();
}

async function documentWithImage() {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: documentBytes,
    imageDecodePort: deferredDecoder().port,
  });
  try {
    const result = await editor.executeImageCommand({
      type: 'insertImage',
      data: galleryPng,
      mime: 'image/png',
      widthPoints: 36,
      heightPoints: 36,
    });
    if (!result.ok) throw Error(result.reason);
    return new Uint8Array(await editor.save());
  } finally {
    editor.destroy();
  }
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const) {
  for (const operation of ['insert', 'replace'] as const)
    test(`an abort during publication keeps a committed ${operation} successful ${clientIds}`, async () => {
      const pair = await createTwoEditorFixture({
        document: operation === 'replace' ? await documentWithImage() : documentBytes,
        clientIds,
        editor: { imageDecodePort: deferredDecoder().port },
      });
      try {
        const editor = pair.editors[0]!;
        await pair.settle();
        const id = drawings(editor)[0];
        if (id)
          expect(
            editor.surface!.selectDrawing(id, editor.surface!.session.paragraphIds()[0]!)
          ).toBe(true);
        const controller = new AbortController();
        const before = await pair.savedDigests();
        const mediaBefore = media(editor);
        const off = editor.surface!.session.subscribe(() => controller.abort());
        const command =
          operation === 'insert'
            ? {
                type: 'insertImage' as const,
                data: replacement,
                mime: 'image/gif' as const,
                widthPoints: 36,
                heightPoints: 36,
              }
            : {
                type: 'replaceImage' as const,
                data: replacement,
                mime: 'image/gif' as const,
                drawingNodeId: id!,
              };
        const result = await editor.executeImageCommand(command, { signal: controller.signal });
        off();
        expect(controller.signal.aborted).toBe(true);
        expect(result).toEqual({ ok: true, changed: true });
        await pair.settle();
        const after = await pair.savedDigests();
        expect(after[0]).toEqual(after[1]);
        expect(after).not.toEqual(before);
        const mediaAfter = media(editor);
        expect(mediaAfter).not.toEqual(mediaBefore);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(await pair.savedDigests()).toEqual(before);
        expect(media(editor)).toEqual(mediaBefore);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(await pair.savedDigests()).toEqual(after);
        expect(media(editor)).toEqual(mediaAfter);
      } finally {
        pair.destroy();
      }
    });
  test(`native insertion cancels during decoding without shared writes ${clientIds}`, async () => {
    const decoder = deferredDecoder();
    const pair = await createTwoEditorFixture({
      document: documentBytes,
      clientIds,
      editor: { imageDecodePort: decoder.port },
    });
    const delayed = decoder.hold();
    try {
      const editor = pair.editors[0]!;
      const controller = new AbortController();
      const before = await pair.savedDigests();
      const shared = pair.documents.map((doc) => Array.from(Y.encodeStateAsUpdate(doc)));
      const pending = editor.executeImageCommand(
        {
          type: 'insertImage',
          data: galleryPng,
          mime: 'image/png',
          widthPoints: 36,
          heightPoints: 36,
        },
        { signal: controller.signal }
      );
      await delayed.started;
      controller.abort();
      delayed.release();
      const result = await pending;
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.reason).toBe('image command was cancelled');
      await pair.settle();
      expect(drawings(editor)).toEqual([]);
      expect(media(editor)).toEqual([]);
      expect(await pair.savedDigests()).toEqual(before);
      expect(pair.documents.map((doc) => Array.from(Y.encodeStateAsUpdate(doc)))).toEqual(shared);
      expect(editor.can({ type: 'undo' }).ok).toBe(false);
    } finally {
      delayed.release();
      pair.destroy();
    }
  });
  for (const change of ['cancel', 'reattach', 'selection', 'delete', 'load', 'view'] as const)
    test(`native replacement keeps every original decode guard with cancellation ${clientIds}/${change}`, async () => {
      const decoder = deferredDecoder();
      const pair = await createTwoEditorFixture({
        document: documentBytes,
        clientIds,
        editor: { imageDecodePort: decoder.port },
      });
      const editor = pair.editors[0]!;
      let release = () => {};
      try {
        expect(
          (
            await editor.executeImageCommand({
              type: 'insertImage',
              data: galleryPng,
              mime: 'image/png',
              widthPoints: 36,
              heightPoints: 36,
            })
          ).ok
        ).toBe(true);
        await pair.settle();
        const owner = editor.surface!;
        const id = drawings(editor)[0]!;
        expect(owner.selectDrawing(id, owner.session.paragraphIds()[0]!)).toBe(true);
        const controller = new AbortController();
        const delayed = decoder.hold();
        release = delayed.release;
        const replacing = editor.executeImageCommand(
          { type: 'replaceImage', data: replacement, mime: 'image/gif', drawingNodeId: id },
          { signal: controller.signal }
        );
        await delayed.started;
        if (change === 'cancel') controller.abort();
        if (change === 'reattach') {
          const generation = editor.mountGeneration;
          editor.detach();
          editor.attach(pair.containers[0]!);
          expect(editor.surface).toBe(owner);
          expect(editor.mountGeneration).toBeGreaterThan(generation);
        }
        if (change === 'selection') {
          const paragraphId = owner.session.paragraphIds()[1]!;
          owner.setSelection({
            anchor: { paragraphId, offset: 1 },
            head: { paragraphId, offset: 1 },
          });
        }
        if (change === 'delete')
          expect(editor.exec({ type: 'deleteImage', drawingNodeId: id }).ok).toBe(true);
        if (change === 'load') editor.load(await editor.save());
        if (change === 'view') editor.surface!.setEditingMode('view');
        await pair.settle();
        const before = editor.surface!.session.currentPackage();
        const revision = editor.surface!.session.packageRevision();
        const beforeMedia = media(editor);
        const shared = pair.documents.map((doc) => Array.from(Y.encodeStateAsUpdate(doc)));
        const history = pair.editors.map((peer) => ({
          undo: peer.can({ type: 'undo' }),
          redo: peer.can({ type: 'redo' }),
        }));
        release();
        const result = await replacing;
        expect(result.ok).toBe(false);
        if (!result.ok && change === 'cancel')
          expect(result.reason).toBe('image command was cancelled');
        expect(editor.surface!.session.currentPackage()).toBe(before);
        expect(editor.surface!.session.packageRevision()).toBe(revision);
        expect(media(editor)).toEqual(beforeMedia);
        expect(pair.documents.map((doc) => Array.from(Y.encodeStateAsUpdate(doc)))).toEqual(shared);
        expect(
          pair.editors.map((peer) => ({
            undo: peer.can({ type: 'undo' }),
            redo: peer.can({ type: 'redo' }),
          }))
        ).toEqual(history);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        release();
        pair.destroy();
      }
    });
  test(`cancellable replacement still commits once and retains history ${clientIds}`, async () => {
    const decoder = deferredDecoder();
    const pair = await createTwoEditorFixture({
      document: await documentWithImage(),
      clientIds,
      editor: { imageDecodePort: decoder.port },
    });
    const editor = pair.editors[0]!;
    try {
      await pair.settle();
      const id = drawings(editor)[0]!;
      expect(editor.surface!.selectDrawing(id, editor.surface!.session.paragraphIds()[0]!)).toBe(
        true
      );
      const before = await pair.savedDigests();
      const originalMedia = media(editor);
      const result = await editor.executeImageCommand(
        { type: 'replaceImage', data: replacement, mime: 'image/gif', drawingNodeId: id },
        { signal: new AbortController().signal }
      );
      expect(result.ok).toBe(true);
      await pair.settle();
      const after = await pair.savedDigests();
      expect(after[0]).toEqual(after[1]);
      expect(after[0]).not.toEqual(before[0]);
      const changedMedia = media(editor);
      expect(changedMedia).not.toEqual(originalMedia);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(await pair.savedDigests()).toEqual(before);
      expect(media(editor)).toEqual(originalMedia);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      expect(await pair.savedDigests()).toEqual(after);
      expect(media(editor)).toEqual(changedMedia);
    } finally {
      pair.destroy();
    }
  });
}

test('a cancelled image command never starts decoding', async () => {
  let calls = 0;
  const pair = await createTwoEditorFixture({
    document: documentBytes,
    editor: {
      imageDecodePort: {
        async decode() {
          calls++;
          return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
        },
      },
    },
  });
  try {
    const controller = new AbortController();
    controller.abort();
    const before = await pair.savedDigests();
    expect(
      (
        await pair.editors[0]!.executeImageCommand(
          {
            type: 'insertImage',
            data: galleryPng,
            mime: 'image/png',
            widthPoints: 36,
            heightPoints: 36,
          },
          { signal: controller.signal }
        )
      ).ok
    ).toBe(false);
    expect(calls).toBe(0);
    expect(await pair.savedDigests()).toEqual(before);
  } finally {
    pair.destroy();
  }
});
