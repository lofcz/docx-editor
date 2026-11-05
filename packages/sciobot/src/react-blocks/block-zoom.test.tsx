import { expect, test } from 'bun:test';
import { act, useState } from 'react';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { galleryDocument } from '../../test/support/gallery';
import { readGallery, updateGallery } from '../blocks/gallery';
import { mountReactBlockHost, type ReactBlockViewProps } from './host';

const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
function Widget({ fragment }: ReactBlockViewProps) {
  const [draft, setDraft] = useState('draft');
  return (
    <input
      value={draft}
      data-content-scale={fragment.contentScale}
      onChange={(event) => setDraft(event.target.value)}
    />
  );
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`block zoom retains native flow and peer geometry without shared writes ${clientIds.join('/')}`, async () => {
    const previous = environment.IS_REACT_ACT_ENVIRONMENT;
    environment.IS_REACT_ACT_ENVIRONMENT = true;
    const pair = await createTwoEditorFixture({ document: galleryDocument(), clientIds });
    const hosts: ReturnType<typeof mountReactBlockHost>[] = [];
    const zooms = [1, 1];
    const settle = async () => {
      await act(async () => {
        await pair.settle();
        await new Promise((resolve) => setTimeout(resolve, 10));
      });
    };
    const roots = () =>
      pair.containers.map(
        (container) => container.querySelector<HTMLElement>('[data-sciobot-block-measure]')!
      );
    const tails = () =>
      pair.editors.map((editor) => {
        const session = editor.surface!.session;
        const id = session
          .paragraphIds()
          .find((id) => paragraphTextOf(session.part(), id)?.startsWith('Tail'))!;
        return { id, point: caretAt(editor.surface!.layout(), { paragraphId: id, offset: 0 })! };
      });
    try {
      await act(async () => {
        pair.editors.forEach((editor, index) =>
          hosts.push(
            mountReactBlockHost(
              editor,
              new Map([['image_gallery', { view: Widget, reserveHeight: true }]]),
              // Simulate physical browser height; the browser suite verifies actual row geometry.
              { measure: () => 200 * zooms[index]! }
            )
          )
        );
      });
      await settle();
      const [alice, bob] = pair.editors;
      const initial = roots().map((root) => ({
        root,
        width: root.style.width,
        maxHeight: root.style.maxHeight,
      }));
      const initialTail = tails().map((tail) => tail.point.y);
      const revision = alice!.surface!.session.packageRevision();
      const input = initial[0]!.root.querySelector('input')!;
      input.focus();
      for (const zoom of [0.75, 1.5, 1]) {
        await act(async () => {
          zooms[0] = zoom;
          expect(alice!.setZoom(zoom).ok).toBe(true);
        });
        await settle();
        const [first, second] = roots();
        expect(first).toBe(initial[0]!.root);
        expect(first!.style.width).toBe(initial[0]!.width);
        expect(first!.style.maxHeight).toBe(initial[0]!.maxHeight);
        expect(first!.style.transform).toBe(`scale(${zoom})`);
        expect(second!.style.transform).toBe('scale(1)');
        expect(input.dataset.contentScale).toBe(String(zoom));
        expect(input.ownerDocument.activeElement).toBe(input);
        expect(tails().map((tail) => tail.point.y)).toEqual(initialTail);
        expect(alice!.surface!.session.packageRevision()).toBe(revision);
      }
      await act(async () => {
        zooms[0] = 0.75;
        alice!.setZoom(0.75);
      });
      pair.pause();
      const gallery = readGallery(alice!.surface!.session.currentPackage(), 'gallery')!;
      const tail = tails()[1]!;
      await act(async () => {
        expect(
          updateGallery(alice!, {
            blockId: 'gallery',
            change: {
              type: 'caption',
              imageId: gallery.images[0]!.id,
              expectedCaption: 'First image',
              caption: 'Alice caption',
            },
          }).success
        ).toBe(true);
        bob!.surface!.setSelection({
          anchor: { paragraphId: tail.id, offset: 4 },
          head: { paragraphId: tail.id, offset: 4 },
        });
        bob!.surface!.insertPlainText(' peer');
        bob!.surface!.flushPendingInput();
        pair.reconnect(0);
      });
      await settle();
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      )!;
      expect(remote).toBeDefined();
      const localPoint = caretAt(bob!.surface!.layout(), {
        paragraphId: remote.head.nodeId,
        offset: remote.head.offset,
      })!;
      const remotePoint = caretAt(alice!.surface!.layout(), {
        paragraphId: remote.head.nodeId,
        offset: remote.head.offset,
      })!;
      expect({ x: remotePoint.x, y: remotePoint.y }).toEqual({ x: localPoint.x, y: localPoint.y });
      await act(async () => {
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      });
      await settle();
      expect(
        readGallery(bob!.surface!.session.currentPackage(), 'gallery')!.images[0]!.attributes
          .caption
      ).toBe('First image');
      expect(paragraphTextOf(bob!.surface!.session.part(), tail.id)).toBe('Tail peer');
      await act(async () => {
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      });
      await settle();
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
      await act(async () => {
        const saved = await alice!.save();
        await alice!.load(new Uint8Array(saved));
      });
      await settle();
      expect(roots()[0]).toBeDefined();
      expect(roots()[0]).not.toBe(initial[0]!.root);
      expect(roots()[0]!.style.width).toBe(initial[0]!.width);
    } finally {
      await act(async () => {
        hosts.forEach((host) => host.dispose());
        await Promise.resolve();
        pair.destroy();
      });
      environment.IS_REACT_ACT_ENVIRONMENT = previous;
    }
  });
