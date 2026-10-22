import { expect, test } from 'bun:test';
import { act } from 'react';
import { caretAt } from '@docx-editor.dev/core/layout';
import { paragraphTextOf } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { nativeSlotDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { mountReactBlockHost, type ReactBlockViewProps } from './host';

const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
function Container({ block, children }: ReactBlockViewProps) {
  return (
    <div data-container={block.id} style={{ display: 'contents' }}>
      {children}
    </div>
  );
}
for (const receiver of [0, 1] as const)
  test(`native nested slots converge through reconnect, peer-preserving history, remote selection, deletion, and save; receiver ${receiver}`, async () => {
    const previous = environment.IS_REACT_ACT_ENVIRONMENT;
    environment.IS_REACT_ACT_ENVIRONMENT = true;
    const pair = await createTwoEditorFixture({
      document: nativeSlotDocument(),
      clientIds: [73, 41],
    });
    const hosts: ReturnType<typeof mountReactBlockHost>[] = [];
    const [alice] = pair.editors;
    const surfaces = pair.editors.map((editor) => editor.surface!);
    const ids = (peer: number, text: string) => {
      const session = surfaces[peer]!.session;
      const id = session.paragraphIds().find((id) => paragraphTextOf(session.part(), id) === text);
      if (!id) throw new Error(`Missing native paragraph: ${text}`);
      return id;
    };
    const inner = ids(0, 'Inner text'),
      other = ids(1, 'Other cell');
    const append = (peer: number, paragraphId: string, text: string) => {
      const surface = surfaces[peer]!,
        offset = paragraphTextOf(surface.session.part(), paragraphId)!.length;
      surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
      surface.insertPlainText(text);
      surface.flushPendingInput();
    };
    const rendered = (id: string, expected: string) => {
      for (const container of pair.containers) {
        const paragraphs = container.querySelectorAll(
          `.docx-paragraph-fragment[data-paragraph-id="${id}"]`
        );
        expect(paragraphs).toHaveLength(1);
        expect(paragraphs[0]!.textContent).toBe(expected);
        expect(paragraphs[0]!.closest('[data-sciobot-native-slot]') !== null).toBe(true);
        expect(container.querySelectorAll('.docx-paragraph-fragment')).toHaveLength(12);
        expect(container.querySelectorAll('[data-sciobot-native-slot]')).toHaveLength(7);
      }
    };
    try {
      await act(async () => {
        for (const editor of pair.editors)
          hosts.push(
            mountReactBlockHost(
              editor,
              new Map([['native_container', { view: Container, nativeContent: true }]])
            )
          );
      });
      pair.pause();
      await act(async () => {
        append(0, inner, ' Alice');
        append(1, other, ' Bob');
        pair.reconnect(receiver);
        await pair.settle();
      });
      rendered(inner, 'Inner text Alice');
      rendered(other, 'Other cell Bob');
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      );
      expect(Boolean(remote)).toBe(true);
      expect(remote!.head.nodeId).toBe(other);
      const geometry = surfaces.map((surface) =>
        caretAt(surface.layout(), {
          paragraphId: remote!.head.nodeId,
          offset: remote!.head.offset,
        })
      );
      expect(geometry.every((caret) => caret !== null)).toBe(true);
      expect({ x: geometry[0]!.x, y: geometry[0]!.y }).toEqual({
        x: geometry[1]!.x,
        y: geometry[1]!.y,
      });
      await act(async () => {
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
      });
      rendered(inner, 'Inner text');
      rendered(other, 'Other cell Bob');
      await act(async () => {
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
      });
      rendered(inner, 'Inner text Alice');
      rendered(other, 'Other cell Bob');
      const rowCell = ids(0, 'Row cell text'),
        rowOther = ids(1, 'Row other text');
      pair.pause();
      await act(async () => {
        append(0, rowCell, ' Alice');
        append(1, rowOther, ' Bob');
        pair.reconnect(receiver);
        await pair.settle();
      });
      rendered(rowCell, 'Row cell text Alice');
      rendered(rowOther, 'Row other text Bob');
      await act(async () => {
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
      });
      rendered(rowCell, 'Row cell text');
      rendered(rowOther, 'Row other text Bob');
      await act(async () => {
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
      });
      rendered(rowCell, 'Row cell text Alice');
      rendered(rowOther, 'Row other text Bob');
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      pair.pause();
      await act(async () => {
        append(1, inner, ' Editing');
        const outer = buildBlockIndex(surfaces[0]!.session.currentPackage()).byId.get('outer')!;
        expect(
          surfaces[0]!.applyAutomationOps(() => [
            { op: 'removeContentControl', controlId: outer.controlId, keepContent: false },
          ]).rejected
        ).toBe(false);
        pair.reconnect(receiver);
        await pair.settle();
        await Promise.resolve();
      });
      for (const container of pair.containers) {
        expect(container.querySelectorAll('[data-container="outer"]')).toHaveLength(0);
        expect(container.querySelectorAll('[data-container="inner"]')).toHaveLength(0);
        expect(container.querySelectorAll('.docx-paragraph-fragment')).toHaveLength(4);
        expect(container.querySelectorAll('[data-sciobot-native-slot]')).toHaveLength(2);
        expect(container.querySelector('[data-container="cell"]')?.textContent).toBe(
          'Row cell text Alice'
        );
      }
      const deleted = await pair.savedDigests();
      expect(deleted[0]).toEqual(deleted[1]);
      pair.pause();
      await act(async () => {
        append(1, rowCell, ' Still editing');
        const row = buildBlockIndex(surfaces[0]!.session.currentPackage()).byId.get('row')!;
        expect(
          surfaces[0]!.applyAutomationOps(() => [
            { op: 'removeContentControl', controlId: row.controlId, keepContent: false },
          ]).rejected
        ).toBe(false);
        pair.reconnect(receiver);
        await pair.settle();
        await Promise.resolve();
      });
      for (const container of pair.containers) {
        expect(container.querySelectorAll('[data-sciobot-native-slot]')).toHaveLength(0);
        expect(
          [...container.querySelectorAll('.docx-paragraph-fragment')].map((p) => p.textContent)
        ).toEqual(['Intro', 'Tail']);
      }
      const rowDeleted = await pair.savedDigests();
      expect(rowDeleted[0]).toEqual(rowDeleted[1]);
    } finally {
      await act(async () => {
        hosts.forEach((host) => host.dispose());
        await Promise.resolve();
        pair.destroy();
      });
      environment.IS_REACT_ACT_ENVIRONMENT = previous;
    }
  }, 30_000);
