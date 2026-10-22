import { expect, test } from 'bun:test';
import { act } from 'react';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { hostControl, hostParagraph } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { writeBlockPayload } from '../blocks/payload-part';
import { mountReactBlockHost, type ReactBlockViewProps } from './host';
import { lastBodyParagraphIn } from './measurement';

const environment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
const ROW = 60;

/** A widget whose intrinsic height follows its payload: `rows` × 60 px, or 1 px when empty. */
function Tall({ block }: ReactBlockViewProps) {
  const rows = Number(block.payload.attributes.rows ?? 0);
  return <div data-height={rows > 0 ? rows * ROW : 1}>rows {rows}</div>;
}
const measure = (element: HTMLElement) =>
  Number(element.querySelector('[data-height]')?.getAttribute('data-height') ?? 0);

function tallDocument(): Uint8Array {
  const read = readOoxmlPackage(
    docxFixture(
      hostParagraph('Intro') +
        hostControl('tall-1', hostParagraph('Tall fallback')) +
        hostParagraph('Tail') +
        hostParagraph('After')
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const written = withPayloadRecord(
    read.package,
    { id: 'tall-1', type: 'test_tall', version: 1, attributes: { rows: 0 } },
    'Tall'
  );
  if (!written.ok) throw new Error(written.reason);
  return writeOoxmlPackage(written.pkg);
}

const idOf = (editor: DocxEditorInstance, text: string) => {
  const session = editor.surface!.session;
  const id = session.paragraphIds().find((id) => paragraphTextOf(session.part(), id) === text);
  if (!id) throw new Error(`Missing paragraph ${text}`);
  return id;
};
const yAt = (editor: DocxEditorInstance, paragraphId: string, offset = 0) => {
  editor.surface!.flushPendingInput();
  return caretAt(editor.surface!.layout(), { paragraphId, offset })!.y;
};
const tallBox = (editor: DocxEditorInstance) => {
  const block = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('tall-1');
  if (!block) return null;
  const boundary = editor
    .surface!.layout()
    .contentControls?.find((entry) => entry.id === block.controlId);
  return boundary?.fragments.map((fragment) => ({ ...fragment.box })) ?? null;
};
function setRows(editor: DocxEditorInstance, rows: number) {
  const block = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('tall-1')!;
  const result = writeBlockPayload(editor, {
    controlId: block.controlId,
    payload: { ...block.payload, attributes: { rows } },
    label: 'Tall',
    fallbackOps: [],
  });
  expect(result.success).toBe(true);
}
const scaleOf = (editor: DocxEditorInstance) => {
  let scale = 0;
  const stop = editor.surface!.observePaint({ afterPaint: (frame) => (scale = frame.scale) });
  stop();
  return scale;
};
const expectedExtra = (rows: number, nativeHeight: number, scale: number) =>
  Math.max(0, Math.ceil(((rows * ROW) / scale - nativeHeight) * 4 - 1e-6) / 4);

test('the last body paragraph of a control skips nested controls and refuses a trailing table', () => {
  const read = readOoxmlPackage(
    docxFixture(
      hostControl('a', hostParagraph('One') + hostControl('b', hostParagraph('Two'))) +
        hostControl(
          'c',
          hostParagraph('Three') +
            '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
        )
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const part = read.package.parts.get('/word/document.xml')!;
  const controls = part.root.children
    .flatMap((node) => (node.kind === 'body' ? node.children : []))
    .filter((node) => node.kind === 'contentControl');
  const lastA = lastBodyParagraphIn(controls[0]!)!;
  expect(paragraphTextOf(part, lastA)).toBe('Two');
  expect(lastBodyParagraphIn(controls[1]!)).toBeNull();
});

test('an expanding widget moves following native text once, and one undo reverses only the user edit', async () => {
  const previous = environment.IS_REACT_ACT_ENVIRONMENT;
  environment.IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: tallDocument() });
  let host: ReturnType<typeof mountReactBlockHost> | undefined;
  const tail = idOf(editor, 'Tail');
  const tailYs: number[] = [];
  const stop = editor.surface!.observePaint({
    afterPaint: (frame) => tailYs.push(caretAt(frame.layout, { paragraphId: tail, offset: 0 })!.y),
  });
  const settle = async () => {
    for (let round = 0; round < 3; round += 1)
      await act(async () => {
        await new Promise((resolve) => setTimeout(resolve, 15));
      });
  };
  try {
    const plain = new Uint8Array(await editor.save());
    const base = yAt(editor, tail);
    const afterBase = yAt(editor, idOf(editor, 'After'));
    await act(async () => {
      host = mountReactBlockHost(
        editor,
        new Map([['test_tall', { view: Tall, reserveHeight: true }]]),
        {
          measure,
        }
      );
    });
    await settle();
    // An empty widget is shorter than its fallback: nothing is reserved.
    expect(yAt(editor, tail)).toBe(base);
    expect(container.querySelectorAll('[data-sciobot-block-measure]')).toHaveLength(1);
    const native = tallBox(editor)!;
    expect(native).toHaveLength(1);
    const scale = scaleOf(editor);
    expect(scale).toBeGreaterThan(0);
    expect(editor.surface!.state().canUndo).toBe(false);

    tailYs.length = 0;
    setRows(editor, 3);
    await settle();
    const extra = expectedExtra(3, native[0]!.height, scale);
    expect(extra).toBeGreaterThan(0);
    const expanded = yAt(editor, tail);
    expect(expanded).toBeCloseTo(base + extra, 6);
    expect(yAt(editor, idOf(editor, 'After')) - afterBase).toBeCloseTo(extra, 6);
    // The native box never includes the reservation, so the measurement cannot feed back.
    expect(tallBox(editor)).toEqual(native);
    // Following text moved once: from its old position to its final one, never back and forth.
    const distinct = tailYs.filter((value, index) => index === 0 || value !== tailYs[index - 1]);
    expect(distinct.length).toBeLessThanOrEqual(2);
    expect(distinct.at(-1)!).toBeCloseTo(expanded, 6);
    if (distinct.length === 2) expect(distinct[0]!).toBe(base);
    // Stable: further turns publish nothing.
    const frames = tailYs.length;
    await settle();
    expect(tailYs.length).toBe(frames);

    // One user edit, one undo step; measurement adds none.
    expect(editor.surface!.state().canUndo).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    await settle();
    expect(editor.surface!.state().canUndo).toBe(false);
    expect(yAt(editor, tail)).toBe(base);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    await settle();
    expect(editor.surface!.state().canRedo).toBe(false);
    expect(yAt(editor, tail)).toBeCloseTo(expanded, 6);

    // Saved files never carry a measurement.
    const withReservation = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!withReservation.ok) throw new Error(withReservation.reason);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    await settle();
    expect(new Uint8Array(await editor.save())).toEqual(plain);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    await settle();
    const reopenedContainer = document.createElement('div');
    document.body.append(reopenedContainer);
    const reopened = createDocxEditor({
      container: reopenedContainer,
      document: writeOoxmlPackage(withReservation.package),
    });
    try {
      // Without a host, the reopened document lays out at its native positions.
      expect(yAt(reopened, idOf(reopened, 'Tail'))).toBe(base);
    } finally {
      reopened.destroy();
      reopenedContainer.remove();
    }

    // Disposal releases the reservation.
    await act(async () => host!.dispose());
    host = undefined;
    await settle();
    expect(yAt(editor, tail)).toBe(base);
  } finally {
    stop();
    host?.dispose();
    editor.destroy();
    container.remove();
    environment.IS_REACT_ACT_ENVIRONMENT = previous;
  }
});

for (const clientIds of [
  [73, 41],
  [41, 73],
] as const)
  for (const receiver of [0, 1] as const)
    test(`each session reserves its own measured height through offline edits, history, remote carets, and deletion; clients ${clientIds.join('/')}, receiver ${receiver}`, async () => {
      const previous = environment.IS_REACT_ACT_ENVIRONMENT;
      environment.IS_REACT_ACT_ENVIRONMENT = true;
      const pair = await createTwoEditorFixture({ document: tallDocument(), clientIds });
      const hosts: ReturnType<typeof mountReactBlockHost>[] = [];
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const settle = async () => {
        for (let round = 0; round < 3; round += 1)
          await act(async () => {
            await pair.settle();
            await new Promise((resolve) => setTimeout(resolve, 10));
          });
      };
      const healthy = () => {
        for (const room of pair.rooms) expect(room.session.status()).not.toBe('error');
      };
      try {
        const base = pair.editors.map((editor) => yAt(editor, idOf(editor, 'Tail')));
        expect(base[0]).toBe(base[1]!);
        await act(async () => {
          for (const editor of pair.editors)
            hosts.push(
              mountReactBlockHost(
                editor,
                new Map([['test_tall', { view: Tall, reserveHeight: true }]]),
                { measure }
              )
            );
        });
        await settle();
        const native = tallBox(alice)!;
        const extra = expectedExtra(3, native[0]!.height, scaleOf(alice));

        pair.pause();
        setRows(alice, 3);
        const tailBob = idOf(bob, 'Tail');
        bob.surface!.setSelection({
          anchor: { paragraphId: tailBob, offset: 4 },
          head: { paragraphId: tailBob, offset: 4 },
        });
        bob.surface!.insertPlainText(' Bob');
        bob.surface!.flushPendingInput();
        await settle();
        // Offline: only Alice's own view reserves.
        expect(yAt(alice, idOf(alice, 'Tail'))).toBeCloseTo(base[0]! + extra, 6);
        expect(yAt(bob, idOf(bob, 'Tail Bob'))).toBe(base[1]!);

        pair.reconnect(receiver);
        await settle();
        healthy();
        for (const editor of pair.editors) {
          expect(yAt(editor, idOf(editor, 'Tail Bob'))).toBeCloseTo(base[0]! + extra, 6);
          expect(tallBox(editor)).toEqual(native);
        }
        const [savedAlice, savedBob] = await pair.savedDigests();
        expect(savedAlice).toEqual(savedBob!);

        // Bob's caret paints at the reserved position in both sessions.
        const bobTail = idOf(bob, 'Tail Bob');
        bob.surface!.setSelection({
          anchor: { paragraphId: bobTail, offset: 2 },
          head: { paragraphId: bobTail, offset: 2 },
        });
        await settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        )!;
        expect(remote).toBeDefined();
        const remoteY = caretAt(alice.surface!.layout(), {
          paragraphId: remote.head.nodeId,
          offset: remote.head.offset,
        })!.y;
        expect(remoteY).toBeCloseTo(yAt(bob, bobTail, 2), 6);
        expect(remoteY).toBeCloseTo(base[0]! + extra, 6);

        // Alice's undo shrinks both widgets and keeps Bob's text; redo expands both again.
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await settle();
        for (const editor of pair.editors)
          expect(yAt(editor, idOf(editor, 'Tail Bob'))).toBe(base[0]!);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await settle();
        for (const editor of pair.editors)
          expect(yAt(editor, idOf(editor, 'Tail Bob'))).toBeCloseTo(base[0]! + extra, 6);
        healthy();

        // Bob deletes the block while Alice expands it further offline.
        pair.pause();
        const block = buildBlockIndex(bob.surface!.session.currentPackage()).byId.get('tall-1')!;
        expect(
          bob.surface!.applyAutomationOps(() => [
            { op: 'removeContentControl', controlId: block.controlId, keepContent: false },
          ]).rejected
        ).toBe(false);
        setRows(alice, 5);
        pair.reconnect(receiver);
        await settle();
        healthy();
        const remaining = pair.editors.map((editor) => ({
          block: buildBlockIndex(editor.surface!.session.currentPackage()).byId.has('tall-1'),
          roots: editor
            .surface!.hostElement()
            .ownerDocument.querySelectorAll('[data-sciobot-block-measure]').length,
        }));
        // Deletion wins over the concurrent payload write: no widget and no reservation remain.
        expect(remaining).toEqual([
          { block: false, roots: 0 },
          { block: false, roots: 0 },
        ]);
        const tails = pair.editors.map((editor) => yAt(editor, idOf(editor, 'Tail Bob')));
        expect(tails[0]).toBeCloseTo(tails[1]!, 6);
        for (const editor of pair.editors) {
          expect(tallBox(editor)).toBeNull();
          // No widget remains, so nothing stays reserved.
          const fresh = document.createElement('div');
          document.body.append(fresh);
          const reopened = createDocxEditor({
            container: fresh,
            document: new Uint8Array(await editor.save()),
          });
          try {
            expect(yAt(editor, idOf(editor, 'Tail Bob'))).toBe(
              yAt(reopened, idOf(reopened, 'Tail Bob'))
            );
          } finally {
            reopened.destroy();
            fresh.remove();
          }
        }
        const digests = await pair.savedDigests();
        expect(digests[0]).toEqual(digests[1]!);
        // Reservations never reach the saved package.
        for (const editor of pair.editors) {
          const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!saved.ok) throw new Error(saved.reason);
          expect(semanticDigest(saved.package.parts.values())).toEqual(digests[0]!);
        }
      } finally {
        await act(async () => hosts.forEach((host) => host.dispose()));
        pair.destroy();
        environment.IS_REACT_ACT_ENVIRONMENT = previous;
      }
    }, 60_000);
