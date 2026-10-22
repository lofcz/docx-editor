// Internal compatibility test. The demo uses Apache packages; this test also
// evaluates the upstream collaboration runtime under the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { act } from 'react';
import * as Y from 'yjs';
import { Awareness, applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import { createDocxEditor, createBrowserAutomationHost } from '@docx-editor.dev/core/editor';
import { readOoxmlPackage, semanticDigest, paragraphTextOf } from '@docx-editor.dev/core/store';
import { createDocumentCollaboration } from '../../../packages/pro/src/collaboration/document-session.ts';
import { fixture } from '../src/document';
import { REVIEW, createBlockBridge } from '../src/bridge';
import { mountReactBlocks } from '../src/react-block-layer';
import { ActivityBlock, PhaseBlock, QuizBlock } from '../src/blocks';

test('React block hosts retain two real replicas through concurrent edits, reconnect, review history, and deletion', async () => {
  const actEnvironment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean };
  const previousActEnvironment = actEnvironment.IS_REACT_ACT_ENVIRONMENT;
  actEnvironment.IS_REACT_ACT_ENVIRONMENT = true;
  const docs = [new Y.Doc(), new Y.Doc()];
  const awareness = docs.map((doc) => new Awareness(doc));
  const first = await createDocumentCollaboration({
    ydoc: docs[0]!,
    awareness: awareness[0]!,
    documentId: 'react-block-proof',
    identity: { actorId: 'alice', name: 'Alice' },
    bootstrap: { kind: 'create', document: fixture() },
  });
  Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!), 'relay');
  const second = await createDocumentCollaboration({
    ydoc: docs[1]!,
    awareness: awareness[1]!,
    documentId: 'react-block-proof',
    identity: { actorId: 'bob', name: 'Bob' },
    bootstrap: { kind: 'join', timeoutMs: 5000 },
  });
  const rooms = [first, second];
  let connected = true;
  docs.forEach((doc, i) =>
    doc.on('update', (update: Uint8Array, origin: unknown) => {
      if (connected && origin !== 'relay') Y.applyUpdate(docs[1 - i]!, update, 'relay');
    })
  );
  awareness.forEach((source, i) =>
    source.on(
      'update',
      (
        { added, updated, removed }: { added: number[]; updated: number[]; removed: number[] },
        origin: unknown
      ) => {
        if (connected && origin !== 'relay')
          applyAwarenessUpdate(
            awareness[1 - i]!,
            encodeAwarenessUpdate(source, [...added, ...updated, ...removed]),
            'relay'
          );
      }
    )
  );
  const containers = rooms.map(() => {
    const node = document.createElement('div');
    document.body.append(node);
    return node;
  });
  const editors = rooms.map((room, i) =>
    createDocxEditor({
      container: containers[i]!,
      document: room.document,
      author: i ? 'Bob' : 'Alice',
      modules: [REVIEW, { id: 'collaboration', collaboration: { session: room.session } }],
    })
  );
  const bridges = editors.map(createBlockBridge);
  const stop: (() => void)[] = [];
  await act(async () => {
    editors.forEach((editor, i) =>
      stop.push(
        mountReactBlocks(
          editor,
          { phase: PhaseBlock, activity: ActivityBlock, quiz: QuizBlock },
          bridges[i]!
        )
      )
    );
  });
  const paragraph = (i: number, text = 'Plan an experiment.') =>
    [
      ...containers[i]!.querySelectorAll<HTMLElement>('[data-native-slot] [data-paragraph-id]'),
    ].find((p) => p.textContent?.includes(text))!;
  const read = (i: number) => {
    const node = paragraph(i);
    return paragraphTextOf(editors[i]!.surface!.session.part(), node.dataset.paragraphId!)!;
  };
  const flush = () => rooms.forEach((room) => room.session.flushPendingJournals());
  const type = (i: number, text: string) => {
    const node = paragraph(i),
      id = node.dataset.paragraphId!;
    editors[i]!.surface!.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    editors[i]!.surface!.insertPlainText(text);
    flush();
  };
  try {
    expect(paragraph(0).closest('[data-react-block="activity"]')).not.toBeNull();
    await act(async () => {
      type(0, 'Alice: ');
    });
    expect(read(1)).toBe(read(0));
    expect(second.session.remoteSelections().length).toBeGreaterThan(0);
    connected = false;
    await act(async () => {
      type(0, 'A ');
      type(1, 'B ');
    });
    connected = true;
    await act(async () => {
      Y.applyUpdate(docs[0]!, Y.encodeStateAsUpdate(docs[1]!), 'relay');
      Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!), 'relay');
      flush();
    });
    expect(read(0)).toBe(read(1));
    expect(read(0)).toContain('A ');
    expect(read(0)).toContain('B ');
    const combined = read(0);
    await act(async () => {
      expect(editors[0]!.exec({ type: 'undo' }).ok).toBe(true);
      flush();
    });
    expect(read(0)).toBe(read(1));
    await act(async () => {
      expect(editors[0]!.exec({ type: 'redo' }).ok).toBe(true);
      flush();
    });
    expect(read(0)).toBe(combined);
    expect(read(1)).toBe(combined);
    const target = paragraph(1, 'Activity step 3.');
    const targetId = target.dataset.paragraphId!;
    await act(async () => {
      editors[1]!.surface!.setSelection({
        anchor: { paragraphId: targetId, offset: 0 },
        head: { paragraphId: targetId, offset: 0 },
      });
      editors[1]!.surface!.insertPlainText('Peer is editing here. ');
      flush();
      const host = createBrowserAutomationHost(editors[0]!);
      const one = host.execute({ operations: [{ op: 'getDocument' }] }).results[0]!;
      if (one.status !== 'ok' || one.value.kind !== 'handle') throw new Error('No document');
      const body = host.execute({ operations: [{ op: 'getBody', document: one.value.handle }] })
        .results[0]!;
      if (body.status !== 'ok' || body.value.kind !== 'handle') throw new Error('No body');
      const matches = host.execute({
        operations: [
          { op: 'search', scope: { body: body.value.handle }, text: 'Activity step 3.' },
        ],
      }).results[0]!;
      if (matches.status !== 'ok' || matches.value.kind !== 'spans') throw new Error('No match');
      expect(
        host.execute({
          operations: [
            { op: 'deleteParagraph', paragraph: matches.value.spans[0]!.start.paragraph },
          ],
        }).ok
      ).toBe(true);
      host.dispose();
      flush();
    });
    expect(paragraph(0, 'Activity step 3.')).toBeUndefined();
    expect(paragraph(1, 'Activity step 3.')).toBeUndefined();
    const digests = [];
    for (const editor of editors) {
      const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!opened.ok) throw new Error(opened.reason);
      digests.push(semanticDigest(opened.package.parts.values()));
    }
    expect(digests[0]).toEqual(digests[1]);
  } finally {
    await act(async () => {
      stop.forEach((stop) => stop());
      await Promise.resolve();
    });
    bridges.forEach((bridge) => bridge.destroy());
    editors.forEach((editor) => editor.destroy());
    rooms.forEach((room) => room.destroy());
    awareness.forEach((a) => a.destroy());
    docs.forEach((doc) => doc.destroy());
    containers.forEach((node) => node.remove());
    actEnvironment.IS_REACT_ACT_ENVIRONMENT = previousActEnvironment;
  }
}, 30000);
