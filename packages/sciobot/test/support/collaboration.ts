// Test-only compatibility fixture. The evaluated transport uses the EigenPal Pro License.
// Production adapter modules must never import this file.
import * as Y from 'yjs';
import { Awareness, applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import {
  createDocxEditor,
  type DocxEditorConfig,
  type DocxEditorInstance,
} from '@docx-editor.dev/core/editor';
import { readOoxmlPackage, semanticDigest } from '@docx-editor.dev/core/store';
import { createDocumentCollaboration } from '../../../pro/src/collaboration/document-session';
import { attachOwnedReviewHistory } from '../../src/collaboration';
import { prepareReviewGroupDocument } from '../../src/native/prepare-review-groups';

export interface TwoEditorOptions {
  readonly document: Uint8Array;
  readonly documentId?: string;
  readonly hostDocument?: Document;
  readonly editor?: Omit<DocxEditorConfig, 'document' | 'container'>;
  /** Fixed before any shared writes, for deterministic ordering regressions. */
  readonly clientIds?: readonly [number, number];
  /** Application-owned attribute protocol, installed before room bootstrap. */
  readonly ownedAttributes?: boolean;
  /** Reopen the complete shared room instead of creating it from DOCX. */
  readonly roomSnapshot?: Uint8Array;
}

/** Two real editor sessions with a controllable in-memory transport. */
export async function createTwoEditorFixture(options: TwoEditorOptions) {
  const documents = [new Y.Doc(), new Y.Doc()];
  if (options.clientIds)
    documents.forEach((document, index) => {
      document.clientID = options.clientIds![index]!;
    });
  const ownership = options.ownedAttributes
    ? documents.map((doc) =>
        attachOwnedReviewHistory(doc, { documentId: options.documentId ?? 'sciobot-parity' })
      )
    : [];
  const awareness = documents.map((doc) => new Awareness(doc));
  const rooms: Awaited<ReturnType<typeof createDocumentCollaboration>>[] = [];
  const editors: DocxEditorInstance[] = [];
  const containers: HTMLElement[] = [];
  let connected = true;
  let disposed = false;
  const flush = () => rooms.forEach((room) => room.session.flushPendingJournals());
  const destroy = () => {
    if (disposed) return;
    disposed = true;
    connected = false;
    editors.forEach((editor) => editor.destroy());
    rooms.forEach((room) => room.destroy());
    awareness.forEach((source) => source.destroy());
    ownership.forEach((handle) => handle.dispose());
    documents.forEach((doc) => doc.destroy());
    containers.forEach((node) => node.remove());
  };
  try {
    const documentId = options.documentId ?? 'sciobot-parity';
    const prepared = options.roomSnapshot ? null : prepareReviewGroupDocument(options.document);
    if (prepared && !prepared.success) throw new Error(prepared.error);
    if (options.roomSnapshot) Y.applyUpdate(documents[0]!, options.roomSnapshot, 'relay');
    rooms.push(
      await createDocumentCollaboration({
        ydoc: documents[0]!,
        awareness: awareness[0]!,
        documentId,
        identity: { actorId: 'alice', name: 'Alice' },
        bootstrap: options.roomSnapshot
          ? { kind: 'join', timeoutMs: 5000 }
          : { kind: 'create', document: prepared?.success ? prepared.document : options.document },
      })
    );
    Y.applyUpdate(documents[1]!, Y.encodeStateAsUpdate(documents[0]!), 'relay');
    rooms.push(
      await createDocumentCollaboration({
        ydoc: documents[1]!,
        awareness: awareness[1]!,
        documentId,
        identity: { actorId: 'bob', name: 'Bob' },
        bootstrap: { kind: 'join', timeoutMs: 5000 },
      })
    );
    // Both peers can announce application protocols before bootstrap. Complete the
    // initial bidirectional sync before relaying incremental updates from either peer.
    Y.applyUpdate(documents[0]!, Y.encodeStateAsUpdate(documents[1]!), 'relay');
    Y.applyUpdate(documents[1]!, Y.encodeStateAsUpdate(documents[0]!), 'relay');
    documents.forEach((doc, index) =>
      doc.on('update', (update, origin) => {
        if (connected && origin !== 'relay') Y.applyUpdate(documents[1 - index]!, update, 'relay');
      })
    );
    awareness.forEach((source, index) =>
      source.on(
        'update',
        (change: { added: number[]; updated: number[]; removed: number[] }, origin: unknown) => {
          if (connected && origin !== 'relay')
            applyAwarenessUpdate(
              awareness[1 - index]!,
              encodeAwarenessUpdate(source, [
                ...change.added,
                ...change.updated,
                ...change.removed,
              ]),
              'relay'
            );
        }
      )
    );
    for (const [index, room] of rooms.entries()) {
      const container = (options.hostDocument ?? document).createElement('div');
      container.ownerDocument.body.append(container);
      containers.push(container);
      editors.push(
        createDocxEditor({
          ...options.editor,
          container,
          document: room.document,
          author: index === 0 ? 'Alice' : 'Bob',
          modules: [
            ...(options.editor?.modules ?? []),
            {
              id: 'collaboration',
              collaboration: {
                session: ownership[index]?.bindSession(room.session) ?? room.session,
              },
            },
          ],
        })
      );
    }
    return {
      editors,
      rooms,
      containers,
      documents,
      awareness,
      ownership,
      flush,
      async settle() {
        flush();
        // Remote publication updates canonical state before its queued surface paint.
        const hostWindow = containers[0]!.ownerDocument.defaultView!;
        await new Promise<void>((done) => hostWindow.requestAnimationFrame(() => done()));
        await new Promise<void>((done) => hostWindow.requestAnimationFrame(() => done()));
        flush();
      },
      pause() {
        flush();
        connected = false;
      },
      reconnect(first: 0 | 1 = 0) {
        flush();
        connected = true;
        const second = 1 - first;
        Y.applyUpdate(documents[first]!, Y.encodeStateAsUpdate(documents[second]!), 'relay');
        Y.applyUpdate(documents[second]!, Y.encodeStateAsUpdate(documents[first]!), 'relay');
        awareness.forEach((source, index) =>
          applyAwarenessUpdate(
            awareness[1 - index]!,
            encodeAwarenessUpdate(source, [...source.getStates().keys()]),
            'relay'
          )
        );
        flush();
      },
      async savedDigests() {
        flush();
        return Promise.all(
          editors.map(async (editor) => {
            const result = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!result.ok) throw new Error(`Reopen failed: ${result.reason}`);
            return semanticDigest(result.package.parts.values());
          })
        );
      },
      destroy,
    };
  } catch (error) {
    destroy();
    throw error;
  }
}
