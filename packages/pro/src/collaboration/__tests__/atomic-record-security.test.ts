/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { ATOMIC_CONTENT_NAMESPACE } from '@docx-editor.dev/core/store';
import { DocumentRegistry, applyPrimitiveJournal, seedPackage, MemoryBlobStore } from '../document';
import { asLogicalId } from '../document/identity';
import { destroyReplica, findText, loadPackage, packageOf, seedReplica } from './document-support';
import { zipDocument } from './document-peer-support';

const fixture = () =>
  loadPackage(
    zipDocument(
      `<w:p><w:sdt xmlns:a="${ATOMIC_CONTENT_NAMESPACE}" a:atomicContent="1"><w:sdtPr/><w:sdtContent><w:r><w:t>source</w:t></w:r></w:sdtContent></w:sdt></w:p>`
    )
  );

test('atomic fallback seeding and journal replacement enforce text limits', async () => {
  const doc = new Y.Doc();
  const limited = new DocumentRegistry(doc, { maxTextLength: 4 });
  try {
    expect(await seedPackage(limited, fixture(), new MemoryBlobStore())).toMatchObject({
      ok: false,
      code: 'text-too-long',
    });
  } finally {
    limited.destroy();
    doc.destroy();
  }
  const replica = await seedReplica(fixture());
  const bounded = new DocumentRegistry(replica.doc, { maxTextLength: 8 });
  try {
    const id = asLogicalId(findText(packageOf(replica), 'source').id);
    expect(
      applyPrimitiveJournal(bounded, {
        effects: [
          {
            kind: 'spliceText',
            logicalId: id,
            utf16Start: 0,
            deleteCount: 6,
            insert: '123456789',
          },
        ],
      })
    ).toMatchObject({ ok: false, code: 'text-too-long' });
    expect(bounded.record(id)).toMatchObject({ value: 'source' });
  } finally {
    bounded.destroy();
    destroyReplica(replica);
  }
});

test('remote atomic fields reject malformed and oversized values before exposing canonical text', async () => {
  const replica = await seedReplica(fixture());
  const peer = new Y.Doc();
  const bounded = new DocumentRegistry(peer, { maxTextLength: 8 });
  try {
    Y.applyUpdate(peer, Y.encodeStateAsUpdate(replica.doc));
    const id = asLogicalId(findText(packageOf(replica), 'source').id);
    const record = replica.registry.schema.nodes.get(id)!;
    for (const invalid of [null, 42, { text: 'wrong' }, ['wrong'], '123456789']) {
      record.set('atomicText', invalid);
      Y.applyUpdate(peer, Y.encodeStateAsUpdate(replica.doc));
      expect(bounded.record(id)).toBeNull();
      expect(bounded.nodeShape(id)).toBeNull();
    }
    for (const invalid of [null, 42, { text: 'wrong' }, ['wrong'], '123456789']) {
      const history = new Y.Array<unknown>();
      history.push(['valid', invalid]);
      record.set('atomicText', history);
      Y.applyUpdate(peer, Y.encodeStateAsUpdate(replica.doc));
      expect(bounded.record(id)).toBeNull();
      expect(bounded.nodeShape(id)).toBeNull();
    }
    const history = new Y.Array<string>();
    history.push(['12345678']);
    record.set('atomicText', history);
    Y.applyUpdate(peer, Y.encodeStateAsUpdate(replica.doc));
    expect(bounded.record(id)).toMatchObject({ value: '12345678' });
    expect(bounded.nodeShape(id)?.textLength).toBe(8);
  } finally {
    bounded.destroy();
    peer.destroy();
    destroyReplica(replica);
  }
});

test('atomic history entries refuse at the existing child limit before any journal writes', async () => {
  const replica = await seedReplica(fixture());
  const peer = new Y.Doc();
  const bounded = new DocumentRegistry(peer, { maxChildren: 2 });
  try {
    const id = asLogicalId(findText(packageOf(replica), 'source').id);
    const history = replica.registry.schema.nodes.get(id)!.get('atomicText') as Y.Array<string>;
    history.push(['second']);
    Y.applyUpdate(peer, Y.encodeStateAsUpdate(replica.doc));
    const before = Y.encodeStateVector(peer);
    expect(
      applyPrimitiveJournal(bounded, {
        effects: [
          { kind: 'spliceText', logicalId: id, utf16Start: 0, deleteCount: 6, insert: 'first' },
          { kind: 'spliceText', logicalId: id, utf16Start: 0, deleteCount: 5, insert: 'second' },
        ],
      })
    ).toMatchObject({ ok: false, code: 'too-many-children' });
    expect(Y.encodeStateVector(peer)).toEqual(before);
    expect(bounded.record(id)).toMatchObject({ value: 'second' });
    history.push(['third']);
    Y.applyUpdate(peer, Y.encodeStateAsUpdate(replica.doc));
    expect(bounded.record(id)).toBeNull();
    expect(bounded.nodeShape(id)).toBeNull();
  } finally {
    bounded.destroy();
    peer.destroy();
    destroyReplica(replica);
  }
});
