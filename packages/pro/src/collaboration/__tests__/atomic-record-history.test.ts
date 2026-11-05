/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { ATOMIC_CONTENT_NAMESPACE } from '@docx-editor.dev/core/store';
import { applyPrimitiveJournal } from '../document';
import { asLogicalId } from '../document/identity';
import { JOURNAL_ORIGIN } from '../document/schema';
import {
  destroyReplica,
  findText,
  joinReplica,
  loadPackage,
  packageOf,
  seedReplica,
} from './document-support';
import { zipDocument } from './document-peer-support';

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`pre-seeded atomic text retains concurrent peer replacements through repeated history: ${clients}`, async () => {
    const source = await seedReplica(
      loadPackage(
        zipDocument(
          `<w:p><w:sdt xmlns:a="${ATOMIC_CONTENT_NAMESPACE}" a:atomicContent="1"><w:sdtPr/><w:sdtContent><w:r><w:t>original</w:t></w:r></w:sdtContent></w:sdt></w:p>`
        )
      ),
      undefined,
      clients[0]
    );
    const peer = joinReplica(source, clients[1]);
    const replicas = [source, peer];
    const histories = replicas.map(
      (replica) =>
        new Y.UndoManager([...replica.registry.trackedTypes()], {
          trackedOrigins: new Set([JOURNAL_ORIGIN]),
          captureTimeout: 0,
          deleteFilter: replica.registry.undoDeleteFilter(),
        })
    );
    const id = asLogicalId(findText(packageOf(source), 'original').id);
    const sync = () => {
      Y.applyUpdate(peer.doc, Y.encodeStateAsUpdate(source.doc), 'relay');
      Y.applyUpdate(source.doc, Y.encodeStateAsUpdate(peer.doc), 'relay');
    };
    const value = (index: number) => replicas[index]!.registry.record(id);
    const replace = (index: number, next: string) => {
      const before = value(index);
      if (before?.kind !== 'textValue') throw Error('Missing atomic text record');
      expect(
        applyPrimitiveJournal(replicas[index]!.registry, {
          effects: [
            {
              kind: 'spliceText',
              logicalId: id,
              utf16Start: 0,
              deleteCount: before.value.length,
              insert: next,
            },
          ],
        }).ok
      ).toBe(true);
    };
    try {
      expect(value(0)).toEqual(value(1));
      for (const round of [1, 2, 3]) {
        replace(0, `Alice ${round}`);
        replace(1, `Bob ${round}`);
        sync();
        expect(value(0)).toEqual(value(1));
        for (const replica of replicas)
          expect(replica.registry.nodeShape(id)?.atomicEntries).toBeLessThanOrEqual(2);
        expect(histories[0]!.undo()).not.toBeNull();
        sync();
        for (const index of [0, 1]) expect(value(index)).toMatchObject({ value: `Bob ${round}` });
        expect(histories[0]!.redo()).not.toBeNull();
        sync();
        expect(value(0)).toEqual(value(1));
        expect(value(0)).toMatchObject({
          value: clients[0] > clients[1] ? `Alice ${round}` : `Bob ${round}`,
        });
      }
      const reopened = joinReplica(source, 101);
      try {
        expect(reopened.registry.record(id)).toEqual(value(0));
      } finally {
        destroyReplica(reopened);
      }
    } finally {
      histories.forEach((history) => history.destroy());
      replicas.forEach(destroyReplica);
    }
  });
