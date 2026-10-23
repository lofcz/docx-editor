// Test-only evaluated transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  writeOoxmlPackage,
  readOoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import {
  SharedBlobStore,
  SHARED_BLOBS_KEY,
} from '../../../pro/src/collaboration/shared-blob-store';
import { JOURNAL_ORIGIN } from '../../../pro/src/collaboration/document/schema';
import {
  applyPrimitiveJournal,
  DocumentRegistry,
} from '../../../pro/src/collaboration/document/index';
import { asLogicalId } from '../../../pro/src/collaboration/document/identity';
import {
  applyJournal,
  destroyReplica,
  joinReplica,
  packageOf,
} from '../../../pro/src/collaboration/__tests__/document-support';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { table } from '../../test/support/target-table-fixture';
import { nodeChildren, rowCells, tableRows } from '../../test/support/native-table-oracle';
import { prepareNativeDocument } from '../native/prepare-document';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { planTargetDocument } from './plan-target';
import { applyDirectPlan } from '../ai/apply-direct';
const roomBlobs = (doc: Y.Doc) => new SharedBlobStore(doc.getMap<Uint8Array>(SHARED_BLOBS_KEY));
const nativeTable = (part: OoxmlPart) =>
  nodeChildren(bodyStoryRoot(part)!).find((node) => node.kind === 'table')!;
const texts = (part: OoxmlPart) =>
  tableRows(part, nativeTable(part).id).map((row) =>
    rowCells(row).map((cell) =>
      nodeChildren(cell)
        .filter((node) => node.kind === 'paragraph')
        .map((node) => paragraphTextOf(part, node.id))
        .join('\n')
    )
  );
/** Independent expected tree: retain every native node and only reorder observed row identities. */
function reordered(part: OoxmlPart, rows: readonly string[]): OoxmlPart {
  const table = nativeTable(part);
  const children = nodeChildren(table);
  const visit = (node: OoxmlNode): OoxmlNode =>
    node.kind === 'textValue'
      ? node
      : ({
          ...node,
          children:
            node.id === table.id
              ? [
                  ...children.filter((child) => child.kind !== 'tableRow'),
                  ...rows.map((id) => children.find((child) => child.id === id)!),
                ]
              : node.children.map(visit),
        } as OoxmlNode);
  return { ...part, root: visit(part.root) as OoxmlPart['root'] };
}
for (const kind of ['middle', 'end', 'prefix', 'cross-prefix'] as const)
  test(`native ${kind} move addresses visible rows and retains a hidden birth awaiting peer text`, async () => {
    const other =
      kind === 'cross-prefix'
        ? '<w:p><w:r><w:t>Gap</w:t></w:r></w:p>' +
          table([
            ['Other', 'Two'],
            ['Other2', 'More'],
          ])
        : '';
    const prepared = prepareNativeDocument(
      docxFixture(
        table([
          ['A', 'B'],
          ['Keep', 'Peer'],
          ['End', 'Last'],
        ]) + other
      )
    );
    if (!prepared.success) throw Error(prepared.error);
    const pair = await createTwoEditorFixture({ document: prepared.document, clientIds: [41, 73] });
    try {
      const [alice, bob] = pair.editors;
      const initial = alice!.surface!.session.part();
      const sourceTable = nativeTable(initial);
      const otherTable = nodeChildren(bodyStoryRoot(initial)!).filter(
        (node) => node.kind === 'table'
      )[1];
      const originalRows = tableRows(initial, sourceTable.id);
      const context = { documentId: 'hidden-move', epoch: 'native' };
      const revision = readEditorRevision(alice!, context.documentId, context.epoch)!;
      const source = await readWorkingView(alice!.surface!.session, revision);
      if (!source.success) throw Error(source.error);
      const planned = planTargetDocument(
        {
          revision,
          pkg: alice!.surface!.session.currentPackage(),
          projection: source.view.projection,
        },
        '| New | Tail $x$ |\n| - | - |\n| A | B |\n| Keep | Peer |\n| End | Last |' +
          (other ? '\n\nGap\n\n| Other | Two |\n| - | - |\n| Other2 | More |' : '')
      );
      if (!planned.success) throw Error(planned.error);
      const applied = applyDirectPlan(alice!, planned.plan, context);
      if (!applied.success) throw Error(applied.error);
      await pair.settle();
      pair.pause();
      const paragraphId = bob!
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob!.surface!.session.part(), id)?.startsWith('Tail'))!;
      expect(
        bob!.surface!.applyAutomationOps(
          () => [
            {
              op: 'insertText',
              paragraphId,
              offset: paragraphTextOf(bob!.surface!.session.part(), paragraphId)!.length,
              text: ' peerborn',
            },
          ],
          { kind: 'body' },
          [],
          false,
          true
        ).rejected
      ).toBe(false);
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(
        canonicalOoxmlFingerprint(initial)
      );
      const replica = joinReplica(
        {
          doc: pair.documents[0]!,
          blobs: roomBlobs(pair.documents[0]!),
        },
        97,
        undefined,
        roomBlobs
      );
      const undo = new Y.UndoManager([...replica.registry.trackedTypes()], {
        captureTimeout: 0,
        trackedOrigins: new Set([JOURNAL_ORIGIN]),
        deleteFilter: replica.registry.undoDeleteFilter(),
      });
      try {
        const base = nodeChildren(sourceTable).findIndex(
          (child) => child.id === originalRows[0]!.id
        );
        const effects =
          kind === 'cross-prefix'
            ? [
                {
                  kind: 'moveNode' as const,
                  logicalId: originalRows[2]!.id,
                  destinationParentLogicalId: otherTable!.id,
                  destinationIndex: nodeChildren(otherTable!).findIndex(
                    (node) => node.kind === 'tableRow'
                  ),
                },
                {
                  kind: 'moveNode' as const,
                  logicalId: originalRows[2]!.id,
                  destinationParentLogicalId: sourceTable.id,
                  destinationIndex: base + 1,
                },
              ]
            : kind === 'end'
              ? [
                  {
                    kind: 'moveNode' as const,
                    logicalId: originalRows[0]!.id,
                    destinationParentLogicalId: sourceTable.id,
                    destinationIndex: 99,
                  },
                ]
              : [
                  {
                    kind: 'moveNode' as const,
                    logicalId: originalRows[2]!.id,
                    destinationParentLogicalId: sourceTable.id,
                    destinationIndex: base + 1,
                  },
                  ...(kind === 'prefix'
                    ? [
                        {
                          kind: 'moveNode' as const,
                          logicalId: originalRows[0]!.id,
                          destinationParentLogicalId: sourceTable.id,
                          destinationIndex: base + 2,
                        },
                      ]
                    : []),
                ];
        applyJournal(replica, { effects });
        const order =
          kind === 'middle' || kind === 'cross-prefix'
            ? [0, 2, 1]
            : kind === 'end'
              ? [1, 2, 0]
              : [2, 1, 0];
        if (otherTable)
          expect(replica.registry.childArray(asLogicalId(otherTable.id)).toArray()).not.toContain(
            originalRows[2]!.id
          );
        const expected = reordered(
          initial,
          order.map((at) => originalRows[at]!.id)
        );
        const currentPart = () =>
          packageOf(replica).parts.get(packageOf(replica).mainDocumentPart)!;
        expect(canonicalOoxmlFingerprint(currentPart())).toBe(canonicalOoxmlFingerprint(expected));
        if (otherTable) {
          const boundedDoc = new Y.Doc();
          const bounded = new DocumentRegistry(boundedDoc, { maxChildren: 6 });
          try {
            bounded.beginBulkLoad();
            Y.applyUpdate(boundedDoc, Y.encodeStateAsUpdate(replica.doc), 'join');
            bounded.endBulkLoad();
            const before = Y.encodeStateAsUpdate(boundedDoc);
            const row = tableRows(initial, otherTable.id)[0]!;
            expect(
              applyPrimitiveJournal(bounded, {
                effects: [
                  {
                    kind: 'moveNode',
                    logicalId: row.id,
                    destinationParentLogicalId: sourceTable.id,
                    destinationIndex: base,
                  },
                ],
              })
            ).toMatchObject({ ok: false, code: 'too-many-children' });
            expect(Y.encodeStateAsUpdate(boundedDoc)).toEqual(before);
          } finally {
            bounded.destroy();
            boundedDoc.destroy();
          }
        }
        if (kind === 'middle') {
          const before = Y.encodeStateAsUpdate(replica.doc);
          const destination = replica.mint.take();
          expect(
            applyPrimitiveJournal(replica.registry, {
              effects: [
                { kind: 'putNode', descriptor: { kind: 'textValue', logicalId: destination } },
                {
                  kind: 'moveNode',
                  logicalId: originalRows[2]!.id,
                  destinationParentLogicalId: destination,
                  destinationIndex: 0,
                },
              ],
            })
          ).toMatchObject({ ok: false, code: 'invalid-bound' });
          expect(Y.encodeStateAsUpdate(replica.doc)).toEqual(before);
          expect(replica.registry.hasNode(destination)).toBe(false);
        }
        Y.applyUpdate(replica.doc, Y.encodeStateAsUpdate(pair.documents[1]!), 'relay');
        expect(replica.materializer.rebuild().ok).toBe(true);
        const expectedRows = [
          ['', ' peerborn'],
          ...order.map(
            (at) =>
              [
                ['A', 'B'],
                ['Keep', 'Peer'],
                ['End', 'Last'],
              ][at]!
          ),
        ];
        expect(texts(currentPart())).toEqual(expectedRows);
        const withPeer = canonicalOoxmlFingerprint(currentPart());
        const expectedUndo = canonicalOoxmlFingerprint(
          reordered(currentPart(), [
            tableRows(currentPart(), sourceTable.id)[0]!.id,
            ...originalRows.map((row) => row.id),
          ])
        );
        const beforeReads = Y.encodeStateAsUpdate(replica.doc);
        const cold = joinReplica(replica, 98, undefined, roomBlobs);
        try {
          expect(
            canonicalOoxmlFingerprint(packageOf(cold).parts.get(packageOf(cold).mainDocumentPart)!)
          ).toBe(withPeer);
        } finally {
          destroyReplica(cold);
        }
        const opened = readOoxmlPackage(writeOoxmlPackage(packageOf(replica)));
        if (!opened.ok) throw Error(opened.reason);
        expect(
          canonicalOoxmlFingerprint(opened.package.parts.get(opened.package.mainDocumentPart)!)
        ).toBe(withPeer);
        expect(Y.encodeStateAsUpdate(replica.doc)).toEqual(beforeReads);
        undo.undo();
        expect(replica.materializer.rebuild().ok).toBe(true);
        expect(texts(currentPart())).toEqual([
          ['', ' peerborn'],
          ['A', 'B'],
          ['Keep', 'Peer'],
          ['End', 'Last'],
        ]);
        expect(canonicalOoxmlFingerprint(currentPart())).toBe(expectedUndo);
        undo.redo();
        expect(replica.materializer.rebuild().ok).toBe(true);
        expect(canonicalOoxmlFingerprint(currentPart())).toBe(withPeer);
      } finally {
        undo.destroy();
        destroyReplica(replica);
      }
    } finally {
      pair.destroy();
    }
  }, 60_000);
