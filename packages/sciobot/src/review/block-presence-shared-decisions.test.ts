// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  withPart,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { presenceAttribution, presencePackage } from '../../test/support/block-presence';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { createBlockPresenceRecord } from './block-presence-plan';
import { withBlockPresenceRecord } from './block-presence-store';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { readReviewGroups } from './groups';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';

const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
const word = (pkg: OoxmlPackage) => {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return native(read.package);
};
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const remove of [false, true])
    for (const action of ['accept', 'reject'] as const)
      test(`presence shared decision ${clients}/${remove}/${action} retains outside edits, exact native output, views, histories, and room reopening`, async () => {
        const source = presencePackage(),
          inner = buildBlockIndex(source).byId.get('inner')!;
        const changed = withPayloadRecord(
          source,
          { ...inner.payload, type: 'lesson_visual' },
          'inner label'
        );
        if (!changed.ok) throw Error(changed.reason);
        const pair = await createTwoEditorFixture({
          document: writeOoxmlPackage(changed.pkg),
          documentId: propertyContext.documentId,
          ownedAttributes: true,
          clientIds: clients,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors;
          const baseline = alice!.surface!.session.currentPackage();
          const rootId = buildBlockIndex(baseline).byId.get('outer')!.controlId;
          const outside = bob!
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'Outside')!;
          const prepared = createBlockPresenceRecord(
            remove ? baseline : presencePackage(false),
            baseline,
            [{ blockId: 'outer', before: { present: remove }, after: { present: !remove } }],
            presenceAttribution
          );
          if (!prepared.success) throw Error(prepared.error);
          const stored = alice!.surface!.applyAutomationOps(
            () => [],
            { kind: 'body' },
            [
              (working) => {
                const result = withBlockPresenceRecord(working, prepared.record);
                if (!result.success) throw Error(result.error);
                return result.pkg;
              },
            ],
            true,
            true,
            'after'
          );
          expect(stored.rejected).toBe(false);
          await pair.settle();
          const beforeBytes = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
          let read: ReturnType<typeof readCommittedView> | undefined;
          const inspected = alice!.surface!.applyAutomationOps(
            () => [],
            { kind: 'body' },
            [
              (working) => {
                read = readCommittedView(alice!.surface!.session, propertyToken(alice!));
                return working;
              },
            ],
            true,
            true,
            'after'
          );
          expect(inspected.rejected).toBe(false);
          expect((await read!)!.success).toBe(true);
          await pair.settle();
          pair.documents.forEach((doc, index) =>
            expect(Y.encodeStateAsUpdate(doc)).toEqual(beforeBytes[index]!)
          );
          const groups = readReviewGroups(alice!.surface!.session.currentPackage());
          if (!groups.success || groups.groups.length !== 1) throw Error('Missing presence group');
          const id = groups.groups[0]!.id;
          pair.pause();
          expect(
            bob!.surface!.applyAutomationOps(() => [
              { op: 'insertText', paragraphId: outside, offset: 7, text: ' Peer' },
            ]).rejected
          ).toBe(false);
          const decided = (action === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
            alice!,
            id,
            propertyToken(alice!),
            propertyContext
          );
          if (!decided.success) throw Error(decided.error);
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const retained = remove ? action === 'reject' : action === 'accept';
          const expectedPeer = applyTreeOp(baseline.parts.get(baseline.mainDocumentPart)!, {
            op: 'insertText',
            paragraphId: outside,
            offset: 7,
            text: ' Peer',
          });
          if (!expectedPeer.ok) throw Error(expectedPeer.reason);
          const expected = withPart(
            baseline,
            retained
              ? expectedPeer.part
              : {
                  ...expectedPeer.part,
                  root: {
                    ...expectedPeer.part.root,
                    children: expectedPeer.part.root.children.map((node) =>
                      node.kind === 'body'
                        ? ({
                            ...node,
                            children: node.children.filter((child) => child.id !== rootId),
                          } as OoxmlElement)
                        : node
                    ),
                  } as OoxmlElement,
                }
          );
          const check = async () => {
            for (const editor of pair.editors) {
              const pkg = editor.surface!.session.currentPackage();
              expect(word(pkg)).toBe(word(expected));
              expect(buildBlockIndex(pkg).byId.size).toBe(retained ? 2 : 0);
              const payloads = readPayloadStore(pkg);
              expect(payloads.status === 'ready' && [...payloads.entries.keys()].sort()).toEqual(
                retained ? ['inner', 'outer'] : []
              );
              const rows = readReviewGroups(pkg);
              expect(rows.success && rows.groups.length).toBe(0);
              const original = await readCommittedView(
                editor.surface!.session,
                propertyToken(editor)
              );
              const working = await readWorkingView(editor.surface!.session, propertyToken(editor));
              if (!original.success || !working.success) throw Error('Missing resolved views');
              expect(original.view.projection.text).toBe(working.view.projection.text);
              expect(original.view.projection.text).toContain('Outside Peer');
            }
          };
          await check();
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            const pkg = editor.surface!.session.currentPackage();
            expect(word(pkg)).toBe(word(withPart(baseline, expectedPeer.part)));
            const rows = readReviewGroups(pkg);
            expect(rows.success && rows.groups.map((row) => row.id)).toEqual([id]);
          }
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check();
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(paragraphTextOf(editor.surface!.session.part(), outside)).toBe('Outside');
            const rows = readReviewGroups(editor.surface!.session.currentPackage());
            expect(rows.success && rows.groups.length).toBe(0);
          }
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check();
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          const reopened = await createTwoEditorFixture({
            document: new Uint8Array(await bob!.save()),
            documentId: propertyContext.documentId,
            ownedAttributes: true,
            editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
            roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
          });
          try {
            await reopened.settle();
            for (const editor of reopened.editors)
              expect(word(editor.surface!.session.currentPackage())).toBe(word(expected));
            reopened.ownership.forEach((handle) => handle.assertHealthy());
          } finally {
            reopened.destroy();
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 30000);
