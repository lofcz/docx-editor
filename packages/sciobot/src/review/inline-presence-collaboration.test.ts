import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
// The evaluated collaboration fixture uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { sourceControl } from '../../test/support/target-source-fixture';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { createMentionPayload } from '../blocks/mention-codec';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { createBlockPresenceRecord } from './block-presence-plan';
import { withBlockPresenceRecord } from './block-presence-store';
import { stageTrackedInlineControl } from './tracked-stage-inline-control';
import { readReviewGroups } from './groups';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import { readEditorRevision } from '../projection/revision-cache';
const identity = { documentId: 'inline-presence-room', epoch: 'shared' };
const attribution = { id: 'inline-proposal', author: 'AI', date: '2026-10-24T12:00:00.000Z' };
const payload = createMentionPayload('inline-chip', 'user-7', 'Ada', { privateData: { keep: 7 } });
const run = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const word = (part: OoxmlPart) => canonicalOoxmlFingerprint(part);
function fixture(present: boolean) {
  const content = present
    ? `${run('Lead. ')}${sourceControl(payload.id, run('Ada'))}${run(' Tail.')}`
    : run('Lead.  Tail.');
  const read = readOoxmlPackage(
    docxFixture(
      `<w:p><w:hyperlink w:anchor="target" w:tooltip="Keep">${content}</w:hyperlink></w:p>`
    )
  );
  if (!read.ok) throw Error(read.reason);
  if (!present) return read.package;
  const stored = withPayloadRecord(read.package, payload, 'Mention');
  if (!stored.ok) throw Error(stored.reason);
  return stored.pkg;
}
function native(part: OoxmlPart, op: TreeDocOp) {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
const payloads = (pkg: OoxmlPackage) => {
  const index = buildBlockIndex(pkg);
  expect(index.diagnostics).toEqual([]);
  return [...index.byId]
    .map(([id, block]) => [id, block.payload])
    .sort(([a], [b]) => String(a).localeCompare(String(b)));
};
for (const remove of [false, true])
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`inline native and payload ownership preserve both actor histories: ${remove}/${clients}/${decision}`, async () => {
        const options = {
          document: writeOoxmlPackage(fixture(remove)),
          documentId: identity.documentId,
          ownedAttributes: true,
          clientIds: clients,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        };
        const pair = await createTwoEditorFixture(options),
          actor = pair.editors[0]!,
          peer = pair.editors[1]!;
        try {
          const before = actor.surface!.session.currentPackage(),
            original = actor.surface!.session.part();
          const paragraphId = actor.surface!.session.paragraphIds()[0]!;
          const controlId = remove ? buildBlockIndex(before).byId.get(payload.id)!.controlId : null;
          const oldPayloads = payloads(before);
          pair.pause();
          const installed = actor.surface!.applyAutomationOps(
            () =>
              remove
                ? [
                    {
                      op: 'removeContentControl',
                      controlId: controlId!,
                      keepContent: false,
                      revision: attribution,
                    },
                  ]
                : [],
            { kind: 'body' },
            [
              (working) => {
                let next = working;
                if (!remove) {
                  stageTrackedInlineControl(
                    {
                      currentPackage: () => next,
                      apply(partName, op) {
                        next = withPart(next, native(next.parts.get(partName)!, op));
                      },
                      edit(change) {
                        next = change(next);
                      },
                    },
                    next.mainDocumentPart,
                    {
                      op: 'insertInlineContentControl',
                      paragraphId,
                      offset: 6,
                      tag: 'scio:1:' + payload.id,
                      text: 'Ada',
                      alias: 'Mention',
                    },
                    attribution
                  );
                  const stored = withPayloadRecord(next, payload, 'Mention');
                  if (!stored.ok) throw Error(stored.reason);
                  next = stored.pkg;
                }
                const record = createBlockPresenceRecord(
                  before,
                  next,
                  [
                    {
                      blockId: payload.id,
                      before: { present: remove },
                      after: { present: !remove },
                    },
                  ],
                  attribution,
                  { nativeInline: true }
                );
                if (!record.success) throw Error(record.error);
                const stored = withBlockPresenceRecord(next, record.record);
                if (!stored.success) throw Error(stored.error);
                return stored.pkg;
              },
            ],
            false,
            true,
            'after'
          );
          expect(installed.rejected).toBe(false);
          const proposed = actor.surface!.session.part(),
            proposedPayloads = payloads(actor.surface!.session.currentPackage());
          const groups = readReviewGroups(actor.surface!.session.currentPackage());
          if (!groups.success) throw Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          expect(groups.groups[0]!.readOnly).toBe(false);
          expect(groups.groups[0]!.native).toHaveLength(1);
          expect(groups.groups[0]!.properties).toHaveLength(1);
          expect(
            peer.surface!.applyAutomationOps(
              () => [{ op: 'insertText', paragraphId, offset: 1, text: ' peer' }],
              { kind: 'body' },
              [],
              false,
              true
            ).rejected
          ).toBe(false);
          pair.reconnect(0);
          await pair.settle();
          const check = (pending: boolean, typed: boolean) => {
            const source = pending ? proposed : original;
            const expected = typed
              ? native(source, { op: 'insertText', paragraphId, offset: 1, text: ' peer' })
              : source;
            for (const editor of pair.editors) {
              expect(word(editor.surface!.session.part())).toBe(word(expected));
              expect(payloads(editor.surface!.session.currentPackage())).toEqual(
                pending ? proposedPayloads : oldPayloads
              );
              const read = readReviewGroups(editor.surface!.session.currentPackage());
              if (!read.success) throw Error(read.error);
              expect(read.groups).toHaveLength(Number(pending));
            }
          };
          check(true, true);
          expect(actor.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(false, true);
          expect(actor.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(true, true);
          expect(peer.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(true, false);
          expect(actor.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(false, false);
          expect(peer.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(false, true);
          expect(actor.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(true, true);
          const pending = actor.surface!.session.part();
          const revision = readEditorRevision(actor, identity.documentId, identity.epoch)!;
          const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
            actor,
            groups.groups[0]!.id,
            revision,
            identity
          );
          if (!result.success) throw Error(result.error);
          await pair.settle();
          const decided = actor.surface!.session.part();
          const decidedPayloads = payloads(actor.surface!.session.currentPackage());
          expect(collectRevisionSites(decided)).toEqual([]);
          expect(
            buildBlockIndex(actor.surface!.session.currentPackage()).byId.has(payload.id)
          ).toBe(decision === 'accept' ? !remove : remove);
          if (decision === 'reject')
            expect(word(decided)).toBe(
              word(native(original, { op: 'insertText', paragraphId, offset: 1, text: ' peer' }))
            );
          expect(actor.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(word(editor.surface!.session.part())).toBe(word(pending));
            expect(payloads(editor.surface!.session.currentPackage())).toEqual(proposedPayloads);
          }
          expect(actor.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          peer.surface!.setSelection({
            anchor: { paragraphId, offset: 1 },
            head: { paragraphId, offset: 1 },
          });
          await pair.settle();
          const geometry = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId, offset: 1 })
          );
          expect(geometry[0]).not.toBeNull();
          expect(geometry[1]).not.toBeNull();
          expect({ x: geometry[0]!.x, y: geometry[0]!.y }).toEqual({
            x: geometry[1]!.x,
            y: geometry[1]!.y,
          });
          for (const shared of [false, true]) {
            const reopened = await createTwoEditorFixture({
              ...options,
              document: new Uint8Array(await actor.save()),
              clientIds: undefined,
              ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
            });
            try {
              for (const editor of reopened.editors) {
                expect(word(editor.surface!.session.part())).toBe(word(decided));
                expect(payloads(editor.surface!.session.currentPackage())).toEqual(decidedPayloads);
              }
            } finally {
              reopened.destroy();
            }
          }
        } finally {
          pair.destroy();
        }
      });
