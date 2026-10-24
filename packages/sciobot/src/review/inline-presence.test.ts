import { stageTrackedInlineControl } from './tracked-stage-inline-control';
import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  applyTreeOp,
  withPart,
  collectRevisionSites,
  contentControlsIn,
  readOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { sourceControl } from '../../test/support/target-source-fixture';
import {
  propertyEditor,
  propertyToken,
  propertyContext,
} from '../../test/support/property-execution';
import { createMentionPayload } from '../blocks/mention-codec';
import { buildBlockIndex } from '../blocks/block-index';
import { withPayloadRecord, readPayloadStore } from '../blocks/payload-opc-codec';
import { createBlockPresenceRecord, bindBlockPresenceRecord } from './block-presence-plan';
import { withBlockPresenceRecord } from './block-presence-store';
import { readReviewGroups } from './groups';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { planOwnedProposalRemoval } from './owned-proposal-plan';
import { stageOwnedProposalRemoval } from './owned-proposal-stage';
import { TreePackageStore } from '@docx-editor.dev/core/store';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const payload = createMentionPayload('inline-chip', 'user-7', 'Ada', { privateData: { keep: 7 } });
const attribution = {
  id: 'inline-proposal',
  author: propertyContext.author,
  date: '2026-10-24T12:00:00.000Z',
};
function fixture(present: boolean, linked: boolean) {
  const control = sourceControl(payload.id, run('Ada'));
  const content = present ? `${run('Lead. ')}${control}${run(' Tail.')}` : run('Lead.  Tail.');
  const read = readOoxmlPackage(
    docxFixture(
      `<w:p>${linked ? `<w:hyperlink w:anchor="target" w:tooltip="Keep">${content}</w:hyperlink>` : content}</w:p>`
    )
  );
  if (!read.ok) throw Error(read.reason);
  if (!present) return read.package;
  const stored = withPayloadRecord(read.package, payload, 'Mention');
  if (!stored.ok) throw Error(stored.reason);
  return stored.pkg;
}
const word = (pkg: OoxmlPackage) => canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
for (const remove of [false, true])
  for (const decision of ['accept', 'reject'] as const)
    for (const linked of [false, true])
      test(`inline presence binds one native receipt through decision, history and reopen: ${remove}/${decision}/${linked}`, async () => {
        const editor = propertyEditor(fixture(remove, linked));
        try {
          const surface = editor.surface!,
            before = surface.session.currentPackage();
          const controlId = remove ? buildBlockIndex(before).byId.get(payload.id)!.controlId : null;
          const paragraphId = surface.session.paragraphIds()[0]!;
          const staged = surface.applyAutomationOps(
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
                        const applied = applyTreeOp(next.parts.get(partName)!, op);
                        if (!applied.ok) throw Error(applied.reason);
                        next = withPart(next, applied.part);
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
                expect(record.record.schemaVersion).toBe(3);
                expect(bindBlockPresenceRecord(next, record.record).success).toBe(true);
                const stored = withBlockPresenceRecord(next, record.record);
                if (!stored.success) throw Error(stored.error);
                return stored.pkg;
              },
            ],
            false,
            true,
            'after'
          );
          expect(staged.rejected).toBe(false);
          const pending = surface.session.currentPackage(),
            pendingWord = word(pending);
          const groups = readReviewGroups(pending);
          if (!groups.success) throw Error(groups.error);
          expect(groups.groups).toHaveLength(1);
          const group = groups.groups[0]!;
          expect(group.readOnly).toBe(false);
          expect(group.native).toHaveLength(1);
          expect(group.properties).toEqual([{ proposalId: attribution.id, blockId: payload.id }]);
          expect(group.anchor?.paragraphId).toBe(paragraphId);
          const committed = await readCommittedView(surface.session, propertyToken(editor));
          const working = await readWorkingView(surface.session, propertyToken(editor));
          if (!committed.success || !working.success) throw Error('Unable to read inline views');
          expect(committed.view.projection.text.includes('Ada')).toBe(remove);
          expect(working.view.projection.text.includes('Ada')).toBe(!remove);
          const inverse = planOwnedProposalRemoval(
            { pkg: pending, revision: propertyToken(editor) },
            [
              {
                groupId: group.id,
                native: group.native.map((member) => ({
                  partName: member.partName,
                  addresses: [member.address],
                })),
                properties: group.properties,
              },
            ]
          );
          if (!inverse.success) throw Error(inverse.error);
          const store = new TreePackageStore(pending, pending.parts.get(pending.mainDocumentPart)!);
          const inverted = store.transact(
            { kind: 'body' },
            (ctx) =>
              stageOwnedProposalRemoval(ctx, inverse.plan, true, (run) => {
                if (!run()) throw Error('Inverse refused');
              }),
            { recordsHistory: false }
          );
          expect(inverted.ok).toBe(true);
          expect(word(store.currentPackage())).toBe(word(before));
          const result = (decision === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
            editor,
            group.id,
            propertyToken(editor),
            propertyContext
          );
          if (!result.success) throw Error(result.error);
          const decided = surface.session.currentPackage();
          expect(collectRevisionSites(surface.session.part())).toEqual([]);
          const remains = decision === 'accept' ? !remove : remove;
          expect(buildBlockIndex(decided).byId.has(payload.id)).toBe(remains);
          expect(buildBlockIndex(decided).diagnostics).toEqual([]);
          expect(contentControlsIn(surface.session.part().root)).toHaveLength(Number(remains));
          if (remains)
            expect(buildBlockIndex(decided).byId.get(payload.id)!.payload).toEqual(payload);
          if (decision === 'reject') expect(word(decided)).toBe(word(before));
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(word(surface.session.currentPackage())).toBe(pendingWord);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(word(surface.session.currentPackage())).toBe(word(decided));
          const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
          if (!opened.ok) throw Error(opened.reason);
          expect(word(opened.package)).toBe(word(decided));
          expect(buildBlockIndex(opened.package).diagnostics).toEqual([]);
          expect(buildBlockIndex(opened.package).byId.has(payload.id)).toBe(remains);
          const payloads = readPayloadStore(opened.package);
          expect(payloads.status).toBe('ready');
        } finally {
          editor.destroy();
        }
      });
