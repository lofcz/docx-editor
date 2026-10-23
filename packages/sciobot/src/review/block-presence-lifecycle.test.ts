import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
  type OoxmlElement,
  withPart,
} from '@docx-editor.dev/core/store';
import { presenceAttribution, presencePackage } from '../../test/support/block-presence';
import {
  propertyEditor,
  propertyToken,
  propertyContext,
} from '../../test/support/property-execution';
import { readPayloadStore, withPayloadRecord } from '../blocks/payload-opc-codec';
import { buildBlockIndex } from '../blocks/block-index';
import { createBlockPresenceRecord } from './block-presence-plan';
import { withBlockPresenceRecord } from './block-presence-store';
import { readApplicationProposals } from './application-proposals';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { readReviewGroups } from './groups';
import { restoreReviewState } from './restore';
import { acceptReviewGroup } from './accept-group';
import { rejectReviewGroup } from './reject-group';
function fixture() {
  const pkg = presencePackage(),
    inner = buildBlockIndex(pkg).byId.get('inner')!;
  const changed = withPayloadRecord(
    pkg,
    { ...inner.payload, type: 'lesson_visual' },
    'inner label'
  );
  if (!changed.ok) throw Error(changed.reason);
  return changed.pkg;
}
const fingerprint = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
for (const remove of [false, true])
  for (const action of ['accept', 'reject'] as const)
    test(`presence ${remove ? 'removal' : 'insertion'} ${action} has combined views, stable groups, exact native output, payload decisions, and history`, async () => {
      const source = fixture(),
        editor = propertyEditor(source);
      try {
        const canonical = editor.surface!.session.currentPackage();
        const part = canonical.parts.get(canonical.mainDocumentPart)!;
        const removedId = buildBlockIndex(canonical).byId.get('outer')!.controlId;
        const absent = withPart(canonical, {
          ...part,
          root: {
            ...part.root,
            children: part.root.children.map((node) =>
              node.kind === 'body'
                ? ({
                    ...node,
                    children: node.children.filter((child) => child.id !== removedId),
                  } as OoxmlElement)
                : node
            ),
          } as OoxmlElement,
        });
        const word = (pkg: OoxmlPackage) => {
          const reopened = readOoxmlPackage(writeOoxmlPackage(pkg));
          if (!reopened.ok) throw Error(reopened.reason);
          return fingerprint(reopened.package);
        };
        const initial = await readWorkingView(editor.surface!.session, propertyToken(editor));
        if (!initial.success) throw Error(initial.error);
        const prepared = createBlockPresenceRecord(
          remove ? canonical : presencePackage(false),
          canonical,
          [{ blockId: 'outer', before: { present: remove }, after: { present: !remove } }],
          presenceAttribution
        );
        if (!prepared.success) throw Error(prepared.error);
        const installed = editor.surface!.applyAutomationOps(
          () => [],
          { kind: 'body' },
          [
            (working) => {
              const saved = withBlockPresenceRecord(working, prepared.record);
              if (!saved.success) throw Error(saved.error);
              return saved.pkg;
            },
          ],
          true,
          true,
          'after'
        );
        expect(installed.rejected).toBe(false);
        const token = propertyToken(editor),
          pending = editor.surface!.session.currentPackage();
        const nativeBefore = fingerprint(pending),
          revisionBefore = editor.surface!.session.packageRevision();
        const original = await readCommittedView(editor.surface!.session, token),
          proposed = await readWorkingView(editor.surface!.session, token);
        if (!original.success || !proposed.success) throw Error('Missing combined views');
        expect(original.view.layout.revision).toBe(token.revision);
        expect(proposed.view.layout.revision).toBe(token.revision);
        expect(original.view.projection.text.trimEnd()).toBe(
          remove ? initial.view.projection.text.trimEnd() : '**Outside**'
        );
        expect(proposed.view.projection.text.trimEnd()).toBe(
          remove ? '**Outside**' : initial.view.projection.text.trimEnd()
        );
        expect(fingerprint(editor.surface!.session.currentPackage())).toBe(nativeBefore);
        expect(editor.surface!.session.packageRevision()).toBe(revisionBefore);
        const groups = readReviewGroups(pending);
        if (!groups.success) throw Error(groups.error);
        expect(groups.groups).toHaveLength(1);
        const restored = restoreReviewState(pending, token);
        if (!restored.success) throw Error(restored.error);
        expect(restored.state.pendingPresence.map((record) => record.id)).toEqual([
          presenceAttribution.id,
        ]);
        expect(restored.state.reviewableGroupIds).toEqual(groups.groups.map((group) => group.id));
        const group = groups.groups[0]!;
        expect(group.readOnly).toBe(false);
        expect(group.native).toHaveLength(0);
        expect(group.properties).toEqual([
          { proposalId: presenceAttribution.id, blockId: 'outer' },
        ]);
        const decided = (action === 'accept' ? acceptReviewGroup : rejectReviewGroup)(
          editor,
          group.id,
          token,
          propertyContext
        );
        if (!decided.success) throw Error(decided.error);
        const retained = remove ? action === 'reject' : action === 'accept';
        const current = editor.surface!.session.currentPackage();
        const read = readReviewGroups(current);
        if (!read.success) throw Error(read.error);
        expect(read.groups).toHaveLength(0);
        const proposals = readApplicationProposals(current, propertyToken(editor));
        expect(proposals.success && proposals.presence.length).toBe(0);
        expect(buildBlockIndex(current).byId.size).toBe(retained ? 2 : 0);
        const payloads = readPayloadStore(current);
        if (payloads.status !== 'ready') throw Error('Missing payload store');
        expect(payloads.entries.size).toBe(retained ? 2 : 0);
        const reopened = readOoxmlPackage(writeOoxmlPackage(current));
        if (!reopened.ok) throw Error(reopened.reason);
        const restoredDecision = restoreReviewState(reopened.package, {
          ...token,
          epoch: 'word-reopen',
        });
        if (!restoredDecision.success) throw Error(restoredDecision.error);
        expect(restoredDecision.state.pendingPresence).toHaveLength(0);
        expect(fingerprint(reopened.package)).toBe(word(retained ? canonical : absent));
        if (retained) expect(fingerprint(reopened.package)).toBe(word(canonical));
        else {
          // Retain the exact imported Outside paragraph, including its Word identity.
          const expected = await readCommittedView(editor.surface!.session, propertyToken(editor));
          if (!expected.success) throw Error(expected.error);
          expect(expected.view.projection.text.trimEnd()).toBe('**Outside**');
          expect(buildBlockIndex(reopened.package).diagnostics).toEqual([]);
        }
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(fingerprint(editor.surface!.session.currentPackage())).toBe(nativeBefore);
        const undone = readReviewGroups(editor.surface!.session.currentPackage());
        if (!undone.success) throw Error(undone.error);
        expect(undone.groups.map((group) => group.id)).toEqual([group.id]);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(
          readApplicationProposals(editor.surface!.session.currentPackage(), propertyToken(editor))
            .success
        ).toBe(true);
      } finally {
        editor.destroy();
      }
    }, 30000);

for (const conflict of ['duplicate', 'native-edit'] as const)
  test(`presence ${conflict} refuses decisions and retains the complete current document`, async () => {
    const editor = propertyEditor(fixture());
    try {
      const prepared = createBlockPresenceRecord(
        presencePackage(false),
        editor.surface!.session.currentPackage(),
        [{ blockId: 'outer', before: { present: false }, after: { present: true } }],
        presenceAttribution
      );
      if (!prepared.success) throw Error(prepared.error);
      const installed = editor.surface!.applyAutomationOps(
        () => [],
        { kind: 'body' },
        [
          (working) => {
            const first = withBlockPresenceRecord(working, prepared.record);
            if (!first.success) throw Error(first.error);
            if (conflict !== 'duplicate') return first.pkg;
            const second = withBlockPresenceRecord(first.pkg, {
              ...prepared.record,
              id: 'presence-two',
            });
            if (!second.success) throw Error(second.error);
            return second.pkg;
          },
        ],
        true,
        true,
        'after'
      );
      expect(installed.rejected).toBe(false);
      if (conflict === 'native-edit') {
        const session = editor.surface!.session;
        const paragraph = session
          .paragraphIds()
          .find((id) => paragraphTextOf(session.part(), id) === 'Opaque child body')!;
        expect(
          editor.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: paragraph, offset: 0, text: 'Manual ' },
          ]).rejected
        ).toBe(false);
      }
      const current = editor.surface!.session.currentPackage();
      const before = writeOoxmlPackage(current);
      const token = propertyToken(editor);
      const read = readReviewGroups(current);
      if (!read.success) throw Error(read.error);
      expect(read.groups.length).toBe(conflict === 'duplicate' ? 2 : 1);
      expect(read.groups.every((group) => group.readOnly)).toBe(true);
      expect((await readCommittedView(editor.surface!.session, token)).success).toBe(false);
      for (const group of read.groups)
        expect(rejectReviewGroup(editor, group.id, token, propertyContext).success).toBe(false);
      expect(writeOoxmlPackage(editor.surface!.session.currentPackage())).toEqual(before);
      expect(propertyToken(editor)).toEqual(token);
    } finally {
      editor.destroy();
    }
  }, 30000);
