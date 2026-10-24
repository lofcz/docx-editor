import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { createMentionPayload } from '../blocks/mention-codec';
import { withPayloadRecord, readPayloadStore, blockTag } from '../blocks/payload-opc-codec';
import { buildBlockIndex } from '../blocks/block-index';
import { planTargetDocument } from '../diff/plan-target';
import { previewDirectPlan } from '../ai/apply-direct';
import { readWorkingView } from '../review/working-view';
import { readCommittedView } from '../review/committed-view';
import { readReviewGroups } from '../review/groups';
import { recomposeOwnedProposals } from '../review/recompose';
import { acceptAllReviewGroups } from '../review/accept-all';
import { rejectAllReviewGroups } from '../review/reject-all';

const propertyContext = {
  documentId: 'recompose-inline-removal',
  epoch: 'native',
  author: 'ScioBot',
  date: '2026-10-24T12:00:00.000Z',
};
const propertyEditor = (pkg: OoxmlPackage) =>
  createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(pkg),
    author: 'User',
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
const propertyToken = (editor: ReturnType<typeof propertyEditor>) =>
  readEditorRevision(editor, propertyContext.documentId, propertyContext.epoch)!;
const sourceControl = (id: string, body: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${blockTag(id)}"/></w:sdtPr><w:sdtContent>${body}</w:sdtContent></w:sdt>`;
const payload = createMentionPayload('owned-mention', 'user', 'Ada', { privateData: { keep: 7 } });
const word = (pkg: OoxmlPackage) => canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
function fixture(linked: boolean, locked: boolean) {
  let control = sourceControl(payload.id, '<w:r><w:t>Ada</w:t></w:r>');
  if (locked)
    control = control.replace('</w:sdtPr>', '<w:lock w:val="sdtContentLocked"/></w:sdtPr>');
  const contents =
    '<w:r><w:t xml:space="preserve">Lead. </w:t></w:r>' +
    control +
    '<w:r><w:t xml:space="preserve"> Tail.</w:t></w:r>';
  const loaded = readOoxmlPackage(
    docxFixture(
      '<w:p>' +
        (linked ? '<w:hyperlink w:anchor="target">' + contents + '</w:hyperlink>' : contents) +
        '</w:p>'
    )
  );
  if (!loaded.ok) throw Error(loaded.reason);
  const stored = withPayloadRecord(loaded.package, payload, 'Mention');
  if (!stored.ok) throw Error(stored.reason);
  return stored.pkg;
}
for (const linked of [false, true])
  for (const locked of [false, true])
    for (const decision of ['accept', 'reject'] as const)
      test(`AI recomposition removes whole inline controls through native review: ${linked}/${locked}/${decision}`, async () => {
        const editor = propertyEditor(fixture(linked, locked));
        try {
          const surface = editor.surface!,
            before = surface.session.currentPackage();
          const initial = await readWorkingView(surface.session, propertyToken(editor));
          if (!initial.success) throw Error(initial.error);
          const target = initial.view.projection.text.replace(
            /<scio-mention\b[^>]*\/>|\[Ada\]\(mention:user\)/g,
            ''
          );
          expect(target).not.toBe(initial.view.projection.text);
          const planned = planTargetDocument(
            { pkg: before, revision: propertyToken(editor), projection: initial.view.projection },
            target
          );
          if (!planned.success) throw Error(planned.error);
          expect(planned.plan.operations.map((edit) => edit.kind)).toEqual(['inline-remove']);
          const direct = previewDirectPlan(before, planned.plan);
          if (!direct.success) throw Error(direct.error);
          const result = await recomposeOwnedProposals(
            editor,
            { revision: propertyToken(editor), targetMarkdown: target, ownership: [] },
            propertyContext,
            { payloads: { presence: { proposalId: 'remove-mention' } } }
          );
          if (!result.success) throw Error(result.error);
          expect(result.committed).toBe(true);
          expect(result.ownership).toHaveLength(1);
          expect(result.ownership[0]!.properties).toEqual([
            { proposalId: 'remove-mention', blockId: payload.id },
          ]);
          const pending = surface.session.currentPackage(),
            pendingWord = word(pending);
          expect(buildBlockIndex(pending).byId.get(payload.id)?.payload).toEqual(payload);
          const groups = readReviewGroups(pending);
          expect(groups.success && groups.groups.length).toBe(1);
          const committed = await readCommittedView(surface.session, propertyToken(editor));
          const working = await readWorkingView(surface.session, propertyToken(editor));
          expect(committed.success && committed.view.projection.text).toBe(
            initial.view.projection.text
          );
          expect(working.success && working.view.projection.text).toBe(target);
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(word(surface.session.currentPackage())).toBe(word(before));
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(word(surface.session.currentPackage())).toBe(pendingWord);
          const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            editor,
            propertyToken(editor),
            propertyContext
          );
          if (!decided.success) throw Error(decided.error);
          const expected = decision === 'accept' ? direct.pkg : before;
          expect(word(surface.session.currentPackage())).toBe(word(expected));
          const payloads = readPayloadStore(surface.session.currentPackage());
          expect(payloads.status === 'ready' && payloads.entries.has(payload.id)).toBe(
            decision === 'reject'
          );
          if (decision === 'reject') {
            const original = buildBlockIndex(before).byId.get(payload.id)!;
            expect(findNode(surface.session.part(), original.controlId)).toEqual(
              findNode(before.parts.get(before.mainDocumentPart)!, original.controlId)
            );
          }
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(word(surface.session.currentPackage())).toBe(pendingWord);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(word(surface.session.currentPackage())).toBe(word(expected));
          const saved = readOoxmlPackage(writeOoxmlPackage(surface.session.currentPackage()));
          if (!saved.ok) throw Error(saved.reason);
          const reopened = propertyEditor(saved.package);
          try {
            expect(word(reopened.surface!.session.currentPackage())).toBe(word(expected));
            const reopenedGroups = readReviewGroups(reopened.surface!.session.currentPackage());
            expect(reopenedGroups.success && reopenedGroups.groups).toEqual([]);
          } finally {
            reopened.destroy();
          }
        } finally {
          editor.destroy();
        }
      });

for (const mixed of [false, true])
  test(`AI inline removal refuses missing ownership or dependent coordinates before publication: ${mixed}`, async () => {
    const editor = propertyEditor(fixture(true, true));
    try {
      const surface = editor.surface!,
        before = surface.session.currentPackage();
      const initial = await readWorkingView(surface.session, propertyToken(editor));
      if (!initial.success) throw Error(initial.error);
      let target = initial.view.projection.text.replace(
        /<scio-mention\b[^>]*\/>|\[Ada\]\(mention:user\)/g,
        ''
      );
      if (mixed) target = target.replace('Lead', 'Changed');
      const result = await recomposeOwnedProposals(
        editor,
        { revision: propertyToken(editor), targetMarkdown: target, ownership: [] },
        propertyContext,
        mixed ? { payloads: { presence: { proposalId: 'mixed-removal' } } } : {}
      );
      expect(result.success).toBe(false);
      if (!result.success) expect(result.code).toBe('unsupported-operation');
      expect(surface.session.currentPackage()).toBe(before);
      expect(surface.session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  });

test('AI replaces an owned deletion twice without reusing native ownership or losing the original control', async () => {
  const editor = propertyEditor(fixture(true, true));
  try {
    const surface = editor.surface!,
      before = surface.session.currentPackage();
    const initial = await readWorkingView(surface.session, propertyToken(editor));
    if (!initial.success) throw Error(initial.error);
    const target = initial.view.projection.text.replace(
      /<scio-mention\b[^>]*\/>|\[Ada\]\(mention:user\)/g,
      ''
    );
    const first = await recomposeOwnedProposals(
      editor,
      { revision: propertyToken(editor), targetMarkdown: target, ownership: [] },
      propertyContext,
      { payloads: { presence: { proposalId: 'first-removal' } } }
    );
    if (!first.success) throw Error(first.error);
    const firstPending = word(surface.session.currentPackage());
    const restored = await recomposeOwnedProposals(
      editor,
      {
        revision: propertyToken(editor),
        targetMarkdown: initial.view.projection.text,
        ownership: first.ownership,
        contentProofs: first.contentProofs,
      },
      propertyContext
    );
    if (!restored.success) throw Error(restored.error);
    expect(word(surface.session.currentPackage())).toBe(word(before));
    const afterRestore = readReviewGroups(surface.session.currentPackage());
    expect(afterRestore.success && afterRestore.groups).toEqual([]);
    const third = await recomposeOwnedProposals(
      editor,
      {
        revision: propertyToken(editor),
        targetMarkdown: target,
        ownership: restored.ownership,
        contentProofs: restored.contentProofs,
      },
      propertyContext,
      { payloads: { presence: { proposalId: 'third-removal' } } }
    );
    if (!third.success) throw Error(third.error);
    const groups = readReviewGroups(surface.session.currentPackage());
    if (!groups.success) throw Error(groups.error);
    expect(groups.groups).toHaveLength(1);
    expect(groups.groups[0]!.properties).toEqual([
      { proposalId: 'third-removal', blockId: payload.id },
    ]);
    expect(groups.groups[0]!.native[0]!.address).not.toEqual(
      first.ownership[0]!.native[0]!.addresses[0]
    );
    expect(buildBlockIndex(surface.session.currentPackage()).byId.get(payload.id)?.payload).toEqual(
      payload
    );
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(word(surface.session.currentPackage())).toBe(word(before));
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(word(surface.session.currentPackage())).toBe(firstPending);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(word(surface.session.currentPackage())).toBe(word(before));
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const decided = rejectAllReviewGroups(editor, propertyToken(editor), propertyContext);
    if (!decided.success) throw Error(decided.error);
    expect(word(surface.session.currentPackage())).toBe(word(before));
  } finally {
    editor.destroy();
  }
});
