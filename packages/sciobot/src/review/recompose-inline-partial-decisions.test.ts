import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { createMentionPayload } from '../blocks/mention-codec';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { readEditorRevision } from '../projection/revision-cache';
import { previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { recomposeOwnedProposals } from '../review/recompose';
import { readWorkingView } from '../review/working-view';
import { readReviewGroups } from '../review/groups';
import { acceptReviewGroup } from '../review/accept-group';
import { rejectReviewGroup } from '../review/reject-group';

for (const atomFirst of [false, true])
  test(`mixed atom and text proposals retain independent decisions: ${atomFirst}`, async () => {
    const control =
      '<w:sdt><w:sdtPr><w:tag w:val="scio:1:mention"/><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:r><w:t>Ada</w:t></w:r></w:sdtContent></w:sdt>';
    const loaded = readOoxmlPackage(
      docxFixture(
        '<w:p><w:hyperlink w:anchor="target"><w:r><w:t xml:space="preserve">Lead. </w:t></w:r>' +
          control +
          '<w:r><w:t xml:space="preserve"> Tail.</w:t></w:r></w:hyperlink></w:p>'
      )
    );
    if (!loaded.ok) throw Error(loaded.reason);
    const stored = withPayloadRecord(
      loaded.package,
      createMentionPayload('mention', 'user', 'Ada'),
      'Mention'
    );
    if (!stored.ok) throw Error(stored.reason);
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(stored.pkg),
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
    });
    const context = {
      documentId: 'partial-inline',
      epoch: 'load',
      author: 'AI',
      date: '2026-10-24T12:00:00.000Z',
    };
    const token = () => readEditorRevision(editor, context.documentId, context.epoch)!;
    try {
      const surface = editor.surface!,
        before = surface.session.currentPackage();
      const initial = await readWorkingView(surface.session, token());
      if (!initial.success) throw Error(initial.error);
      const removed = initial.view.projection.text.replace(
        /<scio-mention\b[^>]*\/>|\[Ada\]\(mention:user\)/g,
        ''
      );
      expect(removed).not.toBe(initial.view.projection.text);
      const purePlan = planTargetDocument(
        { pkg: before, revision: token(), projection: initial.view.projection },
        removed
      );
      if (!purePlan.success) throw Error(purePlan.error);
      const pure = previewDirectPlan(before, purePlan.plan);
      if (!pure.success) throw Error(pure.error);
      const result = await recomposeOwnedProposals(
        editor,
        {
          revision: token(),
          targetMarkdown: removed.replace('Lead', 'Changed').replace('Tail', '**Finish**'),
          ownership: [],
        },
        context,
        { payloads: { presence: { proposalId: 'partial-removal' } } }
      );
      if (!result.success) throw Error(result.error);
      const groups = readReviewGroups(surface.session.currentPackage());
      if (!groups.success) throw Error(groups.error);
      const atom = groups.groups.find((group) => group.properties.length > 0)!;
      const text = groups.groups.filter((group) => group.id !== atom.id);
      expect(atom.native).toHaveLength(1);
      expect(text.length).toBeGreaterThan(0);
      const acceptAtom = () => {
        const decided = acceptReviewGroup(editor, atom.id, token(), context);
        if (!decided.success) throw Error(decided.error);
      };
      if (atomFirst) acceptAtom();
      for (const group of text) {
        const decided = rejectReviewGroup(editor, group.id, token(), context);
        if (!decided.success) throw Error(decided.error);
      }
      if (!atomFirst) acceptAtom();
      expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(
        canonicalOoxmlFingerprint(pure.pkg.parts.get(pure.pkg.mainDocumentPart)!)
      );
      const remaining = readReviewGroups(surface.session.currentPackage());
      expect(remaining.success && remaining.groups).toEqual([]);
      const saved = readOoxmlPackage(writeOoxmlPackage(surface.session.currentPackage()));
      if (!saved.ok) throw Error(saved.reason);
      expect(
        canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
      ).toBe(canonicalOoxmlFingerprint(pure.pkg.parts.get(pure.pkg.mainDocumentPart)!));
    } finally {
      editor.destroy();
    }
  });
