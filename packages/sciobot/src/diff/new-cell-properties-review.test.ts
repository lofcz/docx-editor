import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  serializeOoxmlPart,
  type OoxmlPart,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { table } from '../../test/support/target-table-fixture';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { trackTableRowPlan } from '../review/track-table-rows';
import { trackTableColumnPlan } from '../review/track-table-columns';
import { previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';
// Independent native allocation produces fresh paragraph identity attributes, as in R08 oracles.
const allocatedXml = (part: OoxmlPart) =>
  serializeOoxmlPart(part).replace(/ w14:(paraId|textId)="[0-9A-F]+"/g, '');
const context = { documentId: 'new-cell-properties-review', epoch: '1' };
for (const kind of ['row', 'column'] as const)
  test(`tracked ${kind} owns new properties through exact rejection and direct-equivalent acceptance`, async () => {
    const input = await targetDocumentInput(table([['A']]));
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(input.pkg),
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
    });
    try {
      const session = editor.surface!.session,
        revision = readEditorRevision(editor, context.documentId, context.epoch)!;
      const view = await readWorkingView(session, revision, { tableSyntax: 'html' });
      if (!view.success) throw new Error(view.error);
      const original = session.currentPackage(),
        before = canonicalOoxmlFingerprint(session.part());
      const target =
        '<table>\n<tr>\n<td>\n\nA\n\n</td>\n' +
        (kind === 'row' ? '</tr>\n<tr>\n' : '') +
        '<td style="background-color: #123456; border: 1pt solid red; text-align: right; vertical-align: bottom">\n\nNew\n\n</td>\n</tr>\n</table>';
      const planned = planTargetDocument(
        { pkg: original, revision, projection: view.view.projection },
        target
      );
      if (!planned.success) throw new Error(planned.error);
      const expected = previewDirectPlan(original, planned.plan);
      if (!expected.success) throw new Error(expected.error);
      const applied = (kind === 'row' ? trackTableRowPlan : trackTableColumnPlan)(
        editor,
        planned.plan,
        { ...context, author: 'ScioBot', date: '2026-10-22T08:00:00Z' }
      );
      if (!applied.success) throw new Error(applied.error);
      expect(editor.getReviewItems()).toHaveLength(kind === 'row' ? 1 : 2);
      expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
      expect(allocatedXml(session.part())).toBe(
        allocatedXml(expected.pkg.parts.get(expected.pkg.mainDocumentPart)!)
      );
      const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!saved.ok) throw new Error(saved.reason);
      expect(
        canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
      ).toBe(canonicalOoxmlFingerprint(session.part()));
    } finally {
      editor.destroy();
    }
  });
