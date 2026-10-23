// The test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { propertyEditor } from '../../test/support/property-execution';
import { readEditorRevision } from '../projection/revision-cache';
import { planRunFormat } from '../diff/run-format';
import { planParagraphFormat } from '../diff/paragraph-format';
import { applyDirectPlan } from '../ai/apply-direct';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackFormatPlan } from './track-format';

const context = {
  documentId: 'format-peer',
  epoch: 'mount',
  author: 'Alice',
  date: '2026-10-23T00:00:00Z',
};
const documentBytes = () =>
  docxFixture(
    '<w:p><w:pPr><w:spacing w:after="120"/><w:customKeep w:val="yes"/></w:pPr><w:r><w:t>Before </w:t></w:r><w:r><w:rPr><w:color w:val="123456"/></w:rPr><w:t>Item</w:t></w:r><w:r><w:t> After</w:t></w:r></w:p>'
  );
function plan(editor: DocxEditorInstance, heading: boolean) {
  const session = editor.surface!.session;
  const source = {
    part: session.part(),
    paragraphId: session.paragraphIds()[0]!,
    revision: readEditorRevision(editor, context.documentId, context.epoch)!,
  };
  const result = heading
    ? planParagraphFormat(source, {
        properties: [{ localName: 'pStyle', attributes: { val: 'Heading2' } }],
      })
    : planRunFormat(
        { ...source, start: 7, end: 11, displayMode: 'proposed' },
        { properties: [{ localName: 'b' }] }
      );
  if (!result.success) throw Error(result.error);
  return result.plan;
}
function peerEdit(editor: DocxEditorInstance, remove: boolean) {
  const session = editor.surface!.session,
    paragraphId = session.paragraphIds()[0]!;
  const end = paragraphTextOf(session.part(), paragraphId)!.length;
  expect(
    editor.surface!.applyAutomationOps(() =>
      remove
        ? [{ op: 'deleteText', paragraphId, start: 8, end: 9 }]
        : [{ op: 'insertText', paragraphId, offset: end, text: ' Peer' }]
    ).rejected
  ).toBe(false);
}
const fingerprint = (editor: DocxEditorInstance) =>
  canonicalOoxmlFingerprint(editor.surface!.session.part());
for (const heading of [false, true])
  for (const remove of [false, true])
    for (const receiver of [0, 1] as const)
      test(`${heading ? 'heading' : 'run'} formatting retains offline peer ${remove ? 'deletion' : 'suffix'}, receiver ${receiver}`, async () => {
        const bytes = documentBytes();
        const opened = readOoxmlPackage(bytes);
        if (!opened.ok) throw Error(opened.reason);
        const pair = await createTwoEditorFixture({
          document: bytes,
          documentId: context.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        const original = propertyEditor(opened.package),
          proposed = propertyEditor(opened.package);
        try {
          const direct = applyDirectPlan(proposed, plan(proposed, heading), context);
          if (!direct.success) throw Error(direct.error);
          peerEdit(original, remove);
          peerEdit(proposed, remove);
          const before = fingerprint(original),
            after = fingerprint(proposed);
          const [alice, bob] = pair.editors;
          pair.pause();
          const tracked = trackFormatPlan(alice!, plan(alice!, heading), context);
          if (!tracked.success) throw Error(tracked.error);
          peerEdit(bob!, remove);
          pair.reconnect(receiver);
          await pair.settle();
          const pending = fingerprint(alice!);
          expect(fingerprint(bob!)).toBe(pending);
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(fingerprint(editor)).toBe(before);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(fingerprint(editor)).toBe(pending);
          const paragraphId = bob!.surface!.session.paragraphIds()[0]!;
          const offset = paragraphTextOf(bob!.surface!.session.part(), paragraphId)!.length;
          bob!.surface!.setSelection({
            anchor: { paragraphId, offset },
            head: { paragraphId, offset },
          });
          await pair.settle();
          const presence = pair.rooms[0]!.session.remoteSelections().find(
            (entry) => entry.actorId === 'bob'
          );
          expect(presence?.head.nodeId).toBe(paragraphId);
          expect(presence?.head.offset).toBe(offset);
          const carets = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId, offset })
          );
          expect(carets[0]).not.toBeNull();
          expect(carets[1]).not.toBeNull();
          expect({ x: carets[0]!.x, y: carets[0]!.y }).toEqual({
            x: carets[1]!.x,
            y: carets[1]!.y,
          });
          for (const action of ['accept', 'reject'] as const) {
            expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors) {
              expect(fingerprint(editor)).toBe(action === 'accept' ? after : before);
              const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
              if (!saved.ok) throw Error(saved.reason);
              expect(
                canonicalOoxmlFingerprint(saved.package.parts.get(saved.package.mainDocumentPart)!)
              ).toBe(action === 'accept' ? after : before);
            }
            const digests = await pair.savedDigests();
            expect(digests[0]).toEqual(digests[1]);
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors) expect(fingerprint(editor)).toBe(pending);
            expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(fingerprint(editor)).toBe(action === 'accept' ? after : before);
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
          original.destroy();
          proposed.destroy();
        }
      });
