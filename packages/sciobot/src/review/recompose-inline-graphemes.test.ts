import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { buildBlockIndex } from '../blocks/block-index';
import { createMentionPayload } from '../blocks/mention-codec';
import { withPayloadRecord } from '../blocks/payload-opc-codec';
import { createEditPlan } from '../diff/plan';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { replaceOwnedProposals } from '../review/replace-owned-proposals';

for (const insertion of [false, true])
  test(`atom removal cannot make a canonical text edit split a proposed grapheme: ${insertion}`, () => {
    const control =
      '<w:sdt><w:sdtPr><w:tag w:val="scio:1:mention"/></w:sdtPr><w:sdtContent><w:r><w:t>Ada</w:t></w:r></w:sdtContent></w:sdt>';
    const loaded = readOoxmlPackage(
      docxFixture('<w:p><w:r><w:t>a</w:t></w:r>' + control + '<w:r><w:t>\u0301z</w:t></w:r></w:p>')
    );
    if (!loaded.ok) throw Error(loaded.reason);
    const payload = createMentionPayload('mention', 'user', 'Ada');
    const stored = withPayloadRecord(loaded.package, payload, 'Mention');
    if (!stored.ok) throw Error(stored.reason);
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(stored.pkg),
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
    });
    const context = {
      documentId: 'inline-grapheme',
      epoch: 'load',
      author: 'AI',
      date: '2026-10-24T12:00:00.000Z',
    };
    try {
      const surface = editor.surface!,
        before = surface.session.currentPackage();
      const token = readEditorRevision(editor, context.documentId, context.epoch)!;
      const block = buildBlockIndex(before).byId.get(payload.id)!;
      const paragraphId = surface.session.paragraphIds()[0]!;
      const plan = createEditPlan(token, [
        {
          kind: 'inline-remove',
          partName: before.mainDocumentPart,
          paragraphId,
          blockId: payload.id,
          controlId: block.controlId,
          before: payload,
          fingerprint: canonicalOoxmlFingerprint(
            findNode(surface.session.part(), block.controlId)!
          ),
        },
        {
          kind: 'text',
          partName: before.mainDocumentPart,
          paragraphId,
          start: insertion ? 1 : 0,
          end: 1,
          before: insertion ? '' : 'a',
          after: 'x',
        },
      ]);
      if (!plan.success) throw Error(plan.error);
      // Visible a + combining accent becomes one grapheme after atom removal.
      // Raw source positions are individually valid; the proposed coordinate model must refuse them.
      const result = replaceOwnedProposals(editor, [], plan.plan, context, {
        presence: { proposalId: 'grapheme-removal' },
      });
      expect(result.success).toBe(false);
      if (!result.success) expect(result.code).toBe('ambiguous-target');
      expect(surface.session.currentPackage()).toBe(before);
      expect(readEditorRevision(editor, context.documentId, context.epoch)).toEqual(token);
      expect(surface.session.canUndo()).toBe(false);
    } finally {
      editor.destroy();
    }
  });
