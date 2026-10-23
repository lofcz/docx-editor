import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { prepareNativeDocument } from '../native/prepare-document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { buildBlockIndex } from '../blocks/block-index';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';
const context = { documentId: 'code-atoms-peer', epoch: 'native' };
const document = () =>
  docxFixture(
    '<w:p><w:r><w:t xml:space="preserve">Keep. </w:t></w:r><w:hyperlink r:id="existing" w:tooltip="Read" w:history="1"><w:r><w:rPr><w:color w:val="112233"/></w:rPr><w:t xml:space="preserve"> and </w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> End.</w:t></w:r></w:p>',
    {
      relationships:
        '<Relationship Id="existing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://outer.test" TargetMode="External"/>',
    }
  );
const word = (part: OoxmlPart) => canonicalOoxmlFingerprint(part);
for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const prefix of [false, true])
      test(`code-marked atom birth preserves peer text and imported native link through both histories: ${clientIds}/${receiver}/prefix=${prefix}`, async () => {
        const prepared = prepareNativeDocument(document());
        if (!prepared.success) throw Error(prepared.error);
        const fixture = { document: prepared.document, clientIds };
        const pair = await createTwoEditorFixture(fixture);
        try {
          const [alice, bob] = pair.editors;
          const initial = alice!.surface!.session.part();
          const initialWord = word(initial);
          const paragraphId = alice!.surface!.session.paragraphIds()[0]!;
          const initialOffset = paragraphTextOf(initial, paragraphId)!.indexOf('and') + 1;
          const expected = applyTreeOp(initial, {
            op: 'insertText',
            paragraphId,
            offset: initialOffset,
            text: ' peer ',
          });
          if (!expected.ok) throw Error(expected.reason);
          const revision = readEditorRevision(alice!, context.documentId, context.epoch)!;
          const source = await readWorkingView(alice!.surface!.session, revision);
          if (!source.success) throw Error(source.error);
          const target = `Keep. [<code>${prefix ? 'lead ' : ''}**$x$** and <source title="Paper" url="https://source.test" /></code>](https://outer.test "Read") End.`;
          let next = 0;
          const planned = planTargetDocument(
            {
              revision,
              pkg: alice!.surface!.session.currentPackage(),
              projection: source.view.projection,
            },
            target,
            { allocateBlockId: () => `code-peer-${++next}` }
          );
          if (!planned.success) throw Error(planned.error);
          const applied = applyDirectPlan(alice!, planned.plan, context);
          if (!applied.success) throw Error(applied.error);
          await pair.settle();
          pair.pause();
          const offset =
            paragraphTextOf(bob!.surface!.session.part(), paragraphId)!.indexOf('and') + 1;
          expect(
            bob!.surface!.applyAutomationOps(
              () => [{ op: 'insertText', paragraphId, offset, text: ' peer ' }],
              { kind: 'body' },
              [],
              false,
              true
            ).rejected
          ).toBe(false);
          const authoredPeerWord = word(bob!.surface!.session.part());
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          pair.reconnect(receiver);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(word(editor.surface!.session.part())).toBe(word(expected.part));
            expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(0);
          }
          for (const room of [false, true]) {
            const reopened = await createTwoEditorFixture({
              ...fixture,
              clientIds: undefined,
              document: new Uint8Array(await alice!.save()),
              ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
            });
            try {
              await reopened.settle();
              for (const editor of reopened.editors)
                expect(word(editor.surface!.session.part())).toBe(word(expected.part));
            } finally {
              reopened.destroy();
            }
          }
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(word(editor.surface!.session.part())).toBe(initialWord);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(word(editor.surface!.session.part())).toBe(word(expected.part));
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            expect(word(editor.surface!.session.part())).toBe(authoredPeerWord);
            expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(2);
          }
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
        } finally {
          pair.destroy();
        }
      }, 60_000);
