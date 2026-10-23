import { prepareInlineImageResources } from '../native/inline-image-resources';
// Test-only collaboration transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  runPropertyEdits,
  type ImageDecodePort,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { caretAt } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { buildBlockIndex } from '../blocks/block-index';
import { applyDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from './plan-target';
const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const decode: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const contentLockPolicy = { allowFormattingInLockedContent: true };
const context = { documentId: 'atom-marks-shared', epoch: 'native', contentLockPolicy };
const sentence =
  'A $x$ [Ada](mention:user) <source title="Paper" url="https://source.test" /> ![Cat](asset:cat) Z';
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw Error(view.error);
  let n = 0;
  const result = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: view.view.projection },
    target,
    { allocateBlockId: () => `shared-mark-${++n}` }
  );
  if (!result.success) throw Error(result.error);
  return result.plan;
}
function apply(part: OoxmlPart, op: TreeDocOp) {
  const result = applyTreeOp(part, op, { contentLockPolicy });
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const writer of [0, 1] as const)
    test(`retained atom formatting keeps peer text and exact histories ${clients}/${writer}`, async () => {
      const seed = createDocxEditor({
        container: document.createElement('div'),
        document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
        contentLockPolicy,
        imageDecodePort: decode,
      });
      let documentBytes: Uint8Array;
      try {
        const birth = await plan(seed, sentence);
        const prepared = await prepareInlineImageResources(
          birth,
          async () => ({ bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 }),
          decode
        );
        if (!prepared.success) throw Error(prepared.error);
        const result = applyDirectPlan(seed, birth, { ...context, images: prepared.images });
        if (!result.success) throw Error(result.error);
        documentBytes = new Uint8Array(await seed.save());
      } finally {
        seed.destroy();
      }
      const pair = await createTwoEditorFixture({
        document: documentBytes,
        clientIds: clients,
        editor: { contentLockPolicy, imageDecodePort: decode },
      });
      try {
        const actor = pair.editors[writer]!,
          peer = pair.editors[1 - writer]!;
        const original = actor.surface!.session.part(),
          paragraphId = actor.surface!.session.paragraphIds()[0]!;
        const length = paragraphTextOf(original, paragraphId)!.length;
        const index = buildBlockIndex(actor.surface!.session.currentPackage());
        const planned = await plan(actor, `**${sentence}**`);
        let marked = original;
        for (const op of runPropertyEdits(
          original,
          paragraphId,
          0,
          length,
          [{ localName: 'b', attributes: { val: '1' } }],
          'proposed'
        ))
          marked = apply(marked, { op: 'setRunProperties', paragraphId, ...op });
        pair.pause();
        const formatted = applyDirectPlan(actor, planned, context);
        if (!formatted.success) throw Error(formatted.error);
        expect(canonicalOoxmlFingerprint(actor.surface!.session.part())).toBe(
          canonicalOoxmlFingerprint(marked)
        );
        const typed = peer.surface!.applyAutomationOps(
          () => [{ op: 'insertText', paragraphId, offset: length, text: ' Peer' }],
          { kind: 'body' },
          [],
          false,
          true
        );
        expect(typed.rejected).toBe(false);
        const peerOnly = apply(original, {
          op: 'insertText',
          paragraphId,
          offset: length,
          text: ' Peer',
        });
        const both = apply(marked, {
          op: 'insertText',
          paragraphId,
          offset: length,
          text: ' Peer',
        });
        pair.reconnect(writer);
        await pair.settle();
        const check = (expected: OoxmlPart) => {
          for (const editor of pair.editors) {
            expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
              canonicalOoxmlFingerprint(expected)
            );
            const current = buildBlockIndex(editor.surface!.session.currentPackage());
            expect(current.diagnostics).toEqual([]);
            expect([...current.byId]).toEqual([...index.byId]);
          }
        };
        check(both);
        expect(actor.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(peerOnly);
        expect(actor.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(both);
        expect(peer.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(marked);
        expect(peer.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(both);
        const offset = length + 5;
        peer.surface!.setSelection({
          anchor: { paragraphId, offset },
          head: { paragraphId, offset },
        });
        await pair.settle();
        const remote = pair.rooms[writer]!.session.remoteSelections().find(
          (selection) => selection.actorId === (writer === 0 ? 'bob' : 'alice')
        );
        expect(remote?.head.nodeId).toBe(paragraphId);
        expect(remote?.head.offset).toBe(offset);
        const a = caretAt(actor.surface!.layout(), { paragraphId, offset }),
          b = caretAt(peer.surface!.layout(), { paragraphId, offset });
        expect(a).not.toBeNull();
        expect(b).not.toBeNull();
        expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
        for (const shared of [false, true]) {
          const reopened = await createTwoEditorFixture({
            document: new Uint8Array(await actor.save()),
            editor: { contentLockPolicy, imageDecodePort: decode },
            ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[writer]!) } : {}),
          });
          try {
            for (const editor of reopened.editors) {
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
                canonicalOoxmlFingerprint(both)
              );
              expect(
                [...buildBlockIndex(editor.surface!.session.currentPackage()).byId].map(
                  ([id, block]) => [id, block.payload]
                )
              ).toEqual([...index.byId].map(([id, block]) => [id, block.payload]));
              expect((await plan(editor, `**${sentence} Peer**`)).operations).toEqual([]);
            }
          } finally {
            reopened.destroy();
          }
        }
      } finally {
        pair.destroy();
      }
    });
