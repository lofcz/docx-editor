import { prepareInlineImageResources } from '../native/inline-image-resources';
// Test-only collaboration transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  ensureHyperlinkRelationship,
  TreePackageStore,
  replaceNode,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
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
    { allocateBlockId: () => `shared-link-${++n}` }
  );
  if (!result.success) throw Error(result.error);
  return result.plan;
}
function apply(part: OoxmlPart, op: TreeDocOp) {
  const result = applyTreeOp(part, op, { contentLockPolicy });
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
function descendants(node: OoxmlNode): OoxmlNode[] {
  return [node, ...(node.kind === 'textValue' ? [] : node.children.flatMap(descendants))];
}
for (const [mode, compound] of (['add', 'retarget', 'remove'] as const).flatMap((mode) =>
  [false, true].map((compound) => [mode, compound] as const)
))
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const writer of [0, 1] as const)
      test(`retained atom ${mode} links keep peer text, deletion and exact histories ${clients}/${writer}/compound=${compound}`, async () => {
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
          if (mode !== 'add') {
            const part = seed.surface!.session.part(),
              paragraphId = seed.surface!.session.paragraphIds()[0]!;
            const result = seed.surface!.applyAutomationOps(
              (relate) => [
                {
                  op: 'insertHyperlink',
                  paragraphId,
                  start: 0,
                  end: paragraphTextOf(part, paragraphId)!.length,
                  relationshipId: relate('https://old.test')!,
                  tooltip: 'Original',
                },
              ],
              { kind: 'body' },
              [
                (pkg) => {
                  const main = pkg.parts.get(pkg.mainDocumentPart)!;
                  const link = descendants(main.root).find((node) => node.kind === 'hyperlink')!;
                  if (!link) throw Error('Missing imported link');
                  const replaced = replaceNode(main, link.id, {
                    ...link,
                    attributes: [
                      ...link.attributes.filter(
                        (attribute) => !['history', 'tgtFrame'].includes(attribute.localName)
                      ),
                      ...['history', 'tgtFrame'].map((localName, index) => ({
                        kind: 'genericExtension' as const,
                        namespaceUri: WML_NAMESPACE_URI,
                        prefix: 'w',
                        localName,
                        value: index === 0 ? '1' : '_blank',
                      })),
                    ],
                  });
                  if (!replaced.ok) throw Error(JSON.stringify(replaced.issues));
                  return withPart(pkg, replaced.part);
                },
              ],
              false,
              true,
              'after'
            );
            expect(result.rejected).toBe(false);
          }
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
          const linkedSentence = sentence.replace(
            '[Ada](mention:user)',
            '<scio-mention key="user" value="Ada" />'
          );
          const target =
            mode === 'remove' ? sentence : `[${linkedSentence}](https://new.test "Changed")`;
          const planned = await plan(actor, target);
          const oldLink = descendants(original.root).find((node) => node.kind === 'hyperlink');
          const oracleStore = new TreePackageStore(
            actor.surface!.session.currentPackage(),
            original
          );
          let relationshipId = '';
          oracleStore.transact(
            { kind: 'body' },
            (ctx) => {
              ctx.applyPackage((pkg) => {
                const made = ensureHyperlinkRelationship(pkg, 'https://new.test');
                if (!made) throw Error('Invalid native target');
                relationshipId = made.relationshipId;
                return made.pkg;
              });
            },
            { actorId: writer === 0 ? 'alice' : 'bob' }
          );
          const marked = apply(
            original,
            mode === 'remove'
              ? { op: 'removeHyperlink', linkId: oldLink!.id }
              : mode === 'retarget'
                ? {
                    op: 'setHyperlinkTarget',
                    linkId: oldLink!.id,
                    relationshipId,
                    tooltip: 'Changed',
                  }
                : {
                    op: 'insertHyperlink',
                    paragraphId,
                    start: 0,
                    end: length,
                    relationshipId,
                    tooltip: 'Changed',
                  }
          );
          pair.pause();
          const formatted = applyDirectPlan(actor, planned, context);
          if (!formatted.success) throw Error(formatted.error);
          expect(canonicalOoxmlFingerprint(actor.surface!.session.part())).toBe(
            canonicalOoxmlFingerprint(marked)
          );
          const paragraphOp: TreeDocOp = {
            op: 'setParagraphProperties',
            paragraphId,
            properties: [{ localName: 'spacing', attributes: { after: '120' } }],
          };
          const withParagraphProperty = (part: OoxmlPart) =>
            compound ? apply(part, paragraphOp) : part;
          const typed = peer.surface!.applyAutomationOps(
            () => [
              { op: 'insertText', paragraphId, offset: 1, text: ' peer' },
              ...(compound ? [paragraphOp] : []),
            ],
            { kind: 'body' },
            [],
            false,
            true
          );
          expect(typed.rejected).toBe(false);
          let peerOnly = apply(original, {
            op: 'insertText',
            paragraphId,
            offset: 1,
            text: ' peer',
          });
          peerOnly = withParagraphProperty(peerOnly);
          expect(canonicalOoxmlFingerprint(peer.surface!.session.part())).toBe(
            canonicalOoxmlFingerprint(peerOnly)
          );
          let both = apply(marked, {
            op: 'insertText',
            paragraphId,
            offset: 1,
            text: ' peer',
          });
          both = withParagraphProperty(both);
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
          const deleteOp = { op: 'deleteText' as const, paragraphId, start: 3, end: 4 };
          expect(
            peer.surface!.applyAutomationOps(() => [deleteOp], { kind: 'body' }, [], false, true)
              .rejected
          ).toBe(false);
          await pair.settle();
          const deletedBoth = apply(both, deleteOp),
            deletedPeer = apply(peerOnly, deleteOp);
          check(deletedBoth);
          expect(actor.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(deletedPeer);
          expect(actor.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(deletedBoth);
          expect(peer.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(both);
          expect(peer.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(deletedBoth);
          const offset = length + 4;
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
                  canonicalOoxmlFingerprint(deletedBoth)
                );
                expect(
                  [...buildBlockIndex(editor.surface!.session.currentPackage()).byId].map(
                    ([id, block]) => [id, block.payload]
                  )
                ).toEqual([...index.byId].map(([id, block]) => [id, block.payload]));
                expect(
                  (
                    await plan(
                      editor,
                      mode === 'remove'
                        ? sentence.replace('A ', 'A per ')
                        : `[${linkedSentence.replace('A ', 'A per ')}](https://new.test "Changed")`
                    )
                  ).operations
                ).toEqual([]);
              }
            } finally {
              reopened.destroy();
            }
          }
        } finally {
          pair.destroy();
        }
      });
