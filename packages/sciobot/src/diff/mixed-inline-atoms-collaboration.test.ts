// Evaluation-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  type ImageDecodePort,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { caretAt } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { buildBlockIndex } from '../blocks/block-index';
import { applyDirectPlan } from '../ai/apply-direct';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { planTargetDocument } from './plan-target';
const context = { documentId: 'mixed-atoms', epoch: 'native' };
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
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const source = await readWorkingView(editor.surface!.session, revision);
  if (!source.success) throw Error(source.error);
  const planned = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: source.view.projection },
    target
  );
  if (!planned.success) throw Error(planned.error);
  return planned.plan;
}
async function apply(editor: DocxEditorInstance, target: string, expected: readonly string[]) {
  const planned = await plan(editor, target),
    resolved: string[] = [];
  const prepared = await prepareInlineImageResources(
    planned,
    async (reference) => {
      resolved.push(reference);
      return { bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 };
    },
    decode
  );
  if (!prepared.success) throw Error(prepared.error);
  expect(resolved).toEqual([...expected]);
  const result = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
  if (!result.success) throw Error(result.error);
}
function snapshot(editor: DocxEditorInstance) {
  const pkg = editor.surface!.session.currentPackage();
  const index = buildBlockIndex(pkg);
  expect(index.diagnostics).toEqual([]);
  return {
    word: canonicalOoxmlFingerprint(editor.surface!.session.part()),
    payloads: [...index.byId]
      .map(([id, block]) => [id, block.payload])
      .sort(([a], [b]) => String(a).localeCompare(String(b))),
    media: [...pkg.partBytes]
      .filter(([name]) => /(^|\/)media\//.test(name))
      .sort(([a], [b]) => a.localeCompare(b)),
  };
}
function typed(part: OoxmlPart, paragraphId: string) {
  const result = applyTreeOp(part, { op: 'insertText', paragraphId, offset: 1, text: ' peer' });
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
for (const retained of ['math', 'image', 'mention', 'source'] as const)
  for (const side of ['before', 'after'] as const)
    for (const clients of [
      [41, 73],
      [73, 41],
    ] as const)
      for (const writer of [0, 1] as const)
        test(`mixed linked atoms preserve native peer histories: ${retained}/${side}/${clients}/${writer}`, async () => {
          const atom =
            retained === 'math'
              ? '$x$'
              : retained === 'image'
                ? '![Cat](asset:cat "Photo")'
                : retained === 'mention'
                  ? '<scio-mention key="user" value="Ada" />'
                  : '<source title="Paper" url="https://source.test" />';
          const fresh = retained === 'image' ? '$x$' : '![Cat](asset:cat "Photo")';
          const wrap = (value: string) => `[<code>${value}</code>](https://outer.test "Read")`;
          const seedTarget = `Keep. ${wrap(atom)} Z`;
          const target = `Keep. ${wrap(side === 'after' ? `${atom} ${fresh}` : `${fresh} ${atom}`)} Z`;
          const seed = createDocxEditor({
            container: document.createElement('div'),
            imageDecodePort: decode,
            document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
          });
          let documentBytes: Uint8Array;
          try {
            await apply(seed, seedTarget, retained === 'image' ? ['asset:cat'] : []);
            documentBytes = new Uint8Array(await seed.save());
          } finally {
            seed.destroy();
          }
          const fixture = {
            document: documentBytes,
            clientIds: clients,
            editor: { imageDecodePort: decode },
          };
          const pair = await createTwoEditorFixture(fixture);
          try {
            const actor = pair.editors[writer]!,
              peer = pair.editors[1 - writer]!;
            const original = actor.surface!.session.part(),
              paragraphId = actor.surface!.session.paragraphIds()[0]!;
            const before = snapshot(actor);
            const oldBlock = [
              ...buildBlockIndex(actor.surface!.session.currentPackage()).byId.values(),
            ][0]!;
            const oldWord = canonicalOoxmlFingerprint(findNode(original, oldBlock.controlId)!);
            pair.pause();
            await apply(actor, target, retained === 'image' ? [] : ['asset:cat']);
            const marked = actor.surface!.session.part(),
              after = snapshot(actor);
            expect(canonicalOoxmlFingerprint(findNode(marked, oldBlock.controlId)!)).toBe(oldWord);
            expect(
              buildBlockIndex(actor.surface!.session.currentPackage()).byId.get(oldBlock.id)
                ?.payload
            ).toEqual(oldBlock.payload);
            expect(
              peer.surface!.applyAutomationOps(
                () => [{ op: 'insertText', paragraphId, offset: 1, text: ' peer' }],
                { kind: 'body' },
                [],
                false,
                true
              ).rejected
            ).toBe(false);
            const peerOnly = {
              ...before,
              word: canonicalOoxmlFingerprint(typed(original, paragraphId)),
            };
            const both = { ...after, word: canonicalOoxmlFingerprint(typed(marked, paragraphId)) };
            pair.reconnect(writer);
            await pair.settle();
            const check = async (expected: ReturnType<typeof snapshot>, markdown: string) => {
              for (const editor of pair.editors) {
                expect(snapshot(editor)).toEqual(expected);
                expect((await plan(editor, markdown)).operations).toEqual([]);
                expect(
                  canonicalOoxmlFingerprint(
                    findNode(editor.surface!.session.part(), oldBlock.controlId)!
                  )
                ).toBe(oldWord);
              }
              const [a, b] = await pair.savedDigests();
              expect(a).toEqual(b);
            };
            const peerTarget = target.replace('Keep.', 'K peereep.');
            // Native insertion at offset one yields this exact prefix.
            const withPeer = (value: string) => value.replace('Keep.', 'K peereep.');
            expect(
              paragraphTextOf(pair.editors[0]!.surface!.session.part(), paragraphId)?.startsWith(
                'K peereep.'
              )
            ).toBe(true);
            await check(both, peerTarget);
            expect(actor.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(peerOnly, withPeer(seedTarget));
            expect(actor.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(both, withPeer(target));
            expect(peer.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(after, target);
            expect(peer.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(both, withPeer(target));
            const offset = paragraphTextOf(actor.surface!.session.part(), paragraphId)!.length;
            peer.surface!.setSelection({
              anchor: { paragraphId, offset },
              head: { paragraphId, offset },
            });
            await pair.settle();
            const geometry = pair.editors.map((editor) =>
              caretAt(editor.surface!.layout(), { paragraphId, offset })
            );
            expect(geometry[0]).not.toBeNull();
            expect(geometry[1]).not.toBeNull();
            expect({ x: geometry[0]!.x, y: geometry[0]!.y }).toEqual({
              x: geometry[1]!.x,
              y: geometry[1]!.y,
            });
            const remote = pair.rooms[writer]!.session.remoteSelections().find(
              (selection) => selection.actorId === (writer === 0 ? 'bob' : 'alice')
            );
            expect(remote?.head.nodeId).toBe(paragraphId);
            expect(remote?.head.offset).toBe(offset);
            for (const shared of [false, true]) {
              const reopened = await createTwoEditorFixture({
                ...fixture,
                clientIds: undefined,
                document: new Uint8Array(await actor.save()),
                ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[writer]!) } : {}),
              });
              try {
                await reopened.settle();
                for (const editor of reopened.editors) {
                  expect(snapshot(editor)).toEqual(both);
                  expect((await plan(editor, withPeer(target))).operations).toEqual([]);
                }
              } finally {
                reopened.destroy();
              }
            }
          } finally {
            pair.destroy();
          }
        });
