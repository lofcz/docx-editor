// Evaluation-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  paragraphOffsetIndex,
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
function typed(part: OoxmlPart, paragraphId: string, offset = 1) {
  const result = applyTreeOp(part, { op: 'insertText', paragraphId, offset, text: ' peer' });
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
for (const retained of ['math', 'image', 'mention', 'source'] as const)
  for (const side of ['prefix', 'middle', 'suffix'] as const)
    for (const clients of [
      [41, 73],
      [73, 41],
    ] as const)
      for (const writer of [0, 1] as const)
        test(`identical linked atoms preserve both native identities and peer histories: ${retained}/${side}/${clients}/${writer}`, async () => {
          const atom =
            retained === 'math'
              ? '$x$'
              : retained === 'image'
                ? '![Cat](asset:cat "Photo")'
                : retained === 'mention'
                  ? '<scio-mention key="user" value="Ada" />'
                  : '<source title="Paper" url="https://source.test" />';
          const wrap = (value: string) => `[<code>${value}</code>](https://outer.test "Read")`;
          const body = `Lead. ${atom} bridge ${atom} Tail.`;
          const seedTarget = wrap(body);
          const target = wrap(
            side === 'prefix'
              ? `${atom} Fresh ${body}`
              : side === 'middle'
                ? `Lead. ${atom} ${atom} bridge ${atom} Tail.`
                : `${body} Fresh ${atom}`
          );
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
            const paragraph = findNode(original, paragraphId)!;
            if (paragraph.kind !== 'paragraph') throw Error('Missing paragraph');
            const offsets = paragraphOffsetIndex(paragraph);
            const oldBlocks = [
              ...buildBlockIndex(actor.surface!.session.currentPackage()).byId.values(),
            ]
              .map((block) => ({
                block,
                span: offsets.spanOf(block.controlId)!,
                word: canonicalOoxmlFingerprint(findNode(original, block.controlId)!),
              }))
              .sort((a, b) => a.span.start - b.span.start);
            expect(oldBlocks).toHaveLength(2);
            const width = oldBlocks[0]!.span.end - oldBlocks[0]!.span.start;
            const prefix = side === 'prefix' ? width + ' Fresh '.length : 0;
            pair.pause();
            await apply(actor, target, retained === 'image' ? ['asset:cat'] : []);
            const marked = actor.surface!.session.part(),
              after = snapshot(actor);
            for (const { block, word } of oldBlocks) {
              expect(canonicalOoxmlFingerprint(findNode(marked, block.controlId)!)).toBe(word);
              expect(
                buildBlockIndex(actor.surface!.session.currentPackage()).byId.get(block.id)?.payload
              ).toEqual(block.payload);
            }
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
            const both = {
              ...after,
              word: canonicalOoxmlFingerprint(typed(marked, paragraphId, 1 + prefix)),
            };
            pair.reconnect(writer);
            await pair.settle();
            const check = async (expected: ReturnType<typeof snapshot>, markdown: string) => {
              for (const editor of pair.editors) {
                expect(snapshot(editor)).toEqual(expected);
                expect((await plan(editor, markdown)).operations).toEqual([]);
                for (const { block, word } of oldBlocks)
                  expect(
                    canonicalOoxmlFingerprint(
                      findNode(editor.surface!.session.part(), block.controlId)!
                    )
                  ).toBe(word);
              }
              const [a, b] = await pair.savedDigests();
              expect(a).toEqual(b);
            };
            const withPeer = (value: string) => value.replace('Lead.', 'L peeread.');
            const peerTarget = withPeer(target);
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
            const deleted = (part: OoxmlPart, shift: number) => {
              const result = applyTreeOp(part, {
                op: 'deleteText',
                paragraphId,
                start: 4 + shift,
                end: 5 + shift,
              });
              if (!result.ok) throw Error(result.reason);
              return result.part;
            };
            const deletedBoth = {
              ...both,
              word: canonicalOoxmlFingerprint(
                deleted(typed(marked, paragraphId, 1 + prefix), prefix)
              ),
            };
            const deletedPeer = {
              ...peerOnly,
              word: canonicalOoxmlFingerprint(deleted(typed(original, paragraphId), 0)),
            };
            const withDeletedPeer = (value: string) => value.replace('Lead.', 'L peread.');
            pair.pause();
            expect(actor.exec({ type: 'undo' }).ok).toBe(true);
            expect(
              peer.surface!.applyAutomationOps(
                () => [{ op: 'deleteText', paragraphId, start: 4 + prefix, end: 5 + prefix }],
                { kind: 'body' },
                [],
                false,
                true
              ).rejected
            ).toBe(false);
            pair.reconnect(writer);
            await pair.settle();
            await check(deletedPeer, withDeletedPeer(seedTarget));
            expect(actor.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(deletedBoth, withDeletedPeer(target));
            expect(peer.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(both, withPeer(target));
            expect(peer.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(deletedBoth, withDeletedPeer(target));
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
                  expect(snapshot(editor)).toEqual(deletedBoth);
                  expect((await plan(editor, withDeletedPeer(target))).operations).toEqual([]);
                }
              } finally {
                reopened.destroy();
              }
            }
          } finally {
            pair.destroy();
          }
        });
