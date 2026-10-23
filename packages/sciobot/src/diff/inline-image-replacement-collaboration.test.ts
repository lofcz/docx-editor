// Evaluation-only collaboration transport uses the EigenPal Pro License.
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  resolveImageRelationship,
  type ImageDecodePort,
  type OoxmlNode,
  type OoxmlPart,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { buildBlockIndex } from '../blocks/block-index';
import { planTargetDocument } from './plan-target';
import { serializeEditPlan, decodeEditPlan } from './plan';
import { applyDirectPlan } from '../ai/apply-direct';
import { prepareInlineImageResources } from '../native/inline-image-resources';
const context = { documentId: 'image-replacement', epoch: 'native' };
const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
function labeledPng(label: string) {
  const data = new TextEncoder().encode('Comment\0' + label);
  const chunk = new Uint8Array(12 + data.length);
  new DataView(chunk.buffer).setUint32(0, data.length);
  chunk.set([116, 69, 88, 116], 4);
  chunk.set(data, 8);
  let crc = 0xffffffff;
  for (const byte of chunk.subarray(4, -4)) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  }
  new DataView(chunk.buffer).setUint32(chunk.length - 4, (crc ^ 0xffffffff) >>> 0);
  const result = new Uint8Array(png.length + chunk.length);
  result.set(png.subarray(0, 33));
  result.set(chunk, 33);
  result.set(png.subarray(33), 33 + chunk.length);

  return result;
}
const replacement = labeledPng('Alpha'),
  betaPng = labeledPng('Beta');
const decode: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes]
    .filter(([name]) => /(^|\/)media\//.test(name))
    .sort(([a], [b]) => a.localeCompare(b));
const descendants = (node: OoxmlNode): OoxmlNode[] => [
  node,
  ...(node.kind === 'textValue' ? [] : node.children.flatMap(descendants)),
];
function op(part: OoxmlPart, value: TreeDocOp): OoxmlPart {
  const applied = applyTreeOp(part, value);
  if (!applied.ok) throw Error(applied.reason);
  return applied.part;
}
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const source = await readWorkingView(editor.surface!.session, revision);
  if (!source.success) throw Error(source.error);
  const planned = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: source.view.projection },
    target
  );
  if (!planned.success) throw Error(planned.error);
  const decoded = decodeEditPlan(serializeEditPlan(planned.plan));
  expect(decoded).toEqual({ success: true, plan: planned.plan });
  if (!decoded.success) throw Error(decoded.error);
  return decoded.plan;
}
async function apply(editor: DocxEditorInstance, target: string) {
  const planned = await plan(editor, target),
    resolved: string[] = [];
  const prepared = await prepareInlineImageResources(
    planned,
    async (reference) => {
      resolved.push(reference);
      return {
        bytes: reference === 'asset:cat' ? png : reference === 'asset:beta' ? betaPng : replacement,
        mime: 'image/png',
        widthPoints: reference === 'asset:cat' ? 12 : 40,
        heightPoints: reference === 'asset:cat' ? 8 : 30,
      };
    },
    decode
  );
  if (!prepared.success) throw Error(prepared.error);
  const result = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
  if (!result.success) throw Error(result.error);
  return { planned, resolved };
}
function snapshot(editor: DocxEditorInstance) {
  const pkg = editor.surface!.session.currentPackage(),
    index = buildBlockIndex(pkg);
  expect(index.diagnostics).toEqual([]);
  return {
    pkg,
    part: editor.surface!.session.part(),
    payloads: new Map([...index.byId].map(([id, b]) => [id, b.payload])),
    media: media(pkg),
  };
}
function overlayMedia(
  base: ReturnType<typeof media>,
  states: readonly ReturnType<typeof snapshot>[]
) {
  const result = new Map(base);
  const original = new Map(base);
  for (const state of states) {
    const changed = new Map(state.media);
    for (const [name] of base) if (!changed.has(name)) result.delete(name);
    for (const [name, bytes] of changed) {
      const previous = original.get(name);
      if (
        !previous ||
        previous.length !== bytes.length ||
        previous.some((byte, index) => byte !== bytes[index])
      )
        result.set(name, bytes);
    }
  }
  return [...result].sort(([a], [b]) => a.localeCompare(b));
}
for (const conflict of [false, true])
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const writer of [0, 1] as const)
      for (const receiver of [0, 1] as const)
        test(`image replacements retain coherent native resources through both peer histories: conflict=${conflict}/${clients}/${writer}/${receiver}`, async () => {
          const image = '[<code>![Cat](asset:cat "Photo")</code>](https://outer.test "Read")';
          const seedTarget = `Lead. ${image} bridge ${image} Tail.`;
          const alpha = '[<code>![Alpha](asset:alpha "Alpha")</code>](https://outer.test "Read")';
          const beta = '[<code>![Beta](asset:beta "Beta")</code>](https://outer.test "Read")';
          const seed = createDocxEditor({
            container: document.createElement('div'),
            imageDecodePort: decode,
            document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
          });
          let bytes: Uint8Array;
          try {
            await apply(seed, seedTarget);
            bytes = new Uint8Array(await seed.save());
          } finally {
            seed.destroy();
          }
          const fixture = {
            document: bytes,
            clientIds: clients,
            editor: { imageDecodePort: decode },
          };
          const pair = await createTwoEditorFixture(fixture);
          try {
            const actor = pair.editors[writer]!,
              peer = pair.editors[1 - writer]!;
            const original = snapshot(actor),
              paragraphId = actor.surface!.session.paragraphIds()[0]!;
            const controls = descendants(original.part.root).filter(
              (node) => node.kind === 'contentControl'
            );
            const first = controls[0]!,
              second = conflict ? first : controls[1]!;
            const firstBlock = [...buildBlockIndex(original.pkg).byId.values()].find(
              (block) => block.controlId === first.id
            )!;
            const secondBlock = [...buildBlockIndex(original.pkg).byId.values()].find(
              (block) => block.controlId === second.id
            )!;
            const drawA = descendants(first).find((node) => node.kind === 'drawing')!,
              drawB = descendants(second).find((node) => node.kind === 'drawing')!;
            const targetA = `Lead. ${alpha} bridge ${image} Tail.`,
              targetB = conflict
                ? `Lead. ${beta} bridge ${image} Tail.`
                : `Lead. ${image} bridge ${beta} Tail.`;
            pair.pause();
            expect((await apply(actor, targetA)).resolved).toEqual(['asset:alpha']);
            const a = snapshot(actor);
            expect((await apply(peer, targetB)).resolved).toEqual(['asset:beta']);
            const b = snapshot(peer);
            const relation = (part: OoxmlPart, id: string) =>
              descendants(findNode(part, id)!).flatMap((node) =>
                node.kind === 'textValue'
                  ? []
                  : node.attributes
                      .filter((attr) => attr.localName === 'embed')
                      .map((attr) => attr.value)
              )[0]!;
            const change = (
              part: OoxmlPart,
              controlId: string,
              drawingNodeId: string,
              relationshipId: string,
              name: string
            ) => {
              let result = op(part, {
                op: 'setContentControlProperties',
                controlId,
                lock: 'unlocked',
              });
              result = op(result, { op: 'replaceDrawingResource', drawingNodeId, relationshipId });
              result = op(result, {
                op: 'setDrawingMetadata',
                drawingNodeId,
                title: name,
                description: name,
              });
              return op(result, {
                op: 'setContentControlProperties',
                controlId,
                lock: 'sdtContentLocked',
              });
            };
            const wordA = change(
              original.part,
              first.id,
              drawA.id,
              relation(a.part, drawA.id),
              'Alpha'
            );
            const wordB = change(
              original.part,
              second.id,
              drawB.id,
              relation(b.part, drawB.id),
              'Beta'
            );
            expect(canonicalOoxmlFingerprint(a.part)).toBe(canonicalOoxmlFingerprint(wordA));
            expect(canonicalOoxmlFingerprint(b.part)).toBe(canonicalOoxmlFingerprint(wordB));
            expect(
              peer.surface!.applyAutomationOps(
                () => [{ op: 'insertText', paragraphId, offset: 1, text: ' peer' }],
                { kind: 'body' },
                [],
                false,
                true
              ).rejected
            ).toBe(false);
            pair.reconnect(receiver);
            await pair.settle();
            const check = async (activeA: boolean, activeB: boolean, typed: boolean) => {
              for (const editor of pair.editors) {
                const current = snapshot(editor),
                  index = buildBlockIndex(current.pkg);
                let expected = original.part;
                let selected: 'a' | 'b' | null = null;
                if (conflict && activeA && activeB) {
                  const rel = relation(current.part, drawA.id);
                  expect([relation(a.part, drawA.id), relation(b.part, drawB.id)]).toContain(rel);
                  selected = rel === relation(a.part, drawA.id) ? 'a' : 'b';
                  expected = selected === 'a' ? wordA : wordB;
                } else {
                  if (activeA) expected = wordA;
                  if (activeB)
                    expected = change(
                      expected,
                      second.id,
                      drawB.id,
                      relation(b.part, drawB.id),
                      'Beta'
                    );
                  selected = activeB ? 'b' : activeA ? 'a' : null;
                }
                if (typed)
                  expected = op(expected, {
                    op: 'insertText',
                    paragraphId,
                    offset: 1,
                    text: ' peer',
                  });
                expect(canonicalOoxmlFingerprint(current.part)).toBe(
                  canonicalOoxmlFingerprint(expected)
                );
                const payloads = new Map(original.payloads);
                if (conflict) {
                  if (selected)
                    payloads.set(
                      firstBlock.id,
                      (selected === 'a' ? a : b).payloads.get(firstBlock.id)!
                    );
                } else {
                  if (activeA) payloads.set(firstBlock.id, a.payloads.get(firstBlock.id)!);
                  if (activeB) payloads.set(secondBlock.id, b.payloads.get(secondBlock.id)!);
                }
                expect([...current.payloads].sort(([x], [y]) => x.localeCompare(y))).toEqual(
                  [...payloads].sort(([x], [y]) => x.localeCompare(y))
                );
                expect(current.media).toEqual(
                  overlayMedia(original.media, [...(activeA ? [a] : []), ...(activeB ? [b] : [])])
                );
                for (const block of index.byId.values()) {
                  const rel = relation(
                    current.part,
                    descendants(findNode(current.part, block.controlId)!).find(
                      (node) => node.kind === 'drawing'
                    )!.id
                  );
                  const resource = resolveImageRelationship(
                    current.pkg.relationships.get(current.pkg.mainDocumentPart),
                    current.pkg.mainDocumentPart,
                    rel
                  );
                  expect(resource.mode).toBe('internal');
                  if (resource.mode !== 'internal') throw Error('Image lost its native media');
                  expect(resource.partName).toBe(block.payload.attributes.mediaPartName);
                  const expectedBytes =
                    block.payload.attributes.reference === 'asset:cat'
                      ? png
                      : block.payload.attributes.reference === 'asset:alpha'
                        ? replacement
                        : betaPng;
                  expect(
                    current.pkg.partBytes.get(resource.partName) ??
                      current.pkg.partBytes.get(resource.partName.slice(1))
                  ).toEqual(expectedBytes);
                }
              }
            };
            await check(true, true, true);
            expect(actor.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(false, true, true);
            expect(actor.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(true, true, true);
            expect(peer.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(true, true, false);
            expect(peer.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(true, false, false);
            expect(actor.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(false, false, false);
            expect(peer.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(false, true, false);
            expect(actor.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(true, true, false);
            expect(peer.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(true, true, true);
            const selectionOffset = 1;
            peer.surface!.setSelection({
              anchor: { paragraphId, offset: selectionOffset },
              head: { paragraphId, offset: selectionOffset },
            });
            await pair.settle();
            const geometry = pair.editors.map((editor) =>
              caretAt(editor.surface!.layout(), { paragraphId, offset: selectionOffset })
            );
            expect(geometry[0]).not.toBeNull();
            expect(geometry[1]).not.toBeNull();
            expect({ x: geometry[0]!.x, y: geometry[0]!.y }).toEqual({
              x: geometry[1]!.x,
              y: geometry[1]!.y,
            });
            const final = snapshot(actor),
              updates = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
            for (const shared of [false, true]) {
              const reopened = await createTwoEditorFixture({
                ...fixture,
                clientIds: undefined,
                document: new Uint8Array(await actor.save()),
                ...(shared ? { roomSnapshot: updates[writer] } : {}),
              });
              try {
                await reopened.settle();
                for (const editor of reopened.editors) {
                  const current = snapshot(editor);
                  expect(canonicalOoxmlFingerprint(current.part)).toBe(
                    canonicalOoxmlFingerprint(final.part)
                  );
                  expect(current.media).toEqual(final.media);
                  expect([...current.payloads].sort(([x], [y]) => x.localeCompare(y))).toEqual(
                    [...final.payloads].sort(([x], [y]) => x.localeCompare(y))
                  );
                }
              } finally {
                reopened.destroy();
              }
            }
          } finally {
            pair.destroy();
          }
        });
