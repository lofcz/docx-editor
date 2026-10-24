// Evaluation-only collaboration transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphOffsetIndex,
  paragraphTextOf,
  readOoxmlPackage,
  type ImageDecodePort,
  type OoxmlNode,
  type OoxmlPart,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { buildBlockIndex } from '../blocks/block-index';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from '../ai/apply-direct';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { planTargetDocument } from './plan-target';
import { decodeEditPlan, serializeEditPlan } from './plan';
const context = { documentId: 'image-birth-batch', epoch: 'native' };
const bytes = (value: string) => Uint8Array.from(atob(value), (c) => c.charCodeAt(0));
const cat = bytes(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
);
const dog = bytes('R0lGODlhAQABAPAAAP///wAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw==');
const decode: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const descendants = (node: OoxmlNode): OoxmlNode[] => [
  node,
  ...(node.kind === 'textValue' ? [] : node.children.flatMap(descendants)),
];
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes]
    .filter(([name]) => /(^|\/)media\//.test(name))
    .sort(([a], [b]) => a.localeCompare(b));
const kinds = {
  image: { markdown: '![Cat](asset:cat "Photo")', fallback: '\ufffc' },
  math: { markdown: '$y$', fallback: 'y' },
  mention: { markdown: '<scio-mention key="user" value="Ada" />', fallback: 'Ada' },
  source: { markdown: '<source title="Paper" url="https://source.test" />', fallback: 'Paper' },
} as const;
type Kind = keyof typeof kinds;
const originalImage = kinds.image.markdown;
const replacedImage = '![Dog](asset:dog "Changed")';
function targets(kind: Kind, before: boolean, ownership: 'plain' | 'linked' | 'code-linked') {
  const wrap = (value: string) =>
    ownership === 'plain'
      ? value
      : `[${ownership === 'code-linked' ? `<code>${value}</code>` : value}](https://outer.test "Read")`;
  return {
    seed: wrap(`Lead. anchor ${originalImage} bridge $x$ Tail.`),
    target: wrap(
      before
        ? `Lead. ${kinds[kind].markdown} Fresh anchor ${replacedImage} bridge $x$ Tail.`
        : `Lead. anchor ${replacedImage} bridge Fresh ${kinds[kind].markdown} $x$ Tail.`
    ),
  };
}
async function plan(editor: DocxEditorInstance, target: string) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw Error(view.error);
  const planned = planTargetDocument(
    { revision, pkg: editor.surface!.session.currentPackage(), projection: view.view.projection },
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
        bytes: reference === 'asset:dog' ? dog : cat,
        mime: reference === 'asset:dog' ? 'image/gif' : 'image/png',
        widthPoints: reference === 'asset:dog' ? 40 : 12,
        heightPoints: reference === 'asset:dog' ? 30 : 8,
      };
    },
    decode
  );
  if (!prepared.success) throw Error(prepared.error);
  const result = applyDirectPlan(editor, planned, { ...context, images: prepared.images });
  if (!result.success) throw Error(result.error);
  return { planned, resolved };
}
function native(part: OoxmlPart, op: TreeDocOp) {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw Error(result.reason);
  return result.part;
}
function snapshot(editor: DocxEditorInstance) {
  const pkg = editor.surface!.session.currentPackage(),
    index = buildBlockIndex(pkg);
  expect(index.diagnostics).toEqual([]);
  return {
    part: editor.surface!.session.part(),
    pkg,
    word: canonicalOoxmlFingerprint(editor.surface!.session.part()),
    media: media(pkg),
    payloads: [...index.byId]
      .map(([id, b]) => [id, b.payload])
      .sort(([a], [b]) => String(a).localeCompare(String(b))),
  };
}
function verify(
  before: ReturnType<typeof snapshot>,
  editor: DocxEditorInstance,
  kind: Kind,
  first: boolean
) {
  const current = snapshot(editor),
    old = buildBlockIndex(before.pkg),
    index = buildBlockIndex(current.pkg);
  expect(old.byId.size).toBe(2);
  expect(index.byId.size).toBe(3);
  const picture = [...old.byId.values()].find((b) => b.payload.type === 'inline_image')!;
  const math = [...old.byId.values()].find((b) => b.id !== picture.id)!;
  const fresh = [...index.byId.values()].filter((b) => !old.byId.has(b.id));
  expect(fresh).toHaveLength(1);
  const paragraph = descendants(before.part.root).find((n) => n.kind === 'paragraph')!;
  if (paragraph.kind !== 'paragraph') throw Error('Missing paragraph');
  const nextParagraph = findNode(current.part, paragraph.id)!;
  if (nextParagraph.kind !== 'paragraph') throw Error('Missing retained paragraph');
  const spans = paragraphOffsetIndex(paragraph),
    nextSpans = paragraphOffsetIndex(nextParagraph);
  const oldSpan = spans.spanOf(picture.controlId)!;
  const shift = kinds[kind].fallback.length + ' Fresh '.length;
  expect(nextSpans.spanOf(picture.controlId)).toEqual({
    start: oldSpan.start + (first ? shift : 0),
    end: oldSpan.end + (first ? shift : 0),
  });
  const mathSpan = spans.spanOf(math.controlId)!;
  expect(nextSpans.spanOf(math.controlId)).toEqual({
    start: mathSpan.start + shift,
    end: mathSpan.end + shift,
  });
  expect(nextSpans.spanOf(fresh[0]!.controlId)).toEqual({
    start: first ? 'Lead. '.length : oldSpan.end + ' bridge Fresh '.length,
    end:
      (first ? 'Lead. '.length : oldSpan.end + ' bridge Fresh '.length) +
      kinds[kind].fallback.length,
  });
  expect(paragraphTextOf(current.part, paragraph.id)).toBe(
    first
      ? `Lead. ${kinds[kind].fallback} Fresh anchor \ufffc bridge x Tail.`
      : `Lead. anchor \ufffc bridge Fresh ${kinds[kind].fallback} x Tail.`
  );
  expect(index.byId.get(math.id)).toEqual(math);
  expect(canonicalOoxmlFingerprint(findNode(current.part, math.controlId)!)).toBe(
    canonicalOoxmlFingerprint(findNode(before.part, math.controlId)!)
  );
  expect(index.byId.get(picture.id)?.controlId).toBe(picture.controlId);
  const drawing = descendants(findNode(before.part, picture.controlId)!).find(
    (n) => n.kind === 'drawing'
  )!;
  const updatedDrawing = findNode(current.part, drawing.id)!;
  expect(updatedDrawing.kind).toBe('drawing');
  const relationshipId = descendants(updatedDrawing).flatMap((n) =>
    n.kind === 'textValue'
      ? []
      : n.attributes.filter((a) => a.localName === 'embed').map((a) => a.value)
  )[0]!;
  let expected = native(before.part, {
    op: 'setContentControlProperties',
    controlId: picture.controlId,
    lock: 'unlocked',
  });
  expected = native(expected, {
    op: 'replaceDrawingResource',
    drawingNodeId: drawing.id,
    relationshipId,
  });
  expected = native(expected, {
    op: 'setDrawingMetadata',
    drawingNodeId: drawing.id,
    title: 'Changed',
    description: 'Dog',
  });
  expected = native(expected, {
    op: 'setContentControlProperties',
    controlId: picture.controlId,
    lock: 'sdtContentLocked',
  });
  expect(canonicalOoxmlFingerprint(findNode(current.part, picture.controlId)!)).toBe(
    canonicalOoxmlFingerprint(findNode(expected, picture.controlId)!)
  );
  const changed = index.byId.get(picture.id)!.payload;
  expect(changed).toEqual({
    ...picture.payload,
    attributes: {
      ...picture.payload.attributes,
      reference: 'asset:dog',
      alt: 'Dog',
      title: 'Changed',
      contentHash: changed.attributes.contentHash,
      mediaPartName: changed.attributes.mediaPartName,
    },
  });
  const name = changed.attributes.mediaPartName;
  if (typeof name !== 'string') throw Error('Missing image media');
  expect(current.pkg.partBytes.get(name) ?? current.pkg.partBytes.get(name.slice(1))).toEqual(dog);
  expect(current.media).toHaveLength(kind === 'image' ? 2 : 1);
  return current;
}
for (const kind of Object.keys(kinds) as Kind[])
  for (const first of [false, true])
    for (const ownership of ['plain', 'linked', 'code-linked'] as const)
      test(`image replacement and fresh atoms preserve native ownership: ${kind}/${first}/${ownership}`, async () => {
        const editor = createDocxEditor({
          container: document.createElement('div'),
          imageDecodePort: decode,
          document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
        });
        const target = targets(kind, first, ownership);
        try {
          await apply(editor, target.seed);
          const before = snapshot(editor);
          const { planned, resolved } = await apply(editor, target.target);
          expect(planned.operations.filter((op) => op.kind === 'inline-image')).toHaveLength(1);
          expect(
            planned.operations.flatMap((op) => (op.kind === 'inline-atoms' ? op.atoms : []))
          ).toHaveLength(1);
          expect(resolved.sort()).toEqual(
            kind === 'image' ? ['asset:cat', 'asset:dog'] : ['asset:dog']
          );
          const after = verify(before, editor, kind, first);
          expect((await plan(editor, target.target)).operations).toEqual([]);
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          const undone = snapshot(editor);
          expect({ word: undone.word, media: undone.media, payloads: undone.payloads }).toEqual({
            word: before.word,
            media: before.media,
            payloads: before.payloads,
          });
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          const redone = snapshot(editor);
          expect({ word: redone.word, media: redone.media, payloads: redone.payloads }).toEqual({
            word: after.word,
            media: after.media,
            payloads: after.payloads,
          });
          const saved = new Uint8Array(await editor.save()),
            loaded = readOoxmlPackage(saved);
          if (!loaded.ok) throw Error(loaded.reason);
          expect(
            canonicalOoxmlFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart)!)
          ).toBe(after.word);
          expect(media(loaded.package)).toEqual(after.media);
          const reopened = createDocxEditor({
            container: document.createElement('div'),
            document: saved,
            imageDecodePort: decode,
          });
          try {
            expect((await plan(reopened, target.target)).operations).toEqual([]);
            expect(snapshot(reopened).payloads).toEqual(after.payloads);
          } finally {
            reopened.destroy();
          }
        } finally {
          editor.destroy();
        }
      });
for (const kind of Object.keys(kinds) as Kind[])
  for (const first of [false, true])
    for (const clients of [
      [41, 73],
      [73, 41],
    ] as const)
      test(`replacement plus birth preserves peer text and both histories: ${kind}/${first}/${clients}`, async () => {
        const target = targets(kind, first, 'code-linked');
        const seed = createDocxEditor({
          container: document.createElement('div'),
          imageDecodePort: decode,
          document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
        });
        let documentBytes: Uint8Array;
        try {
          await apply(seed, target.seed);
          documentBytes = new Uint8Array(await seed.save());
        } finally {
          seed.destroy();
        }
        const fixture = {
          document: documentBytes,
          ownedAttributes: true,
          clientIds: clients,
          editor: { imageDecodePort: decode },
        };
        const pair = await createTwoEditorFixture(fixture);
        const writer = first ? 1 : 0,
          actor = pair.editors[writer]!,
          peer = pair.editors[1 - writer]!;
        try {
          const before = snapshot(actor),
            paragraphId = actor.surface!.session.paragraphIds()[0]!;
          pair.pause();
          await apply(actor, target.target);
          const after = verify(before, actor, kind, first);
          expect(
            peer.surface!.applyAutomationOps(
              () => [{ op: 'insertText', paragraphId, offset: 1, text: ' peer' }],
              { kind: 'body' },
              [],
              false,
              true
            ).rejected
          ).toBe(false);
          pair.reconnect(writer);
          await pair.settle();
          const check = (active: boolean, typed: boolean) => {
            const state = active ? after : before;
            const word = typed
              ? canonicalOoxmlFingerprint(
                  native(state.part, { op: 'insertText', paragraphId, offset: 1, text: ' peer' })
                )
              : state.word;
            for (const editor of pair.editors) {
              const current = snapshot(editor);
              expect(current.word).toBe(word);
              expect(current.payloads).toEqual(state.payloads);
              expect(current.media).toEqual(state.media);
            }
          };
          check(true, true);
          expect(actor.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(false, true);
          expect(actor.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(true, true);
          expect(peer.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(true, false);
          expect(actor.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(false, false);
          expect(peer.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(false, true);
          expect(actor.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(true, true);
          peer.surface!.setSelection({
            anchor: { paragraphId, offset: 1 },
            head: { paragraphId, offset: 1 },
          });
          await pair.settle();
          const geometry = pair.editors.map((e) =>
            caretAt(e.surface!.layout(), { paragraphId, offset: 1 })
          );
          expect(geometry[0]).not.toBeNull();
          expect(geometry[1]).not.toBeNull();
          expect({ x: geometry[0]!.x, y: geometry[0]!.y }).toEqual({
            x: geometry[1]!.x,
            y: geometry[1]!.y,
          });
          const updates = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
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
                expect(current.word).toBe(snapshot(actor).word);
                expect(current.media).toEqual(after.media);
                expect(current.payloads).toEqual(after.payloads);
              }
            } finally {
              reopened.destroy();
            }
          }
        } finally {
          pair.destroy();
        }
      });
