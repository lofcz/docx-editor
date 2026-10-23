import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
// Test-only collaboration transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  resolveImageRelationship,
  type OoxmlNode,
  type OoxmlPackage,
  type ImageDecodePort,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { prepareNativeDocument } from '../native/prepare-document';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { parseTargetMarkdown } from '../projection/parse-target';
import { targetAstKey } from './target-key';
import { planTargetDocument } from './plan-target';
import { applyDirectPlan } from '../ai/apply-direct';
import { buildBlockIndex } from '../blocks/block-index';
const context = { documentId: 'shared-images', epoch: 'native' };
const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
// A valid ancillary PNG chunk gives the peer an independently identifiable byte stream.
const chunkData = new TextEncoder().encode('Comment\0Peer');
const chunk = new Uint8Array(12 + chunkData.length);
new DataView(chunk.buffer).setUint32(0, chunkData.length);
chunk.set([116, 69, 88, 116], 4);
chunk.set(chunkData, 8);
let crc = 0xffffffff;
for (const byte of chunk.subarray(4, -4)) {
  crc ^= byte;
  for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
}
new DataView(chunk.buffer).setUint32(chunk.length - 4, (crc ^ 0xffffffff) >>> 0);
const peerPng = new Uint8Array(png.length + chunk.length);
peerPng.set(png.subarray(0, 33));
peerPng.set(chunk, 33);
peerPng.set(png.subarray(33), 33 + chunk.length);
function pictureRelationship(node: OoxmlNode): string[] {
  if (node.kind === 'textValue') return [];
  const id =
    node.localName === 'blip'
      ? node.attributes.find(
          (attribute) =>
            attribute.localName === 'embed' &&
            attribute.namespaceUri ===
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        )?.value
      : undefined;
  return [...(id ? [id] : []), ...node.children.flatMap(pictureRelationship)];
}
const decode: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes]
    .filter(([name]) => /(^|\/)media\//.test(name))
    .sort(([a], [b]) => a.localeCompare(b));
const key = (source: string) => {
  const parsed = parseTargetMarkdown(source);
  if (!parsed.success) throw Error(parsed.error);
  return targetAstKey(parsed.target.normalizedAst);
};
for (const ownership of ['plain', 'linked', 'code-linked'] as const)
  for (const alternate of [false, true])
    for (const clients of [
      [41, 73],
      [73, 41],
    ] as const)
      for (const receiver of [0, 1] as const)
        test(`inline images retain peer media and text through both histories, geometry and DOCX/room reopen: ${clients}/${receiver}/alternate=${alternate}/ownership=${ownership}`, async () => {
          const initial = docxFixture(
            '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>Neighbor</w:t></w:r></w:p><w:p><w:r><w:t>Peer</w:t></w:r></w:p>'
          );
          const entries = unzipSync(initial);
          if (alternate) {
            entries['custom/main.xml'] = entries['word/document.xml']!;
            entries['custom/_rels/main.xml.rels'] = entries['word/_rels/document.xml.rels']!;
            delete entries['word/document.xml'];
            delete entries['word/_rels/document.xml.rels'];
            entries['_rels/.rels'] = strToU8(
              strFromU8(entries['_rels/.rels']!).replace('word/document.xml', 'custom/main.xml')
            );
            entries['[Content_Types].xml'] = strToU8(
              strFromU8(entries['[Content_Types].xml']!).replace(
                '/word/document.xml',
                '/custom/main.xml'
              )
            );
          }
          const prepared = prepareNativeDocument(alternate ? zipSync(entries) : initial);
          if (!prepared.success) throw Error(prepared.error);
          const fixture = {
            document: prepared.document,
            clientIds: clients,
            editor: { imageDecodePort: decode },
          };
          const pair = await createTwoEditorFixture(fixture);
          try {
            const [alice, bob] = pair.editors;
            const neighbor = alice!.surface!.session.paragraphIds()[1]!;
            const imageA =
              ownership === 'plain'
                ? 'Neighbor A ![Cat](asset:cat "Photo") Z'
                : ownership === 'linked'
                  ? 'Neighbor A [![Cat](asset:cat "Photo")](https://outer.test "Read") Z'
                  : 'Neighbor A [<strong><code>![Cat](asset:cat "Photo")</code></strong>](https://outer.test "Read") Z';
            const imageB = 'Peer B ![Dog](asset:dog) Y';
            const apply = async (
              editor: (typeof pair.editors)[number],
              transform: (text: string) => string,
              expected: string
            ) => {
              const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
              const source = await readWorkingView(editor.surface!.session, revision);
              if (!source.success) throw Error(source.error);
              const planned = planTargetDocument(
                {
                  revision,
                  pkg: editor.surface!.session.currentPackage(),
                  projection: source.view.projection,
                },
                transform(source.view.projection.text)
              );
              if (!planned.success) throw Error(planned.error);
              const resolved: string[] = [];
              const resources = await prepareInlineImageResources(
                planned.plan,
                async (reference) => {
                  resolved.push(reference);
                  return {
                    bytes: reference === 'asset:dog' ? peerPng : png,
                    mime: 'image/png',
                    widthPoints: 12,
                    heightPoints: 8,
                  };
                },
                decode
              );
              if (!resources.success) throw Error(resources.error);
              expect(resolved).toEqual([expected]);
              const result = applyDirectPlan(editor, planned.plan, {
                ...context,
                images: resources.images,
              });
              if (!result.success) throw Error(result.error);
            };
            const peerImage = { word: undefined as string | undefined };
            const check = async (a = true, b = true, prefix = ' Peer', tail = '') => {
              const updates = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
              for (const editor of pair.editors) {
                const pkg = editor.surface!.session.currentPackage(),
                  index = buildBlockIndex(pkg);
                expect(index.diagnostics).toEqual([]);
                const byReference = new Map(
                  [...index.byId.values()].map((block) => [
                    block.payload.attributes.reference,
                    block,
                  ])
                );
                expect(byReference.has('asset:cat')).toBe(a);
                expect(byReference.has('asset:dog')).toBe(b);
                expect(media(pkg)).toHaveLength(Number(a) + Number(b));
                for (const block of byReference.values()) {
                  const ids = pictureRelationship(
                    findNode(editor.surface!.session.part(), block.controlId)!
                  );
                  expect(ids).toHaveLength(1);
                  const resource = resolveImageRelationship(
                    pkg.relationships.get(pkg.mainDocumentPart),
                    pkg.mainDocumentPart,
                    ids[0]!
                  );
                  expect(resource.mode).toBe('internal');
                  if (resource.mode !== 'internal')
                    throw Error('The image has no native media relationship.');
                  const mediaPartName = block.payload.attributes.mediaPartName;
                  if (typeof mediaPartName !== 'string')
                    throw Error('The image media part is invalid.');
                  expect(resource.partName).toBe(mediaPartName);
                  expect(
                    pkg.partBytes.get(resource.partName) ??
                      pkg.partBytes.get(resource.partName.replace(/^\//, ''))
                  ).toEqual(block.payload.attributes.reference === 'asset:dog' ? peerPng : png);
                }
                expect(paragraphTextOf(editor.surface!.session.part(), neighbor)).toBe(
                  'Neigh' + prefix + 'bor' + (a ? ' A \uFFFC Z' : '') + tail
                );
                if (b && peerImage.word !== undefined)
                  expect(
                    canonicalOoxmlFingerprint(
                      findNode(
                        editor.surface!.session.part(),
                        byReference.get('asset:dog')!.controlId
                      )!
                    )
                  ).toBe(peerImage.word);
                const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
                const source = await readWorkingView(editor.surface!.session, revision);
                if (!source.success) throw Error(source.error);
                const target =
                  'Alpha\n\n' +
                  (a ? imageA : 'Neighbor').replace('Neighbor', 'Neigh' + prefix + 'bor') +
                  tail +
                  '\n\n' +
                  (b ? imageB : 'Peer');
                expect(key(source.view.projection.text)).toBe(key(target));
                expect(editor.surface!.session.currentPackage()).toBe(pkg);
              }
              pair.documents.forEach((doc, index) =>
                expect(Y.encodeStateAsUpdate(doc)).toEqual(updates[index]!)
              );
              expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(
                canonicalOoxmlFingerprint(bob!.surface!.session.part())
              );
              expect(media(alice!.surface!.session.currentPackage())).toEqual(
                media(bob!.surface!.session.currentPackage())
              );
              const [aDigest, bDigest] = await pair.savedDigests();
              expect(aDigest).toEqual(bDigest);
            };
            pair.pause();
            await apply(alice!, (source) => source.replace('Neighbor', imageA), 'asset:cat');
            expect(
              bob!.surface!.applyAutomationOps(
                () => [{ op: 'insertText', paragraphId: neighbor, offset: 5, text: ' Peer' }],
                { kind: 'body' },
                [],
                false,
                true
              ).rejected
            ).toBe(false);
            pair.reconnect(receiver);
            await pair.settle();
            await check(true, false);
            await apply(bob!, (source) => source.replace(/Peer\s*$/, imageB), 'asset:dog');
            await pair.settle();
            await check();
            const peerBlock = [
              ...buildBlockIndex(alice!.surface!.session.currentPackage()).byId.values(),
            ].find((block) => block.payload.attributes.reference === 'asset:dog')!;
            peerImage.word = canonicalOoxmlFingerprint(
              findNode(alice!.surface!.session.part(), peerBlock.controlId)!
            );
            pair.pause();
            expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
            expect(
              bob!.surface!.applyAutomationOps(
                () => [
                  {
                    op: 'insertText',
                    paragraphId: neighbor,
                    offset: paragraphTextOf(bob!.surface!.session.part(), neighbor)!.length,
                    text: ' More',
                  },
                ],
                { kind: 'body' },
                [],
                false,
                true
              ).rejected
            ).toBe(false);
            pair.reconnect(receiver);
            await pair.settle();
            await check(false, true, ' Peer', ' More');
            expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(true, true, ' Peer', ' More');
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check();
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(true, false);
            expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(true, false, '');
            expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
            expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
            expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(true, true, ' Peer', ' More');
            const offset = paragraphTextOf(bob!.surface!.session.part(), neighbor)!.length;
            bob!.surface!.setSelection({
              anchor: { paragraphId: neighbor, offset },
              head: { paragraphId: neighbor, offset },
            });
            await pair.settle();
            const remote = pair.rooms[0]!.session.remoteSelections().find(
              (selection) => selection.actorId === 'bob'
            );
            expect(remote?.head.nodeId).toBe(neighbor);
            expect(remote?.head.offset).toBe(offset);
            const aCaret = caretAt(alice!.surface!.layout(), { paragraphId: neighbor, offset }),
              bCaret = caretAt(bob!.surface!.layout(), { paragraphId: neighbor, offset });
            expect(aCaret).not.toBeNull();
            expect(bCaret).not.toBeNull();
            expect({ x: aCaret!.x, y: aCaret!.y }).toEqual({ x: bCaret!.x, y: bCaret!.y });
            const expectedWord = canonicalOoxmlFingerprint(alice!.surface!.session.part());
            const expectedMedia = media(alice!.surface!.session.currentPackage());
            const expectedPayloads = [
              ...buildBlockIndex(alice!.surface!.session.currentPackage()).byId,
            ].map(([id, block]) => [id, block.payload]);
            for (const shared of [false, true]) {
              const reopened = await createTwoEditorFixture({
                ...fixture,
                clientIds: undefined,
                document: new Uint8Array(await alice!.save()),
                ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
              });
              try {
                await reopened.settle();
                for (const editor of reopened.editors) {
                  const pkg = editor.surface!.session.currentPackage();
                  expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
                    expectedWord
                  );
                  expect(media(pkg)).toEqual(expectedMedia);
                  expect(
                    [...buildBlockIndex(pkg).byId].map(([id, block]) => [id, block.payload])
                  ).toEqual(expectedPayloads);
                  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
                  const source = await readWorkingView(editor.surface!.session, revision);
                  if (!source.success) throw Error(source.error);
                  expect(key(source.view.projection.text)).toBe(
                    key(
                      'Alpha\n\n' +
                        imageA.replace('Neighbor', 'Neigh Peerbor') +
                        ' More\n\n' +
                        imageB
                    )
                  );
                }
                const [aDigest, bDigest] = await reopened.savedDigests();
                expect(aDigest).toEqual(bDigest);
              } finally {
                reopened.destroy();
              }
            }
          } finally {
            pair.destroy();
          }
        });
