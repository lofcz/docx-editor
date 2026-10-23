// Test-only collaboration transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
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
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`inline images retain peer media and text through both histories, geometry and DOCX/room reopen: ${clients}/${receiver}`, async () => {
      const prepared = prepareNativeDocument(
        docxFixture(
          '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>Neighbor</w:t></w:r></w:p><w:p><w:r><w:t>Peer</w:t></w:r></w:p>'
        )
      );
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
        const imageA = 'Neighbor A ![Cat](asset:cat "Photo") Z';
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
              return { bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 };
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
        let peerImageWord: string | undefined;
        const check = async (a = true, b = true, prefix = ' Peer', tail = '') => {
          const updates = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
          for (const editor of pair.editors) {
            const pkg = editor.surface!.session.currentPackage(),
              index = buildBlockIndex(pkg);
            expect(index.diagnostics).toEqual([]);
            const byReference = new Map(
              [...index.byId.values()].map((block) => [block.payload.attributes.reference, block])
            );
            expect(byReference.has('asset:cat')).toBe(a);
            expect(byReference.has('asset:dog')).toBe(b);
            expect(media(pkg)).toHaveLength(Number(a) + Number(b));
            for (const [, bytes] of media(pkg)) expect(bytes).toEqual(png);
            expect(paragraphTextOf(editor.surface!.session.part(), neighbor)).toBe(
              'Neigh' + prefix + 'bor' + (a ? ' A \uFFFC Z' : '') + tail
            );
            if (b && peerImageWord !== undefined)
              expect(
                canonicalOoxmlFingerprint(
                  findNode(editor.surface!.session.part(), byReference.get('asset:dog')!.controlId)!
                )
              ).toBe(peerImageWord);
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
        peerImageWord = canonicalOoxmlFingerprint(
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
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(expectedWord);
              expect(media(pkg)).toEqual(expectedMedia);
              expect(
                [...buildBlockIndex(pkg).byId].map(([id, block]) => [id, block.payload])
              ).toEqual(expectedPayloads);
              const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
              const source = await readWorkingView(editor.surface!.session, revision);
              if (!source.success) throw Error(source.error);
              expect(key(source.view.projection.text)).toBe(
                key(
                  'Alpha\n\n' + imageA.replace('Neighbor', 'Neigh Peerbor') + ' More\n\n' + imageB
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
