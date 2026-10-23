import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  type OoxmlNode,
  type ImageDecodePort,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { prepareNativeDocument } from '../native/prepare-document';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { planTargetDocument } from './plan-target';
import { serializeEditPlan, decodeEditPlan } from './plan';
import { applyDirectPlan } from '../ai/apply-direct';
import { buildBlockIndex } from '../blocks/block-index';
const context = { documentId: 'imported-boundary', epoch: 'native' };
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
function nodes(node: OoxmlNode): OoxmlNode[] {
  return [node, ...(node.kind === 'textValue' ? [] : node.children.flatMap(nodes))];
}
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`imported image boundary keeps foreign text outside the exact native link through both histories ${clients}/${receiver}`, async () => {
      const initial = docxFixture(
        '<w:p><w:r><w:t>Keep. </w:t></w:r><w:hyperlink r:id="existing" w:tooltip="Read" w:history="1" w:tgtFrame="_blank"><w:r><w:rPr xmlns:proof="urn:sciobot:proof"><w:rStyle w:val="ImportedBase"/><w:color w:val="112233"/><proof:metadata proof:keep="yes"/></w:rPr><w:t>label</w:t></w:r></w:hyperlink><w:r><w:t> End.</w:t></w:r></w:p>',
        {
          relationships:
            '<Relationship Id="existing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://outer.test" TargetMode="External"/>',
          styles:
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="character" w:styleId="ImportedBase"><w:name w:val="Imported base"/><w:rPr><w:color w:val="112233"/></w:rPr></w:style></w:styles>',
        }
      );
      const prepared = prepareNativeDocument(initial);
      if (!prepared.success) throw Error(prepared.error);
      const fixture = {
        document: prepared.document,
        clientIds: clients,
        editor: { imageDecodePort: decode },
      };
      const pair = await createTwoEditorFixture(fixture);
      try {
        const [alice, bob] = pair.editors;
        const part = alice!.surface!.session.part(),
          paragraphId = alice!.surface!.session.paragraphIds()[0]!;
        const link = nodes(part.root).find((n) => n.kind === 'hyperlink')!;
        if (link.kind !== 'hyperlink') throw Error('Missing native link');
        const outside = nodes(part.root).find(
          (n) => n.kind === 'textValue' && n.value === ' End.'
        )!;
        const target =
          'Keep. [label <code>![Cat](asset:cat "Photo")</code>](https://outer.test "Read") End.';
        const revision = readEditorRevision(alice!, context.documentId, context.epoch)!;
        const source = await readWorkingView(alice!.surface!.session, revision);
        if (!source.success) throw Error(source.error);
        const planned = planTargetDocument(
          {
            revision,
            pkg: alice!.surface!.session.currentPackage(),
            projection: source.view.projection,
          },
          target
        );
        if (!planned.success) throw Error(planned.error);
        expect(decodeEditPlan(serializeEditPlan(planned.plan))).toEqual({
          success: true,
          plan: planned.plan,
        });
        const resources = await prepareInlineImageResources(
          planned.plan,
          async () => ({ bytes: png, mime: 'image/png', widthPoints: 12, heightPoints: 8 }),
          decode
        );
        if (!resources.success) throw Error(resources.error);
        const check = async (image: boolean, peer: boolean) => {
          const updates = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
          for (const editor of pair.editors) {
            const current = editor.surface!.session.part(),
              pkg = editor.surface!.session.currentPackage();
            const kept = findNode(current, link.id);
            expect(kept?.kind).toBe('hyperlink');
            if (kept?.kind !== 'hyperlink') throw Error('Imported link disappeared');
            expect(kept.attributes).toEqual(link.attributes);
            expect(
              nodes(kept)
                .filter((n) => n.kind === 'textValue')
                .map((n) => (n.kind === 'textValue' ? n.value : ''))
                .join('')
            ).toBe(image ? 'label ' : 'label');
            expect(findNode(current, outside.id)).toMatchObject({
              kind: 'textValue',
              value: (peer ? ' peer' : '') + ' End.',
            });
            expect(paragraphTextOf(current, paragraphId)).toBe(
              'Keep. label' + (image ? ' \uFFFC' : '') + (peer ? ' peer' : '') + ' End.'
            );
            const index = buildBlockIndex(pkg);
            expect(index.diagnostics).toEqual([]);
            expect(index.byId.size).toBe(Number(image));
            if (image) {
              const block = [...index.byId.values()][0]!;
              expect(nodes(kept).some((n) => n.id === block.controlId)).toBe(true);
              const properties = nodes(findNode(current, block.controlId)!);
              for (const [name, value] of [
                ['rStyle', 'SciobotCode__ImportedBase'],
                ['color', '112233'],
              ] as const)
                expect(
                  properties.some(
                    (n) =>
                      n.kind !== 'textValue' &&
                      n.localName === name &&
                      n.attributes.some((a) => a.localName === 'val' && a.value === value)
                  )
                ).toBe(true);
              expect(
                properties.some(
                  (n) =>
                    n.kind !== 'textValue' &&
                    n.namespaceUri === 'urn:sciobot:proof' &&
                    n.localName === 'metadata'
                )
              ).toBe(true);
            }
            expect([...pkg.partBytes].filter(([name]) => name.includes('/media/'))).toHaveLength(
              Number(image)
            );
          }
          pair.documents.forEach((doc, i) =>
            expect(Y.encodeStateAsUpdate(doc)).toEqual(updates[i]!)
          );
          expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(
            canonicalOoxmlFingerprint(bob!.surface!.session.part())
          );
        };
        pair.pause();
        const result = applyDirectPlan(alice!, planned.plan, {
          ...context,
          images: resources.images,
        });
        if (!result.success) throw Error(result.error);
        expect(
          bob!.surface!.applyAutomationOps(
            () => [{ op: 'insertText', paragraphId, offset: 11, text: ' peer' }],
            { kind: 'body' },
            [],
            false,
            true
          ).rejected
        ).toBe(false);
        pair.reconnect(receiver);
        await pair.settle();
        await check(true, true);
        const combined = canonicalOoxmlFingerprint(alice!.surface!.session.part());
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(false, true);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(true, true);
        expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(combined);
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check(true, false);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check(true, true);
        expect(canonicalOoxmlFingerprint(alice!.surface!.session.part())).toBe(combined);
        const expectedPayloads = [
          ...buildBlockIndex(alice!.surface!.session.currentPackage()).byId,
        ].map(([id, b]) => [id, b.payload]);
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
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(combined);
              expect(
                [...buildBlockIndex(editor.surface!.session.currentPackage()).byId].map(
                  ([id, b]) => [id, b.payload]
                )
              ).toEqual(expectedPayloads);
            }
          } finally {
            reopened.destroy();
          }
        }
      } finally {
        pair.destroy();
      }
    });
