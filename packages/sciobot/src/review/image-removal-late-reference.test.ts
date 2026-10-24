import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  allocateDrawingPropertyId,
  canonicalOoxmlFingerprint,
  contentControlsIn,
  contentControlPropertiesOf,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  imageRemovalFixture,
  imageRemovalContext as context,
  imageRemovalDecodePort,
  imageRemovalMedia as media,
  imageRemovalNodes,
} from '../../test/support/image-removal';
import { expectImageRelationships } from '../../test/support/image-relationships';
import { buildBlockIndex } from '../blocks/block-index';
import { blockTag, withPayloadRecord } from '../blocks/payload-opc-codec';
import { materializeImageReservation } from '../native/inline-image-reservation';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from './working-view';
import { planTargetDocument } from '../diff/plan-target';
import { previewDirectPlan } from '../ai/apply-direct';
import { recomposeOwnedProposals } from './recompose';
import { acceptAllReviewGroups } from './accept-all';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
const word = (pkg: OoxmlPackage) => {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!);
};
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  test(`late offline drawing reference preserves retired image resources ${clients}`, async () => {
    const source = await imageRemovalFixture(false, false);
    const pair = await createTwoEditorFixture({
      document: writeOoxmlPackage(source),
      documentId: context.documentId,
      clientIds: clients,
      ownedAttributes: true,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE], imageDecodePort: imageRemovalDecodePort },
    });
    try {
      const [alice, bob] = pair.editors;
      const token = (editor: (typeof pair.editors)[number]) =>
        readEditorRevision(editor, context.documentId, context.epoch)!;
      const before = bob!.surface!.session.currentPackage();
      const cat = [...buildBlockIndex(before).byId.values()].find(
        (block) => block.payload.attributes.reference === 'asset:cat'
      )!;
      const control = findNode(before.parts.get(cat.partName)!, cat.controlId)!;
      const drawing = imageRemovalNodes(control).find((node) => node.kind === 'drawing')!;
      if (drawing.kind !== 'drawing') throw Error('Missing source drawing.');
      const paragraphId = bob!.surface!.session.paragraphIds()[0]!;
      const current = await readWorkingView(alice!.surface!.session, token(alice!));
      if (!current.success) throw Error(current.error);
      const target = current.view.projection.text.replace('![Cat](asset:cat "Photo")', '');
      expect(target).not.toBe(current.view.projection.text);
      pair.pause();
      const staged = await recomposeOwnedProposals(
        alice!,
        { revision: token(alice!), targetMarkdown: target, ownership: [] },
        context,
        { payloads: { presence: { proposalId: 'late-remove-image' } } }
      );
      if (!staged.success) throw Error(staged.error);
      // Bob duplicates the observed drawing and its existing media reference without new bytes.
      const added = bob!.surface!.applyAutomationOps(
        () => [
          {
            op: 'insertInlineContentControl',
            paragraphId,
            offset: paragraphTextOf(bob!.surface!.session.part(), paragraphId)!.length,
            tag: blockTag('peer-copy'),
            alias: 'Image',
            text: '\uFFFC',
          },
        ],
        { kind: 'body' },
        [
          (pkg) => {
            const part = pkg.parts.get(cat.partName)!;
            const created = contentControlsIn(part.root).find(
              (entry) => contentControlPropertiesOf(entry.node).tag === blockTag('peer-copy')
            )!.node;
            const property = allocateDrawingPropertyId(pkg);
            if (!property.ok) throw Error('Cannot allocate peer drawing identity.');
            const remeta = (node: OoxmlNode): OoxmlNode =>
              node.kind === 'textValue'
                ? node
                : ({
                    ...node,
                    ...(node.kind === 'drawingDocPr'
                      ? {
                          attributes: node.attributes.map((attr) =>
                            attr.namespaceUri === '' && attr.localName === 'id'
                              ? { ...attr, value: String(property.id) }
                              : attr
                          ),
                        }
                      : {}),
                    children: node.children.map(remeta),
                  } as OoxmlNode);
            const copied = materializeImageReservation(
              pkg,
              cat.partName,
              created.id,
              blockTag('peer-copy'),
              'peer-copy',
              remeta(drawing) as typeof drawing,
              []
            );
            const stored = withPayloadRecord(
              copied,
              {
                ...cat.payload,
                id: 'peer-copy',
                attributes: { ...cat.payload.attributes, reference: 'asset:peer' },
              },
              'Image'
            );
            if (!stored.ok) throw Error(stored.reason);
            return stored.pkg;
          },
        ],
        false,
        true,
        'after'
      );
      expect(added).toMatchObject({ committed: true, rejected: false });
      const peer = bob!.surface!.session.currentPackage();
      expect(media(peer)).toEqual(media(before));
      expectImageRelationships(peer, buildBlockIndex(peer).byId.values());
      const peerView = await readWorkingView(bob!.surface!.session, token(bob!));
      if (!peerView.success) throw Error(peerView.error);
      const removed = planTargetDocument(
        { pkg: peer, revision: token(bob!), projection: peerView.view.projection },
        peerView.view.projection.text.replace('![Cat](asset:cat "Photo")', '')
      );
      if (!removed.success) throw Error(removed.error);
      const expected = previewDirectPlan(peer, removed.plan);
      if (!expected.success) throw Error(expected.error);
      expect(media(expected.pkg)).toEqual(media(before));
      const decided = acceptAllReviewGroups(alice!, token(alice!), context);
      if (!decided.success) throw Error(decided.error);
      expect(media(alice!.surface!.session.currentPackage())).toHaveLength(1);
      pair.reconnect(clients[0] === 41 ? 0 : 1);
      await pair.settle();
      const check = () => {
        for (const editor of pair.editors) {
          const pkg = editor.surface!.session.currentPackage();
          expect(word(pkg)).toBe(word(expected.pkg));
          expect(media(pkg)).toEqual(media(expected.pkg));
          expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
        }
      };
      check();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      check();
      const saved = new Uint8Array(await bob!.save());
      const reopened = await createTwoEditorFixture({
        document: saved,
        documentId: context.documentId,
        ownedAttributes: true,
        editor: {
          modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
          imageDecodePort: imageRemovalDecodePort,
        },
        roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
      });
      try {
        for (const editor of reopened.editors) {
          const pkg = editor.surface!.session.currentPackage();
          expect(word(pkg)).toBe(word(expected.pkg));
          expect(media(pkg)).toEqual(media(expected.pkg));
          expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
        }
        reopened.ownership.forEach((handle) => handle.assertHealthy());
      } finally {
        reopened.destroy();
      }
      pair.ownership.forEach((handle) => handle.assertHealthy());
    } finally {
      pair.destroy();
    }
  }, 30000);
