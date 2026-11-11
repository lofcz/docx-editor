/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  contentControlPropertiesOf,
  contentControlsIn,
  findNode,
  replaceNode,
  withEmbeddedImage,
  withAtomicContentMarker,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { buildInlinePictureDrawing } from '../../../../core/src/store/package/drawing-package-edit';
import { BODY, createPeerHarness, walk, zipDocument } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (character) => character.charCodeAt(0)
);
const decodePort = { decode: async () => ({ pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 }) };
const document = zipDocument(
  '<w:p><w:r><w:t>Image</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>Atomic</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>Neighbor</w:t></w:r></w:p>'
);
function control(pkg: OoxmlPackage) {
  return contentControlsIn(pkg.parts.get(pkg.mainDocumentPart)!.root).find(
    ({ node }) => contentControlPropertiesOf(node).tag === 'test:atomic'
  )?.node;
}
function insertAtomic(pkg: OoxmlPackage, paragraphId: string): OoxmlPackage {
  const part = pkg.parts.get(pkg.mainDocumentPart)!;
  const inserted = applyTreeOp(part, {
    op: 'insertInlineContentControl',
    paragraphId,
    offset: 0,
    tag: 'test:atomic',
    text: 'generated',
    lock: 'contentLocked',
  });
  if (!inserted.ok) throw Error(inserted.reason);
  const next = withPart(pkg, inserted.part);
  const node = control(next)!;
  const marked = replaceNode(inserted.part, node.id, withAtomicContentMarker(node));
  if (!marked.ok) throw Error('Cannot mark the atomic control.');
  return withPart(next, marked.part);
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const batch of ['one-transaction', 'separate-transactions'] as const)
    test(`atomic adoption preserves an ordinary image creation result: ${clients}/${batch}`, async () => {
      const harness = createPeerHarness('atomic-image-identity', undefined, {
        alice: clients[0],
        bob: clients[1],
      });
      try {
        const { alice, bob, pause, resume } = await harness.pair(document);
        let drawingId = '';
        if (batch === 'separate-transactions') {
          const image = await alice.store.insertImage(BODY, {
            paragraphId: harness.paragraphIdAt(alice, 0),
            offset: 0,
            bytes: png,
            mime: 'image/png',
            widthPoints: 12,
            heightPoints: 12,
            decodePort,
            expectedPackageRevision: alice.store.packageRevision,
          });
          if (!image.ok || !image.drawingNodeId) throw Error('Image insertion failed.');
          drawingId = image.drawingNodeId;
        }
        const atomicParagraph = harness.paragraphIdAt(alice, 1);
        expect(
          alice.store.transact(BODY, (ctx) => {
            if (batch === 'one-transaction') {
              const pkg = ctx.currentPackage();
              const embedded = withEmbeddedImage(pkg, pkg.mainDocumentPart, {
                bytes: png,
                mime: 'image/png',
              });
              if (!embedded.ok) throw Error(embedded.reason);
              expect(ctx.applyPackage(() => embedded.pkg)).toBe(true);
              const drawing = buildInlinePictureDrawing({
                docPrId: embedded.docPrId,
                relationshipId: embedded.relationshipId,
                extentEmu: { cx: 152400, cy: 152400 },
              });
              if (drawing.kind !== 'drawing') throw Error('Expected a picture drawing.');
              expect(
                ctx.apply({
                  op: 'insertDrawing',
                  paragraphId: harness.paragraphIdAt(alice, 0),
                  offset: 0,
                  drawing,
                })
              ).toBe(true);
              const current = ctx.currentPackage();
              walk(current.parts.get(current.mainDocumentPart)!.root, (node) => {
                if (node.kind === 'drawing') drawingId = node.id;
              });
            }
            expect(ctx.applyPackage((pkg) => insertAtomic(pkg, atomicParagraph))).toBe(true);
          }).ok
        ).toBe(true);
        alice.port.flushPendingJournals();
        const a = harness.packageOf(alice),
          b = harness.packageOf(bob);
        expect(control(a)?.id).toBe(control(b)?.id);
        expect(control(a)?.id).toMatch(/^lid:/);
        expect(findNode(a.parts.get(a.mainDocumentPart)!, drawingId)?.kind).toBe('drawing');
        harness.expectConverged(alice, bob);
        const born = packageFingerprint(a);
        // Use the original creation result. Looking up a replacement ID would hide the defect.
        const deletion = alice.room.session.beginUndoGroup();
        try {
          expect(alice.store.deleteImage(BODY, drawingId).ok).toBe(true);
        } finally {
          deletion.end();
        }
        harness.expectConverged(alice, bob);
        const deleted = packageFingerprint(harness.packageOf(alice));
        expect(deleted).not.toBe(born);
        expect(alice.room.session.undo()).toBe(true);
        expect(packageFingerprint(harness.packageOf(alice))).toBe(born);
        expect(alice.room.session.redo()).toBe(true);
        expect(packageFingerprint(harness.packageOf(alice))).toBe(deleted);
        // A later remote publication adopts remaining local aliases through the normal path.
        pause();
        harness.apply(bob, [
          {
            op: 'insertText',
            paragraphId: harness.paragraphIdAt(bob, 2),
            offset: 8,
            text: ' peer',
          },
        ]);
        resume();
        harness.expectConverged(alice, bob);
        expect(control(harness.packageOf(alice))?.id).toBe(control(harness.packageOf(bob))?.id);
        expect(saveReopenDigest(harness.packageOf(alice))).toEqual(
          saveReopenDigest(harness.packageOf(bob))
        );
      } finally {
        harness.cleanup();
      }
    });
