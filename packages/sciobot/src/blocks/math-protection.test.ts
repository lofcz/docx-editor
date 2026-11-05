import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  ATOMIC_CONTENT_NAMESPACE,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { buildBlockIndex } from './block-index';
import { createMathPayload, type MathKind } from './math-codec';
import { updateMath } from './math-edit';
import { blockTag, withPayloadRecord } from './payload-opc-codec';
import { updateAtomicBlock } from './update-atomic';

const binding =
  '<w:dataBinding w:xpath="/root/value" w:storeItemID="{12345678-1234-1234-1234-123456789ABC}"/>';
function fixture(kind: MathKind, protection: string): Uint8Array {
  const payload = createMathPayload('protected-math', kind, 'x');
  const text = '<w:r><w:t>x</w:t></w:r>';
  const control =
    `<w:sdt xmlns:a="${ATOMIC_CONTENT_NAMESPACE}" a:atomicContent="1"><w:sdtPr>` +
    `<w:tag w:val="${blockTag(payload.id)}"/><w:lock w:val="contentLocked"/>` +
    (protection === 'own-binding' ? binding : '') +
    '</w:sdtPr><w:sdtContent>' +
    (kind === 'inline' ? text : `<w:p>${text}</w:p>`) +
    '</w:sdtContent></w:sdt>';
  const outerProperties =
    protection === 'ancestor-binding'
      ? binding
      : protection === 'own-binding'
        ? ''
        : `<w:lock w:val="${protection}"/>`;
  const bytes = docxFixture(
    `<w:sdt><w:sdtPr>${outerProperties}</w:sdtPr><w:sdtContent>` +
      (kind === 'inline' ? `<w:p>${control}</w:p>` : control) +
      '</w:sdtContent></w:sdt><w:p><w:r><w:t>Neighbor</w:t></w:r></w:p>'
  );
  const read = readOoxmlPackage(bytes);
  if (!read.ok) throw Error(read.reason);
  const stored = withPayloadRecord(read.package, payload, 'Equation');
  if (!stored.ok) throw Error(stored.reason);
  return writeOoxmlPackage(stored.pkg);
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const kind of ['inline', 'block'] as const)
    for (const protection of [
      'contentLocked',
      'sdtContentLocked',
      'ancestor-binding',
      'own-binding',
    ])
      test(`math source and display respect native protection: ${clients}/${kind}/${protection}`, async () => {
        const pair = await createTwoEditorFixture({
          document: fixture(kind, protection),
          clientIds: clients,
        });
        try {
          for (const [index, editor] of pair.editors.entries()) {
            const surface = editor.surface!;
            const before = surface.session.currentPackage();
            const fingerprint = canonicalOoxmlFingerprint(surface.session.part());
            const history = [editor.snapshot().canUndo, editor.snapshot().canRedo];
            const shared = Y.encodeStateAsUpdate(pair.documents[index]!);
            const original = buildBlockIndex(before).byId.get('protected-math')!.payload;
            for (const target of [kind, kind === 'inline' ? 'block' : 'inline'] as const)
              expect(
                updateMath(editor, { original, source: 'changed', kind: target })
              ).toMatchObject({
                success: false,
                code: 'unsupported-operation',
              });
            expect(
              updateAtomicBlock(editor, {
                blockId: original.id,
                attributes: { ...original.attributes, texExpression: 'changed' },
              })
            ).toMatchObject({ success: false, code: 'unsupported-operation' });
            expect(surface.session.currentPackage()).toBe(before);
            expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(fingerprint);
            expect([editor.snapshot().canUndo, editor.snapshot().canRedo]).toEqual(history);
            expect(Y.encodeStateAsUpdate(pair.documents[index]!)).toEqual(shared);
          }
          const bob = pair.editors[1]!.surface!;
          const paragraphId = bob.session
            .paragraphIds()
            .find((id) => paragraphTextOf(bob.session.part(), id) === 'Neighbor')!;
          const point = { paragraphId, offset: 8 };
          bob.setSelection({ anchor: point, head: point });
          bob.insertPlainText(' peer');
          bob.flushPendingInput();
          await pair.settle();
          for (const editor of pair.editors) {
            expect(paragraphTextOf(editor.surface!.session.part(), paragraphId)).toBe(
              'Neighbor peer'
            );
            expect(
              buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('protected-math')!
                .payload.attributes.texExpression
            ).toBe('x');
          }
          const saved = await pair.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
        } finally {
          pair.destroy();
        }
      });

test('an ancestor deletion lock still permits equation content edits and display conversion', async () => {
  const pair = await createTwoEditorFixture({ document: fixture('inline', 'sdtLocked') });
  try {
    const editor = pair.editors[0]!;
    const original = buildBlockIndex(editor.surface!.session.currentPackage()).byId.get(
      'protected-math'
    )!.payload;
    expect(updateMath(editor, { original, source: 'y', kind: 'block' })).toMatchObject({
      success: true,
    });
    await pair.settle();
    for (const peer of pair.editors)
      expect(
        buildBlockIndex(peer.surface!.session.currentPackage()).byId.get('protected-math')!.payload
          .attributes.texExpression
      ).toBe('y');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    for (const peer of pair.editors)
      expect(
        buildBlockIndex(peer.surface!.session.currentPackage()).byId.get('protected-math')!.payload
          .attributes.texExpression
      ).toBe('x');
  } finally {
    pair.destroy();
  }
});
