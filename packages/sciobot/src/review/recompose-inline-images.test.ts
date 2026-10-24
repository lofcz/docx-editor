import { previewDirectPlan } from '../ai/apply-direct';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { planTargetDocument } from '../diff/plan-target';
import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  replaceNode,
  withPart,
  WML_NAMESPACE_URI,
  readOoxmlPackage,
  writeOoxmlPackage,
  type ImageDecodePort,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { buildBlockIndex } from '../blocks/block-index';
import { readPayloadStore } from '../blocks/payload-opc-codec';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { readReviewGroups } from './groups';
import { recomposeOwnedProposals } from './recompose';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
const context = {
  documentId: 'tracked-images',
  epoch: 'native',
  author: 'AI',
  date: '2026-10-24T12:00:00.000Z',
};
const png = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const decodePort: ImageDecodePort = {
  async decode() {
    return { pixelWidth: 1, pixelHeight: 1, dpiX: 96, dpiY: 96 };
  },
};
const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes].filter(([name]) => /(^|\/)media\//.test(name));
// Bind only fresh native control identities into the independent direct execution.
function expectedImageBirth(direct: OoxmlPackage, pending: OoxmlPackage): OoxmlPackage {
  const actual = [...buildBlockIndex(pending).byId.values()][0]!;
  const expected = [...buildBlockIndex(direct).byId.values()][0]!;
  const control = (pkg: OoxmlPackage, id: string) => {
    const node = findNode(pkg.parts.get(pkg.mainDocumentPart)!, id);
    if (node?.kind !== 'contentControl') throw Error('Expected one image control.');
    return node;
  };
  const a = control(pending, actual.controlId);
  const e = control(direct, expected.controlId);
  let part = direct.parts.get(direct.mainDocumentPart)!;
  for (const name of ['id', 'tag']) {
    const property = (node: typeof a) =>
      node.children
        .find((child) => child.kind === 'contentControlProperties')!
        .children.find((child) => child.localName === name)!;
    const source = property(a),
      target = property(e);
    if (!source || !target) {
      expect(Boolean(source)).toBe(Boolean(target));
      continue;
    }
    if (source.kind !== 'generic' || target.kind !== 'generic')
      throw Error('Missing image control identity.');
    const value = source.attributes.find(
      (attr) => attr.namespaceUri === WML_NAMESPACE_URI && attr.localName === 'val'
    )!.value;
    if (name === 'id') expect(value).toMatch(/^[1-9][0-9]*$/);
    const bound = replaceNode(part, target.id, {
      ...target,
      attributes: target.attributes.map((attr) =>
        attr.namespaceUri === WML_NAMESPACE_URI && attr.localName === 'val'
          ? { ...attr, value }
          : attr
      ),
    });
    if (!bound.ok) throw Error('Cannot bind new image identity.');
    part = bound.part;
  }
  return withPart(direct, part);
}
for (const born of [false, true])
  for (const mark of ['plain', 'bold', 'link', 'code-link'] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`tracked image ${born ? 'born' : 'retained'}/${mark}/${decision}`, async () => {
        const editor = createDocxEditor({
          container: document.createElement('div'),
          imageDecodePort: decodePort,
          document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
          modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
        });
        try {
          const surface = editor.surface!;
          const originalPackage = surface.session.currentPackage();
          const before = native(originalPackage);
          const image = '![Cat](asset:cat "Photo")';
          const fragment =
            mark === 'plain'
              ? image
              : mark === 'bold'
                ? `**${image}**`
                : mark === 'link'
                  ? `[${image}](https://target.test "Read")`
                  : `<code>[${image}](https://target.test "Read")</code>`;
          const target = born ? `Keep.\n\nA ${fragment} Z` : `Keep. A ${fragment} Z`;
          const revision = () => readEditorRevision(editor, context.documentId, context.epoch)!;
          const initial = await readWorkingView(surface.session, revision());
          if (!initial.success) throw Error(initial.error);
          const planned = planTargetDocument(
            { pkg: originalPackage, revision: revision(), projection: initial.view.projection },
            target
          );
          if (!planned.success) throw Error(planned.error);
          const prepared = await prepareInlineImageResources(
            planned.plan,
            async () => ({ bytes: png, mime: 'image/png', widthPoints: 24, heightPoints: 18 }),
            decodePort
          );
          if (!prepared.success) throw Error(prepared.error);
          const direct = previewDirectPlan(originalPackage, planned.plan, prepared.images);
          if (!direct.success) throw Error(direct.error);
          let resolutions = 0;
          const staged = await recomposeOwnedProposals(
            editor,
            { revision: revision(), targetMarkdown: target, ownership: [] },
            context,
            {
              payloads: { presence: { proposalId: 'image-birth' } },
              images: {
                decodePort,
                resolve: async (reference) => {
                  expect(reference).toBe('asset:cat');
                  resolutions++;
                  return { bytes: png, mime: 'image/png', widthPoints: 24, heightPoints: 18 };
                },
              },
            }
          );
          expect(staged).toMatchObject({ success: true, committed: true });
          expect(resolutions).toBe(1);
          const pending = surface.session.currentPackage();
          const pendingNative = native(pending);
          const expectedAccepted = expectedImageBirth(direct.pkg, pending);
          const expectedEditor = createDocxEditor({
            container: document.createElement('div'),
            document: writeOoxmlPackage(direct.pkg),
            imageDecodePort: decodePort,
          });
          let expectedMarkdown: string;
          try {
            const expectedView = await readWorkingView(
              expectedEditor.surface!.session,
              readEditorRevision(expectedEditor, context.documentId, context.epoch)!
            );
            if (!expectedView.success) throw Error(expectedView.error);
            expectedMarkdown = expectedView.view.projection.text;
          } finally {
            expectedEditor.destroy();
          }
          const blocks = buildBlockIndex(pending);
          expect(blocks.byId.size).toBe(1);
          expect([...blocks.byId.values()][0]!.payload.type).toBe('inline_image');
          expect([...blocks.byId.values()][0]!.payload.attributes).toMatchObject({
            reference: 'asset:cat',
            alt: 'Cat',
            title: 'Photo',
            widthPoints: 24,
            heightPoints: 18,
          });
          expect(media(pending)).toHaveLength(1);
          expect(media(pending)[0]![1]).toEqual(png);
          const original = await readCommittedView(surface.session, revision());
          expect(original.success && original.view.projection.text).toBe(
            initial.view.projection.text
          );
          const working = await readWorkingView(surface.session, revision());
          expect(working.success && working.view.projection.text).toBe(expectedMarkdown);
          const groups = readReviewGroups(pending);
          expect(groups.success && groups.groups.every((group) => !group.readOnly)).toBe(true);
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(before);
          expect(media(surface.session.currentPackage())).toHaveLength(0);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(pendingNative);
          expect(media(surface.session.currentPackage())[0]![1]).toEqual(png);
          const decide = decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups;
          expect(decide(editor, revision(), context)).toMatchObject({ success: true });
          const resolved = surface.session.currentPackage();
          const resolvedNative = native(resolved);
          if (decision === 'reject') {
            expect(resolvedNative).toBe(before);
            expect(media(resolved)).toHaveLength(0);
            const payloads = readPayloadStore(resolved);
            expect(payloads.status === 'ready' && payloads.entries.size).toBe(0);
          } else {
            expect(media(resolved)).toHaveLength(1);
            expect(media(resolved)[0]![1]).toEqual(png);
            const accepted = await readWorkingView(surface.session, revision());
            expect(accepted.success && accepted.view.projection.text).toBe(expectedMarkdown);
            expect(resolvedNative).toBe(native(expectedAccepted));
          }
          expect(editor.exec({ type: 'undo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(pendingNative);
          expect(media(surface.session.currentPackage())[0]![1]).toEqual(png);
          expect(editor.exec({ type: 'redo' }).ok).toBe(true);
          expect(native(surface.session.currentPackage())).toBe(resolvedNative);
          const reopened = readOoxmlPackage(writeOoxmlPackage(surface.session.currentPackage()));
          expect(reopened.ok).toBe(true);
          if (!reopened.ok) throw Error(reopened.reason);
          expect(native(reopened.package)).toBe(resolvedNative);
          const reopenedAgain = readOoxmlPackage(writeOoxmlPackage(reopened.package));
          if (!reopenedAgain.ok) throw Error(reopenedAgain.reason);
          expect(native(reopenedAgain.package)).toBe(native(reopened.package));
          expect(media(reopened.package).map(([, bytes]) => bytes)).toEqual(
            media(resolved).map(([, bytes]) => bytes)
          );
        } finally {
          editor.destroy();
        }
      });

test('image preparation refuses stale publication without adding media or payloads', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    imageDecodePort: decodePort,
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
  try {
    const surface = editor.surface!;
    let afterManual: OoxmlPackage | undefined;
    const result = await recomposeOwnedProposals(
      editor,
      {
        revision: readEditorRevision(editor, context.documentId, context.epoch)!,
        targetMarkdown: 'Keep. ![Cat](asset:cat)',
        ownership: [],
      },
      context,
      {
        payloads: { presence: { proposalId: 'stale-image' } },
        images: {
          decodePort,
          resolve: async () => {
            expect(
              surface.applyAutomationOps(() => [
                {
                  op: 'insertText',
                  paragraphId: surface.session.paragraphIds()[0]!,
                  offset: 5,
                  text: ' manual',
                },
              ])
            ).toMatchObject({ committed: true });
            afterManual = surface.session.currentPackage();
            return { bytes: png, mime: 'image/png' };
          },
        },
      }
    );
    expect(result).toMatchObject({ success: false, code: 'stale-revision' });
    expect(surface.session.currentPackage()).toBe(afterManual!);
    expect(media(afterManual!)).toHaveLength(0);
    expect(buildBlockIndex(afterManual!).byId.size).toBe(0);
  } finally {
    editor.destroy();
  }
});
