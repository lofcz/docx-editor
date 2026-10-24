import { expectImageRelationships } from '../../test/support/image-relationships';
import {
  prepareInlineImageResources,
  type PreparedInlineImages,
} from '../native/inline-image-resources';
// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  TreePackageStore,
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  withPart,
  replaceNode,
  WML_NAMESPACE_URI,
  type ImageDecodePort,
  writeOoxmlPackage,
  readOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import { withInitializedPayloadStore, readPayloadStore } from '../blocks/payload-opc-codec';
import { buildBlockIndex } from '../blocks/block-index';
import { planTargetDocument } from '../diff/plan-target';
import { previewDirectPlan } from '../ai/apply-direct';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { recomposeOwnedProposals } from './recompose';
import { readWorkingView } from './working-view';
import { readReviewGroups } from './groups';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
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
const images = {
  decodePort,
  resolve: async () => ({
    bytes: png,
    mime: 'image/png' as const,
    widthPoints: 24,
    heightPoints: 18,
  }),
};
const media = (pkg: OoxmlPackage) =>
  [...pkg.partBytes]
    .filter(([name]) => /(^|\/)media\//.test(name))
    .sort(([a], [b]) => a.localeCompare(b));
const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
const word = (pkg: OoxmlPackage) => {
  const loaded = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!loaded.ok) throw Error(loaded.reason);
  return native(loaded.package);
};
function actorDirect(
  pkg: OoxmlPackage,
  plan: Parameters<typeof previewDirectPlan>[1],
  resources: PreparedInlineImages
) {
  const store = new TreePackageStore(pkg, pkg.parts.get(pkg.mainDocumentPart)!);
  const result = store.transact(
    { kind: 'body' },
    (ctx) => {
      if (
        !ctx.applyPackage((working) => {
          const applied = previewDirectPlan(working, plan, resources);
          if (!applied.success) throw Error(applied.error);
          return applied.pkg;
        })
      )
        throw Error('Independent direct transaction refused.');
    },
    { recordsHistory: false, actorId: 'alice' }
  );
  if (!result.ok) throw Error(result.reason);
  return store.currentPackage();
}
function bindWordIdentity(direct: OoxmlPackage, pending: OoxmlPackage): OoxmlPackage {
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
for (const peerOffset of [1, 10])
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`linked image birth ${peerOffset}/${clients}/${decision} preserves offline peer text, exact native decisions, and both histories`, async () => {
        const loaded = readOoxmlPackage(
          docxFixture('<w:p><w:r><w:t>Before. Lead.  Tail. After.</w:t></w:r></w:p>')
        );
        if (!loaded.ok) throw Error(loaded.reason);
        const prepared = withInitializedPayloadStore(loaded.package);
        if (!prepared.ok) throw Error(prepared.reason);
        const pair = await createTwoEditorFixture({
          document: writeOoxmlPackage(prepared.pkg),
          documentId: propertyContext.documentId,
          ownedAttributes: true,
          clientIds: clients,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE], imageDecodePort: decodePort },
        });
        try {
          const [alice, bob] = pair.editors;
          const before = alice!.surface!.session.currentPackage(),
            paragraphId = alice!.surface!.session.paragraphIds()[0]!;
          const target =
            'Before. [Lead. ![Cat](asset:cat "Photo") Tail.](https://target.test "Keep") After.';
          const initial = await readWorkingView(alice!.surface!.session, propertyToken(alice!));
          if (!initial.success) throw Error(initial.error);
          const options = { allocateBlockId: () => 'linked-atom' };
          const plan = planTargetDocument(
            { pkg: before, revision: propertyToken(alice!), projection: initial.view.projection },
            target,
            options
          );
          if (!plan.success) throw Error(plan.error);
          const resources = await prepareInlineImageResources(
            plan.plan,
            images.resolve,
            decodePort
          );
          if (!resources.success) throw Error(resources.error);
          const direct = actorDirect(before, plan.plan, resources.images);
          const peerOp = {
            op: 'insertText',
            paragraphId,
            offset: peerOffset,
            text: ' peer',
          } as const;
          const addPeer = (pkg: OoxmlPackage) => {
            const edited = applyTreeOp(pkg.parts.get(pkg.mainDocumentPart)!, peerOp);
            if (!edited.ok) throw Error(edited.reason);
            return withPart(pkg, edited.part);
          };
          pair.pause();
          expect(bob!.surface!.applyAutomationOps(() => [peerOp]).rejected).toBe(false);
          const staged = await recomposeOwnedProposals(
            alice!,
            { revision: propertyToken(alice!), targetMarkdown: target, ownership: [] },
            propertyContext,
            { target: options, payloads: { presence: { proposalId: 'linked-presence' } }, images }
          );
          if (!staged.success) throw Error(staged.error);
          const imageId = [
            ...buildBlockIndex(alice!.surface!.session.currentPackage()).byId.keys(),
          ][0]!;
          expect(media(alice!.surface!.session.currentPackage())).toHaveLength(1);
          const acceptedExpected = addPeer(
            bindWordIdentity(direct, alice!.surface!.session.currentPackage())
          );
          const rejectedExpected = addPeer(before);
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const check = async () => {
            const states = pair.editors.map((editor) =>
              native(editor.surface!.session.currentPackage())
            );
            expect(states[0]).toBe(states[1]);
            for (const editor of pair.editors) {
              const view = await readWorkingView(editor.surface!.session, propertyToken(editor));
              expect(view.success && view.view.projection.text).toBe(
                peerOffset === 1
                  ? 'B peerefore\\. [Lead\\. ![Cat](asset:cat "Photo") Tail\\.](https://target.test "Keep") After\\.'
                  : 'Before\\. [Le peerad\\. ![Cat](asset:cat "Photo") Tail\\.](https://target.test "Keep") After\\.'
              );
              const pkg = editor.surface!.session.currentPackage();
              expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
              const groups = readReviewGroups(pkg);
              expect(groups.success && groups.groups.length).toBe(2);
              expect(groups.success && groups.groups.every((group) => !group.readOnly)).toBe(true);
            }
          };
          await check();
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(word(editor.surface!.session.currentPackage())).toBe(word(rejectedExpected));
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check();
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check();
          const pending = pair.editors.map((editor) =>
            native(editor.surface!.session.currentPackage())
          );
          const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
            alice!,
            propertyToken(alice!),
            propertyContext
          );
          if (!decided.success) throw Error(decided.error);
          await pair.settle();
          const expected = decision === 'accept' ? acceptedExpected : rejectedExpected;
          for (const editor of pair.editors) {
            expect(word(editor.surface!.session.currentPackage())).toBe(word(expected));
            expect(media(editor.surface!.session.currentPackage())).toEqual(
              decision === 'accept' ? media(acceptedExpected) : []
            );
            const pkg = editor.surface!.session.currentPackage();
            expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
            const payloads = readPayloadStore(pkg);
            expect(payloads.status === 'ready' && [...payloads.entries.keys()]).toEqual(
              decision === 'accept' ? [imageId] : []
            );
          }
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          pair.editors.forEach((editor, i) =>
            expect(native(editor.surface!.session.currentPackage())).toBe(pending[i]!)
          );
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          bob!.surface!.setSelection({
            anchor: { paragraphId, offset: 1 },
            head: { paragraphId, offset: 1 },
          });
          await pair.settle();
          const geometry = pair.editors.map((editor) =>
            caretAt(editor.surface!.layout(), { paragraphId, offset: 1 })
          );
          expect(geometry[0]).not.toBeNull();
          expect(geometry[0]).toEqual(geometry[1]);
          const saved = new Uint8Array(await bob!.save());
          for (const room of [false, true]) {
            const reopened = await createTwoEditorFixture({
              document: saved,
              documentId: propertyContext.documentId,
              ownedAttributes: true,
              editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE], imageDecodePort: decodePort },
              ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
            });
            try {
              for (const editor of reopened.editors) {
                expect(word(editor.surface!.session.currentPackage())).toBe(word(expected));
                expect(media(editor.surface!.session.currentPackage())).toEqual(
                  decision === 'accept' ? media(acceptedExpected) : []
                );
              }
              reopened.ownership.forEach((handle) => handle.assertHealthy());
            } finally {
              reopened.destroy();
            }
          }
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      });
