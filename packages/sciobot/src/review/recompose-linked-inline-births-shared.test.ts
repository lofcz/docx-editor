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
const native = (pkg: OoxmlPackage) =>
  canonicalOoxmlFingerprint(pkg.parts.get(pkg.mainDocumentPart)!);
const word = (pkg: OoxmlPackage) => {
  const loaded = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!loaded.ok) throw Error(loaded.reason);
  return native(loaded.package);
};
function actorDirect(pkg: OoxmlPackage, plan: Parameters<typeof previewDirectPlan>[1]) {
  const store = new TreePackageStore(pkg, pkg.parts.get(pkg.mainDocumentPart)!);
  const result = store.transact(
    { kind: 'body' },
    (ctx) => {
      if (
        !ctx.applyPackage((working) => {
          const applied = previewDirectPlan(working, plan);
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
function bindWordIdentity(expected: OoxmlPackage, actual: OoxmlPackage) {
  const control = (pkg: OoxmlPackage) => {
    const block = buildBlockIndex(pkg).byId.get('linked-atom')!;
    const node = findNode(pkg.parts.get(block.partName)!, block.controlId)!;
    if (node.kind !== 'contentControl') throw Error('Missing native atom.');
    return node.children
      .find((child) => child.kind === 'contentControlProperties')!
      .children.find((child) => child.localName === 'id')!;
  };
  const a = control(actual),
    e = control(expected);
  if (a.kind !== 'generic' || e.kind !== 'generic') throw Error('Missing native Word ID.');
  const value = a.attributes.find(
    (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'val'
  )!.value;
  expect(value).toMatch(/^[1-9][0-9]*$/);
  const replaced = replaceNode(expected.parts.get(expected.mainDocumentPart)!, e.id, {
    ...e,
    attributes: e.attributes.map((attribute) =>
      attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'val'
        ? { ...attribute, value }
        : attribute
    ),
  });
  if (!replaced.ok) throw Error('Cannot bind fresh Word identity.');
  return withPart(expected, replaced.part);
}
for (const peerOffset of [1, 10])
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`linked atom birth ${peerOffset}/${clients}/${decision} preserves offline peer text, exact native decisions, and both histories`, async () => {
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
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        });
        try {
          const [alice, bob] = pair.editors;
          const before = alice!.surface!.session.currentPackage(),
            paragraphId = alice!.surface!.session.paragraphIds()[0]!;
          const target =
            'Before. [Lead. <scio-mention key="user" value="Ada" /> Tail.](https://target.test "Keep") After.';
          const initial = await readWorkingView(alice!.surface!.session, propertyToken(alice!));
          if (!initial.success) throw Error(initial.error);
          const options = { allocateBlockId: () => 'linked-atom' };
          const plan = planTargetDocument(
            { pkg: before, revision: propertyToken(alice!), projection: initial.view.projection },
            target,
            options
          );
          if (!plan.success) throw Error(plan.error);
          const direct = actorDirect(before, plan.plan);
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
            { target: options, payloads: { presence: { proposalId: 'linked-presence' } } }
          );
          if (!staged.success) throw Error(staged.error);
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
                  ? 'B peerefore\\. [Lead\\. <scio-mention key="user" value="Ada" /> Tail\\.](https://target.test "Keep") After\\.'
                  : 'Before\\. [Le peerad\\. <scio-mention key="user" value="Ada" /> Tail\\.](https://target.test "Keep") After\\.'
              );
              const groups = readReviewGroups(editor.surface!.session.currentPackage());
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
            const payloads = readPayloadStore(editor.surface!.session.currentPackage());
            expect(payloads.status === 'ready' && [...payloads.entries.keys()]).toEqual(
              decision === 'accept' ? ['linked-atom'] : []
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
              editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
              ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
            });
            try {
              for (const editor of reopened.editors)
                expect(word(editor.surface!.session.currentPackage())).toBe(word(expected));
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
