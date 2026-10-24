import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  TreePackageStore,
  applyTreeOp,
  findNode,
  replaceNode,
  parentNodeOf,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  readOoxmlPart,
  bodyStoryRoot,
  insertChildren,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  context,
  png,
  decodePort,
  media,
  native,
  bindControls,
  cases,
  preparedCellDocument,
} from '../../test/support/tracked-cell-atoms';
import { expectImageRelationships } from '../../test/support/image-relationships';
import { buildBlockIndex } from '../blocks/block-index';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from './working-view';
import { planTargetDocument } from '../diff/plan-target';
import { prepareInlineImageResources } from '../native/inline-image-resources';
import { previewDirectPlan } from '../ai/apply-direct';
import { recomposeOwnedProposals } from './recompose';
import { acceptAllReviewGroups } from './accept-all';
import { rejectAllReviewGroups } from './reject-all';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
const word = (pkg: OoxmlPackage) => {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return native(read.package);
};
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const decision of ['accept', 'reject'] as const)
    test(`peer text inside an atom-bearing new row survives creation history and ${decision}: ${clients}`, async () => {
      const item = cases[0]!;
      const pair = await createTwoEditorFixture({
        document: preparedCellDocument(item.body),
        documentId: context.documentId,
        clientIds: clients,
        ownedAttributes: true,
        editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE], imageDecodePort: decodePort },
      });
      try {
        const [alice, bob] = pair.editors;
        const token = (editor: typeof alice) =>
          readEditorRevision(editor!, context.documentId, context.epoch)!;
        const before = alice!.surface!.session.currentPackage(),
          initialToken = token(alice);
        const initial = await readWorkingView(alice!.surface!.session, initialToken);
        if (!initial.success) throw Error(initial.error);
        const staged = await recomposeOwnedProposals(
          alice!,
          { revision: token(alice), targetMarkdown: item.target, ownership: [] },
          context,
          {
            payloads: { presence: { proposalId: 'peer-cell-presence' } },
            images: { decodePort, resolve: async () => ({ bytes: png, mime: 'image/png' }) },
          }
        );
        if (!staged.success) throw Error(staged.error);
        await pair.settle();
        const index = buildBlockIndex(bob!.surface!.session.currentPackage());
        let owner = parentNodeOf(
          bob!.surface!.session.part(),
          [...index.byId.values()][0]!.controlId
        );
        while (owner && owner.kind !== 'paragraph')
          owner = parentNodeOf(bob!.surface!.session.part(), owner.id);
        if (owner?.kind !== 'paragraph') throw Error('Missing born paragraph.');
        const paragraphId = owner.id,
          offset = paragraphTextOf(bob!.surface!.session.part(), paragraphId)!.length;
        pair.pause();
        expect(
          bob!.surface!.applyAutomationOps(
            () => [{ op: 'insertText', paragraphId, offset, text: 'Peer', bias: 'right' }],
            { kind: 'body' },
            [],
            false,
            true
          )
        ).toMatchObject({ committed: true, rejected: false });
        pair.reconnect(clients[0] === 41 ? 0 : 1);
        await pair.settle();
        const pending = alice!.surface!.session.currentPackage(),
          pendingWord = word(pending);
        const direct = async (target: string) => {
          const plan = planTargetDocument(
            { pkg: before, revision: initialToken, projection: initial.view.projection },
            target
          );
          if (!plan.success) throw Error(plan.error);
          const images = await prepareInlineImageResources(
            plan.plan,
            async () => ({ bytes: png, mime: 'image/png' }),
            decodePort
          );
          if (!images.success) throw Error(images.error);
          const store = new TreePackageStore(before, before.parts.get(before.mainDocumentPart)!);
          const applied = store.transact(
            { kind: 'body' },
            (ctx) => {
              if (
                !ctx.applyPackage((pkg) => {
                  const result = previewDirectPlan(pkg, plan.plan, images.images);
                  if (!result.success) throw Error(result.error);
                  return result.pkg;
                })
              )
                throw Error('Direct oracle refused.');
            },
            {
              recordsHistory: false,
              actorId: alice!.surface!.collaborationSession()!.identity.actorId,
            }
          );
          if (!applied.ok) throw Error(applied.reason);
          return store.currentPackage();
        };
        const accepted = bindControls(
          await direct(item.target.replace(' /> |', ' />Peer |')),
          pending,
          before
        );
        const originalPart = before.parts.get(before.mainDocumentPart)!;
        const originalBody = bodyStoryRoot(originalPart);
        if (!originalBody || originalBody.kind === 'textValue')
          throw Error('Missing original body.');
        const originalTable = originalBody.children.find((node) => node.kind === 'table');
        if (!originalTable || originalTable.kind !== 'table')
          throw Error('Missing original table.');
        const rows = originalTable.children.filter((node) => node.kind === 'tableRow');
        const rowBirth = applyTreeOp(originalPart, {
          op: 'insertTableRow',
          tableId: originalTable.id,
          rowId: rows.at(-1)!.id,
          where: 'below',
        });
        if (!rowBirth.ok) throw Error(rowBirth.reason);
        const table = findNode(rowBirth.part, originalTable.id);
        if (table?.kind !== 'table') throw Error('Missing recovered row oracle.');
        const row = table.children.filter((node) => node.kind === 'tableRow').at(-1)!;
        if (row.kind !== 'tableRow') throw Error('Missing row oracle.');
        const cell = row.children.find((node) => node.kind === 'tableCell')!;
        if (cell.kind !== 'tableCell') throw Error('Missing cell oracle.');
        const paragraph = cell.children.find((node) => node.kind === 'paragraph')!;
        if (paragraph.kind !== 'paragraph') throw Error('Missing paragraph oracle.');
        const identity = owner.attributes.filter(
          (a) => a.namespaceUri === 'http://schemas.microsoft.com/office/word/2010/wordml'
        );
        const bound = replaceNode(rowBirth.part, paragraph.id, {
          ...paragraph,
          attributes: identity,
        });
        if (!bound.ok) throw Error('Cannot bind retained peer identity.');
        const companion = row.children.filter((node) => node.kind === 'tableCell')[1]!;
        if (companion.kind !== 'tableCell') throw Error('Missing companion cell oracle.');
        const companionParagraph = companion.children.find((node) => node.kind === 'paragraph')!;
        const pendingCompanionId = alice!
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(alice!.surface!.session.part(), id) === 'Tail')!;
        const pendingCompanion = findNode(
          pending.parts.get(pending.mainDocumentPart)!,
          pendingCompanionId
        );
        if (pendingCompanion?.kind !== 'paragraph' || companionParagraph.kind !== 'paragraph')
          throw Error('Missing born companion identity.');
        const companionIdentity = pendingCompanion.attributes.filter(
          (a) => a.namespaceUri === 'http://schemas.microsoft.com/office/word/2010/wordml'
        );
        const boundCompanion = replaceNode(bound.part, companionParagraph.id, {
          ...companionParagraph,
          attributes: companionIdentity,
        });
        if (!boundCompanion.ok) throw Error('Cannot bind born companion identity.');
        const peerText = applyTreeOp(boundCompanion.part, {
          op: 'insertText',
          paragraphId: paragraph.id,
          offset: 0,
          text: 'Peer',
          bias: 'right',
        });
        if (!peerText.ok) throw Error(peerText.reason);
        const recovery = withPart(before, peerText.part);
        const check = (expected: string, blocks: number) => {
          for (const editor of pair.editors) {
            const pkg = editor.surface!.session.currentPackage();
            expect(word(pkg)).toBe(expected);
            expect(native(pkg)).toBe(native(pair.editors[0]!.surface!.session.currentPackage()));
            expect(buildBlockIndex(pkg).byId.size).toBe(blocks);
            expect(media(pkg).map(([, bytes]) => bytes)).toEqual(blocks ? [png] : []);
            expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
          }
        };
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(word(recovery), 0);
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(pendingWord, 4);
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(paragraphTextOf(alice!.surface!.session.part(), paragraphId)?.endsWith('Peer')).toBe(
          false
        );
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(pendingWord, 4);
        const sourcePart = before.parts.get(before.mainDocumentPart)!,
          body = bodyStoryRoot(sourcePart);
        if (!body || body.kind === 'textValue') throw Error('Missing original body.');
        const peer = readOoxmlPart(
          '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr/><w:r><w:t xml:space="preserve">Peer</w:t></w:r></w:p>',
          { name: '/peer-oracle.xml', contentType: sourcePart.contentType }
        );
        if (!peer.ok) throw Error(peer.reason);
        const inserted = insertChildren(sourcePart, body.id, 0, [
          { ...peer.part.root, attributes: identity },
        ]);
        if (!inserted.ok) throw Error('Cannot construct peer rejection oracle.');
        const decided = (decision === 'accept' ? acceptAllReviewGroups : rejectAllReviewGroups)(
          bob!,
          token(bob),
          context
        );
        if (!decided.success) throw Error(decided.error);
        await pair.settle();
        // Recovery preserves the original paragraph's Word identity, captured before the decision.
        const rejectedBase = withPart(before, inserted.part);
        const rejected = rejectedBase;
        const expected = decision === 'accept' ? word(accepted) : word(rejected),
          blocks = decision === 'accept' ? 4 : 0;
        check(expected, blocks);
        expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        check(pendingWord, 4);
        expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        check(expected, blocks);
        const cold = await createTwoEditorFixture({
          document: new Uint8Array(await bob!.save()),
          documentId: context.documentId,
          roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!),
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE], imageDecodePort: decodePort },
        });
        try {
          for (const editor of cold.editors) {
            const pkg = editor.surface!.session.currentPackage();
            expect(word(pkg)).toBe(expected);
            expect(buildBlockIndex(pkg).byId.size).toBe(blocks);
            expectImageRelationships(pkg, buildBlockIndex(pkg).byId.values());
          }
          cold.ownership.forEach((h) => h.assertHealthy());
        } finally {
          cold.destroy();
        }
        pair.ownership.forEach((h) => h.assertHealthy());
      } finally {
        pair.destroy();
      }
    }, 30000);
