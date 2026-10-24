import { caretAt } from '@docx-editor.dev/core/layout';
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  TreePackageStore,
  replaceChildren,
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
  atoms,
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
for (const site of [
  { name: 'math-image', offset: 2, marker: '![' },
  { name: 'image-mention', offset: 4, marker: '[Grace]' },
  { name: 'mention-source', offset: 8, marker: '<source' },
])
  for (const clients of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const decision of ['accept', 'reject'] as const)
      test(`peer text between ${site.name} atoms survives table replacement and ${decision}: ${clients}`, async () => {
        const item = {
          body: '<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>',
          target: `Before\n\n| A | B |\n| - | - |\n| ${atoms} | Tail |\n\nAfter`,
        };
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
            offset = site.offset;
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
          const firstPendingWord = word(bob!.surface!.session.currentPackage());
          const updated = item.target
            .replace('$x$', '$y$')
            .replace('Cat', 'Lynx')
            .replace('Ada', 'Grace')
            .replace('Paper', 'Updated');
          const second = await recomposeOwnedProposals(
            alice!,
            {
              revision: token(alice),
              targetMarkdown: updated,
              ownership: staged.ownership,
              contentProofs: staged.contentProofs,
            },
            context,
            {
              payloads: { presence: { proposalId: 'peer-second-table-presence' } },
              images: { decodePort, resolve: async () => ({ bytes: png, mime: 'image/png' }) },
            }
          );
          if (!second.success) throw Error(second.error);
          pair.reconnect(clients[0] === 41 ? 0 : 1);
          await pair.settle();
          const pending = alice!.surface!.session.currentPackage(),
            pendingWord = word(pending);
          const working = await readWorkingView(alice!.surface!.session, token(alice));
          if (!working.success) throw Error(working.error);
          const replan = planTargetDocument(
            { pkg: pending, revision: token(alice), projection: working.view.projection },
            updated.replace(site.marker, 'Peer' + site.marker)
          );
          if (!replan.success) throw Error(replan.error);
          expect(replan.plan.operations).toEqual([]);
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
          // Empty native recovery slots remain private addresses outside the new table.
          // Export hides only those slots. Bind fresh Word identities before any decision.
          const visible = readOoxmlPackage(writeOoxmlPackage(pending));
          if (!visible.ok) throw Error(visible.reason);
          let directTarget = await direct(updated.replace(site.marker, 'Peer' + site.marker));
          const directPart = directTarget.parts.get(directTarget.mainDocumentPart)!;
          const firstControl = [...buildBlockIndex(directTarget).byId.values()][0]!;
          let directOwner = parentNodeOf(directPart, firstControl.controlId);
          while (directOwner && directOwner.kind !== 'paragraph')
            directOwner = parentNodeOf(directPart, directOwner.id);
          if (directOwner?.kind !== 'paragraph')
            throw Error('Missing independent target paragraph.');
          const textOf = (node: import('@docx-editor.dev/core/store').OoxmlNode): string =>
            node.kind === 'textValue' ? node.value : node.children.map(textOf).join('');
          const combined = directOwner.children.find(
            (node) => node.kind === 'run' && textOf(node) === ' Peer'
          );
          if (!combined) throw Error('Missing independent separator and peer run.');
          const split = readOoxmlPart(
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t xml:space="preserve">Peer</w:t></w:r></w:p>',
            { name: '/independent-peer-runs.xml', contentType: directPart.contentType }
          );
          if (!split.ok || split.part.root.kind !== 'paragraph')
            throw Error('Invalid independent peer run oracle.');
          const splitChildren = split.part.root.children;
          const replaced = replaceChildren(
            directPart,
            directOwner.id,
            directOwner.children.flatMap((node) =>
              node.id === combined.id ? splitChildren : [node]
            )
          );
          if (!replaced.ok) throw Error('Cannot bind independent peer run boundaries.');
          directTarget = withPart(directTarget, replaced.part);
          const accepted = bindControls(directTarget, visible.package, before);
          const originalPart = before.parts.get(before.mainDocumentPart)!;
          const originalBody = bodyStoryRoot(originalPart);
          if (!originalBody || originalBody.kind === 'textValue')
            throw Error('Missing original body.');
          const identity = owner.attributes.filter(
            (a) => a.namespaceUri === 'http://schemas.microsoft.com/office/word/2010/wordml'
          );
          expect(identity).toHaveLength(2);
          // Undo keeps only the independently used column and row. Its 4680-twip width
          // is half this fixture's 9360-twip text area, fixed before the table birth.
          const recovered = readOoxmlPart(
            '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">Peer</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
            { name: '/recovered-table-oracle.xml', contentType: originalPart.contentType }
          );
          if (!recovered.ok) throw Error(recovered.reason);
          const row =
            recovered.part.root.kind === 'table' &&
            recovered.part.root.children.find((n) => n.kind === 'tableRow');
          const cell =
            row && row.kind === 'tableRow' && row.children.find((n) => n.kind === 'tableCell');
          const paragraph =
            cell && cell.kind === 'tableCell' && cell.children.find((n) => n.kind === 'paragraph');
          if (!paragraph || paragraph.kind !== 'paragraph')
            throw Error('Missing recovered paragraph oracle.');
          const bound = replaceNode(recovered.part, paragraph.id, {
            ...paragraph,
            attributes: identity,
          });
          if (!bound.ok)
            throw Error('Cannot bind retained peer identity: ' + JSON.stringify(bound.issues));
          const retained = insertChildren(originalPart, originalBody.id, 1, [bound.part.root]);
          if (!retained.ok) throw Error('Cannot construct retained table oracle.');
          const recovery = withPart(before, retained.part);
          const geometry = async () => {
            const text = paragraphTextOf(bob!.surface!.session.part(), paragraphId)!;
            expect(text).toContain('Peer');
            const offset = text.indexOf('Peer') + 4;
            bob!.surface!.setSelection({
              anchor: { paragraphId, offset },
              head: { paragraphId, offset },
            });
            await pair.settle();
            const remote = pair.rooms[0]!.session.remoteSelections().find(
              (selection) => selection.actorId === 'bob'
            );
            expect(remote?.head.nodeId).toBe(paragraphId);
            expect(remote?.head.offset).toBe(offset);
            const points = pair.editors.map((editor) =>
              caretAt(editor.surface!.layout(), { paragraphId, offset })
            );
            points.forEach((point) => expect(point).not.toBeNull());
            expect({ x: points[0]!.x, y: points[0]!.y }).toEqual({
              x: points[1]!.x,
              y: points[1]!.y,
            });
          };
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
          check(firstPendingWord, 4);
          await geometry();
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          check(word(recovery), 0);
          await geometry();
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(firstPendingWord, 4);
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          check(pendingWord, 4);
          await geometry();
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          expect(
            paragraphTextOf(alice!.surface!.session.part(), paragraphId)?.includes('Peer')
          ).toBe(false);
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
          const inserted = insertChildren(sourcePart, body.id, 1, [
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
          await geometry();
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
