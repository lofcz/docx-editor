// Test-only transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { propertyContext, propertyToken } from '../../test/support/property-execution';
import {
  preparedApplicationDocument,
  applicationFragmentOptions as options,
  applicationFragmentTarget,
} from '../../test/support/application-fragments';
import { buildBlockIndex } from '../blocks/block-index';
import { parseTargetMarkdown } from '../projection/parse-target';
import { targetAstKey } from '../diff/target-key';
import { recomposeOwnedProposals } from './recompose';
import { readWorkingView } from './working-view';
import { readCommittedView } from './committed-view';
import { readReviewGroups } from './groups';
import { rejectAllReviewGroups } from './reject-all';
import { rejectReviewGroup } from './reject-group';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
const target = () => applicationFragmentTarget('First') + '\n\nNeighbor';
const key = (text: string) => {
  const read = parseTargetMarkdown(text);
  if (!read.success) throw Error(read.error);
  return targetAstKey(read.target.normalizedAst);
};
const word = (pkg: OoxmlPackage) => {
  const read = readOoxmlPackage(writeOoxmlPackage(pkg));
  if (!read.ok) throw Error(read.reason);
  return canonicalOoxmlFingerprint(read.package.parts.get(pkg.mainDocumentPart)!);
};
for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const overlap of [true, false])
      test(`unselected presence stays actionable through ${overlap ? 'refused overlap' : 'outside proposal'}, peer edits, histories, decisions, and reopening; ${clients}/${receiver}`, async () => {
        const fixture = {
          document: preparedApplicationDocument(
            docxFixture(
              '<w:p><w:r><w:t>Outside</w:t></w:r></w:p><w:p><w:r><w:t>Neighbor</w:t></w:r></w:p>'
            )
          ),
          documentId: propertyContext.documentId,
          ownedAttributes: true,
          editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
        };
        const pair = await createTwoEditorFixture({ ...fixture, clientIds: clients });
        try {
          const [alice, bob] = pair.editors;
          const baseline = alice!.surface!.session.currentPackage();
          const neighbor = alice!
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(alice!.surface!.session.part(), id) === 'Neighbor')!;
          const first = await recomposeOwnedProposals(
            alice!,
            { revision: propertyToken(alice!), targetMarkdown: target(), ownership: [] },
            propertyContext,
            { target: options('first'), payloads: { presence: { proposalId: 'first-birth' } } }
          );
          if (!first.success) throw Error(first.error);
          await pair.settle();
          const pending = alice!.surface!.session.currentPackage();
          const beforeGroups = readReviewGroups(pending);
          if (!beforeGroups.success) throw Error(beforeGroups.error);
          expect(beforeGroups.groups).toHaveLength(3);
          const owners = new Map(
            [...buildBlockIndex(pending).byId].map(([id, block]) => [
              id,
              canonicalOoxmlFingerprint(
                findNode(pending.parts.get(block.partName)!, block.controlId)!
              ),
            ])
          );
          pair.pause();
          expect(
            bob!.surface!.applyAutomationOps(() => [
              { op: 'insertText', paragraphId: neighbor, offset: 8, text: ' Peer' },
            ]).rejected
          ).toBe(false);
          const before = alice!.surface!.session.currentPackage(),
            bytes = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
          let publications = 0;
          const stop = alice!.surface!.session.subscribe((change) => {
            if (!change.origin.endsWith('.remote')) publications++;
          });
          const result = await recomposeOwnedProposals(
            alice!,
            {
              revision: propertyToken(alice!),
              targetMarkdown: overlap
                ? target().replace(
                    '</phase>',
                    '<activity name="Unowned" duration="2">\nNew body\n</activity>\n</phase>'
                  )
                : target().replace('Outside', 'Outside updated'),
              ownership: [],
            },
            propertyContext,
            {
              target: options('unowned'),
              ...(overlap ? { payloads: { presence: { proposalId: 'unowned-birth' } } } : {}),
            }
          );
          stop();
          if (overlap) {
            expect(result.success).toBe(false);
            if (!result.success)
              expect(result.error).toBe('Replacement would change an unselected review group.');
            expect(publications).toBe(0);
            expect(alice!.surface!.session.currentPackage()).toBe(before);
            pair.documents.forEach((doc, index) =>
              expect(Y.encodeStateAsUpdate(doc)).toEqual(bytes[index]!)
            );
          } else {
            if (!result.success) throw Error(result.error);
            expect(publications).toBe(1);
            expect(result.ownership).toHaveLength(1);
          }
          pair.reconnect(receiver);
          await pair.settle();
          const peer = (pkg: OoxmlPackage) => {
            const edit = applyTreeOp(pkg.parts.get(pkg.mainDocumentPart)!, {
              op: 'insertText',
              paragraphId: neighbor,
              offset: 8,
              text: ' Peer',
            });
            if (!edit.ok) throw Error(edit.reason);
            return withPart(pkg, edit.part);
          };
          const expectedPending = word(peer(pending));
          const currentTarget = target()
            .replace('Neighbor', 'Neighbor Peer')
            .replace('Outside', overlap ? 'Outside' : 'Outside updated');
          const check = async (peerText = true) => {
            for (const editor of pair.editors) {
              const pkg = editor.surface!.session.currentPackage(),
                groups = readReviewGroups(pkg);
              if (!groups.success) throw Error(groups.error);
              for (const group of beforeGroups.groups) {
                const live = groups.groups.find((candidate) => candidate.id === group.id)!;
                expect(live.readOnly).toBe(false);
                expect(live.properties).toEqual(group.properties);
              }
              for (const [id, fingerprint] of owners) {
                const owner = buildBlockIndex(pkg).byId.get(id)!;
                expect(
                  canonicalOoxmlFingerprint(
                    findNode(pkg.parts.get(owner.partName)!, owner.controlId)!
                  )
                ).toBe(fingerprint);
              }
              const original = await readCommittedView(
                editor.surface!.session,
                propertyToken(editor)
              );
              const working = await readWorkingView(editor.surface!.session, propertyToken(editor));
              if (!original.success || !working.success) throw Error('Missing review views');
              expect(key(original.view.projection.text)).toBe(
                key(peerText ? 'Outside\n\nNeighbor Peer' : 'Outside\n\nNeighbor')
              );
              expect(key(working.view.projection.text)).toBe(
                key(peerText ? currentTarget : currentTarget.replace('Neighbor Peer', 'Neighbor'))
              );
              if (overlap && peerText) expect(word(pkg)).toBe(expectedPending);
            }
          };
          await check();
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(false);
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check();
          expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            const read = await readWorkingView(editor.surface!.session, propertyToken(editor));
            if (!read.success) throw Error(read.error);
            expect(key(read.view.projection.text)).toBe(
              key(
                overlap ? 'Outside\n\nNeighbor Peer' : target().replace('Neighbor', 'Neighbor Peer')
              )
            );
          }
          expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check();
          const born = bob!
            .surface!.session.paragraphIds()
            .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'A rich link.')!;
          bob!.surface!.setSelection({
            anchor: { paragraphId: born, offset: 2 },
            head: { paragraphId: born, offset: 2 },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          );
          expect(remote?.head.nodeId).toBe(born);
          expect(remote?.head.offset).toBe(2);
          const a = caretAt(alice!.surface!.layout(), { paragraphId: born, offset: 2 }),
            b = caretAt(bob!.surface!.layout(), { paragraphId: born, offset: 2 });
          expect(a).not.toBeNull();
          expect(b).not.toBeNull();
          expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
          if (overlap) {
            const decision = rejectAllReviewGroups(bob!, propertyToken(bob!), propertyContext);
            if (!decision.success) throw Error(decision.error);
          } else {
            if (!result.success) throw Error(result.error);
            for (const group of result.ownership) {
              const decision = rejectReviewGroup(
                bob!,
                group.groupId,
                propertyToken(bob!),
                propertyContext
              );
              if (!decision.success) throw Error(decision.error);
            }
          }
          await pair.settle();
          const expected = overlap ? word(peer(baseline)) : expectedPending;
          for (const editor of pair.editors)
            expect(word(editor.surface!.session.currentPackage())).toBe(expected);
          expect(bob!.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check();
          expect(bob!.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          const saved = await pair.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
          for (const shared of [false, true]) {
            const reopened = await createTwoEditorFixture({
              ...fixture,
              document: new Uint8Array(await alice!.save()),
              ...(shared ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
            });
            try {
              await reopened.settle();
              for (const editor of reopened.editors) {
                expect(word(editor.surface!.session.currentPackage())).toBe(expected);
                const groups = readReviewGroups(editor.surface!.session.currentPackage());
                expect(groups.success && groups.groups.every((group) => !group.readOnly)).toBe(
                  true
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
      }, 30000);

test('an unprepared shared application list refuses before first publication and history', async () => {
  const pair = await createTwoEditorFixture({
    document: docxFixture('<w:p><w:r><w:t>Outside</w:t></w:r></w:p>'),
    documentId: propertyContext.documentId,
    ownedAttributes: true,
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
  });
  try {
    const [alice] = pair.editors,
      before = alice!.surface!.session.currentPackage(),
      bytes = pair.documents.map((doc) => Y.encodeStateAsUpdate(doc));
    let publications = 0;
    const stop = alice!.surface!.session.subscribe(() => publications++);
    const result = await recomposeOwnedProposals(
      alice!,
      {
        revision: propertyToken(alice!),
        targetMarkdown: applicationFragmentTarget('First'),
        ownership: [],
      },
      propertyContext,
      { target: options('unprepared'), payloads: { presence: { proposalId: 'unprepared-birth' } } }
    );
    stop();
    expect(result.success).toBe(false);
    if (!result.success)
      expect(result.error).toBe(
        'Prepare native numbering infrastructure before sharing application list fragments.'
      );
    expect(publications).toBe(0);
    expect(alice!.surface!.session.currentPackage()).toBe(before);
    expect(alice!.surface!.session.canUndo()).toBe(false);
    expect(alice!.snapshot().canUndo).toBe(false);
    await pair.settle();
    pair.documents.forEach((doc, index) =>
      expect(Y.encodeStateAsUpdate(doc)).toEqual(bytes[index]!)
    );
    for (const editor of pair.editors) expect(editor.snapshot().collaborationStatus).toBe('ready');
    pair.ownership.forEach((handle) => handle.assertHealthy());
  } finally {
    pair.destroy();
  }
});
