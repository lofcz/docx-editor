import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlElement,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { groupToken, groups } from '../../test/support/review-groups';
import { propertyContext } from '../../test/support/property-execution';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { withInitializedReviewGroupStore, withReviewGroupRecord } from './group-record-store';
import { rejectReviewGroup } from './reject-group';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const groupId = 'g1:stored:rg_leading_control';
const documentId = propertyContext.documentId;
function documentBytes(peerProperty = false) {
  const revision = `w:id="17" w:author="ScioBot" w:date="2026-10-19T09:00:00Z"`;
  const alignment = peerProperty ? '<w:jc w:val="right"/>' : '';
  const change = peerProperty
    ? '<w:pPrChange w:id="22" w:author="Peer"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>'
    : '';
  const read = readOoxmlPackage(
    docxFixture(
      `<w:p><w:pPr>${alignment}<w:rPr><w:ins ${revision}/></w:rPr>${change}</w:pPr><w:ins ${revision}><w:r><w:t>New</w:t></w:r></w:ins></w:p>` +
        '<w:sdt><w:sdtPr><w:tag w:val="foreign-control"/></w:sdtPr><w:sdtContent>' +
        p('Control body') +
        '</w:sdtContent></w:sdt>' +
        p('Tail')
    )
  );
  if (!read.ok) throw new Error(read.reason);
  const initialized = withInitializedReviewGroupStore(read.package);
  if (!initialized.ok) throw new Error(initialized.reason);
  const written = withReviewGroupRecord(initialized.pkg, {
    schemaVersion: 1,
    id: 'rg_leading_control',
    members: [
      {
        partName: '/word/document.xml',
        revision: { id: '17', author: 'ScioBot', date: '2026-10-19T09:00:00Z' },
      },
    ],
  });
  if (!written.ok) throw new Error(written.reason);
  return writeOoxmlPackage(written.pkg);
}
const options = (peerProperty = false) => ({
  document: documentBytes(peerProperty),
  documentId,
  ownedAttributes: true,
  editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
});
const pkg = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
function word(source: OoxmlPackage) {
  const read = readOoxmlPackage(writeOoxmlPackage(source));
  if (!read.ok) throw new Error(read.reason);
  return canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!);
}
/** Independent expected Word structure: remove the complete inserted first paragraph. */
function expectedWithoutLeading(source: OoxmlPackage, id: string) {
  const part = source.parts.get(source.mainDocumentPart)!;
  const body = part.root.children.find((node) => node.kind === 'body') as OoxmlElement;
  return word(
    withPart(source, {
      ...part,
      root: {
        ...part.root,
        children: part.root.children.map((node) =>
          node === body
            ? { ...body, children: body.children.filter((child) => child.id !== id) }
            : node
        ),
      } as OoxmlElement,
    })
  );
}
function reject(editor: DocxEditorInstance) {
  const result = rejectReviewGroup(editor, groupId, groupToken(editor), propertyContext);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
}
function append(editor: DocxEditorInstance, id: string, offset: number, text: string) {
  const result = editor.surface!.applyAutomationOps(() => [
    { op: 'insertText', paragraphId: id, offset, text, bias: 'right' },
  ]);
  if (!result.committed || result.rejected) throw new Error(String(result.reason));
}
async function views(editor: DocxEditorInstance) {
  const [before, after] = await Promise.all([
    readCommittedView(editor.surface!.session, groupToken(editor)),
    readWorkingView(editor.surface!.session, groupToken(editor)),
  ]);
  if (!before.success) throw new Error(before.error);
  if (!after.success) throw new Error(after.error);
  return [before.view.projection.text, after.view.projection.text];
}
async function geometry(
  pair: Awaited<ReturnType<typeof createTwoEditorFixture>>,
  id: string,
  offset: number
) {
  const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
  bob.surface!.setSelection({
    anchor: { paragraphId: id, offset },
    head: { paragraphId: id, offset },
  });
  await pair.settle();
  const remote = pair.rooms[0]!.session.remoteSelections().find((row) => row.actorId === 'bob');
  expect(remote?.head.nodeId).toBe(id);
  expect(remote?.head.offset).toBe(offset);
  const left = caretAt(alice.surface!.layout(), { paragraphId: id, offset }),
    right = caretAt(bob.surface!.layout(), { paragraphId: id, offset });
  expect(left).not.toBeNull();
  expect(right).not.toBeNull();
  expect({ x: left!.x, y: left!.y }).toEqual({ x: right!.x, y: right!.y });
  expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
}
async function saved(
  pair: Awaited<ReturnType<typeof createTwoEditorFixture>>,
  expectedGroupIds: readonly string[] = []
) {
  const expected = await views(pair.editors[0]!),
    output = word(pkg(pair.editors[0]!));
  expect(word(pkg(pair.editors[1]!))).toBe(output);
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  for (const room of [false, true]) {
    const reopened = await createTwoEditorFixture({
      ...options(),
      document: new Uint8Array(await pair.editors[0]!.save()),
      clientIds: room ? [401, 433] : [181, 211],
      ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
    });
    try {
      for (const editor of reopened.editors) {
        expect(await views(editor)).toEqual(expected);
        expect(word(pkg(editor))).toBe(output);
        if (expectedGroupIds.length)
          expect(groups(editor).map((group) => group.id)).toEqual([...expectedGroupIds]);
        else expect(groups(editor)).toEqual([]);
      }
    } finally {
      reopened.destroy();
    }
  }
}

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const) {
    for (const firstUndo of [0, 1] as const)
      test(`concurrent leading-control paragraph rejection retains one decision through both actors' repeated history; ${clientIds}/${receiver}/${firstUndo}`, async () => {
        const pair = await createTwoEditorFixture({ ...options(), clientIds });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
            id = alice.surface!.session.paragraphIds()[0]!,
            tail = alice.surface!.session.paragraphIds().at(-1)!;
          append(bob, tail, 4, ' Peer');
          await pair.settle();
          const pending = word(pkg(alice)),
            expected = expectedWithoutLeading(pkg(alice), id);
          pair.pause();
          reject(alice);
          reject(bob);
          pair.reconnect(receiver);
          await pair.settle();
          const assert = async (isPending: boolean) => {
            for (const editor of pair.editors) {
              expect(word(pkg(editor))).toBe(isPending ? pending : expected);
              expect(groups(editor).map((group) => group.id)).toEqual(isPending ? [groupId] : []);
              expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe('Tail Peer');
              expect(findNode(editor.surface!.session.part(), id)).not.toBeNull();
            }
          };
          await assert(false);
          for (let cycle = 0; cycle < 2; cycle++) {
            expect(pair.editors[firstUndo]!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await assert(false);
            expect(pair.editors[1 - firstUndo]!.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await assert(true);
            expect(pair.editors[firstUndo]!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await assert(false);
            expect(pair.editors[1 - firstUndo]!.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await assert(false);
          }
          expect(collectRevisionSites(alice.surface!.session.part())).toEqual([]);
          await geometry(pair, tail, 9);
          await saved(pair);
        } finally {
          pair.destroy();
        }
      }, 30000);

    test(`leading-control rejection preserves unseen peer text, peer history, geometry and reopened rooms; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          id = alice.surface!.session.paragraphIds()[0]!;
        const expected = expectedWithoutLeading(pkg(alice), id);
        pair.pause();
        reject(alice);
        expect(word(pkg(alice))).toBe(expected);
        expect(alice.surface!.session.paragraphIds()).not.toContain(id);
        append(bob, id, 3, 'Peer');
        pair.reconnect(receiver);
        await pair.settle();
        const baseline = 'Peer\n\nControl body\n\nTail';
        for (const editor of pair.editors) {
          expect(await views(editor)).toEqual([baseline, baseline]);
          expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe('Peer');
        }
        await geometry(pair, id, 4);
        for (let cycle = 0; cycle < 2; cycle++) {
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await views(editor)).toEqual([baseline, 'NewPeer\n\nControl body\n\nTail']);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await views(editor)).toEqual([baseline, baseline]);
        }
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) expect(word(pkg(editor))).toBe(expected);
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([baseline, baseline]);
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30000);

    test(`explicit leading-control owner deletion preserves peer text through undo and redo; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          id = alice.surface!.session.paragraphIds()[0]!;
        reject(alice);
        await pair.settle();
        const expected = word(pkg(alice));
        pair.pause();
        const result = alice.surface!.applyAutomationOps(() => [
          { op: 'deleteBlock', blockId: id },
        ]);
        expect(result.rejected).toBe(false);
        append(bob, id, 0, 'Peer');
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(findNode(editor.surface!.session.part(), id)).toBeNull();
          expect(word(pkg(editor))).toBe(expected);
        }
        for (let cycle = 0; cycle < 2; cycle++) {
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors)
            expect(await views(editor)).toEqual([
              'Peer\n\nControl body\n\nTail',
              'Peer\n\nControl body\n\nTail',
            ]);
          await geometry(pair, id, 4);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) expect(word(pkg(editor))).toBe(expected);
        }
        await saved(pair);
      } finally {
        pair.destroy();
      }
    }, 30000);
  }

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`owned leading paragraph rejection preserves an independent pending paragraph format through history and reopening; ${clientIds}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ ...options(true), clientIds });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          leading = alice.surface!.session.paragraphIds()[0]!,
          control = alice.surface!.session.paragraphIds()[1]!,
          tail = alice.surface!.session.paragraphIds().at(-1)!,
          initialGroups = groups(alice),
          foreign = initialGroups.find((group) =>
            group.native.some((member) => member.address.author === 'Peer')
          )!;
        expect(initialGroups).toHaveLength(2);
        const peerFingerprint = (editor: DocxEditorInstance) => {
          const part = editor.surface!.session.part(),
            peer = collectRevisionSites(part).find(
              (site) =>
                site.node.localName === 'pPrChange' &&
                site.node.attributes.some(
                  (attribute) => attribute.localName === 'author' && attribute.value === 'Peer'
                )
            );
          if (!peer) throw new Error('The independent peer property revision was lost.');
          return canonicalOoxmlFingerprint({ ...part, root: peer.node });
        };
        const expectedPeer = peerFingerprint(alice);
        pair.pause();
        reject(alice);
        append(bob, tail, 4, ' Peer');
        pair.reconnect(receiver);
        await pair.settle();
        const assert = (pending: boolean) => {
          for (const editor of pair.editors) {
            const part = editor.surface!.session.part();
            expect(peerFingerprint(editor)).toBe(expectedPeer);
            expect(paragraphTextOf(part, leading)).toBe(pending ? 'New' : '');
            expect(paragraphTextOf(part, control)).toBe('Control body');
            expect(paragraphTextOf(part, tail)).toBe('Tail Peer');
            expect(editor.surface!.session.paragraphIds()).toContain(leading);
            expect(groups(editor).map((group) => group.id)).toEqual(
              pending ? initialGroups.map((group) => group.id) : [foreign.id]
            );
            expect(collectRevisionSites(part)).toHaveLength(pending ? 3 : 1);
            const paragraph = findNode(part, leading) as OoxmlElement,
              properties = paragraph.children.find(
                (node) => node.kind !== 'textValue' && node.localName === 'pPr'
              ) as OoxmlElement,
              alignment = properties.children.find(
                (node) => node.kind !== 'textValue' && node.localName === 'jc'
              ) as OoxmlElement;
            expect(
              alignment.attributes.find((attribute) => attribute.localName === 'val')?.value
            ).toBe('right');
          }
        };
        assert(false);
        await geometry(pair, leading, 0);
        const resolvedWord = word(pkg(alice));
        for (let cycle = 0; cycle < 2; cycle++) {
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          assert(true);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          assert(false);
          for (const editor of pair.editors) expect(word(pkg(editor))).toBe(resolvedWord);
        }
        expect(bob.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe('Tail');
          expect(peerFingerprint(editor)).toBe(expectedPeer);
          expect(groups(editor).map((group) => group.id)).toEqual([foreign.id]);
        }
        expect(bob.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        assert(false);
        await saved(pair, [foreign.id]);
      } finally {
        pair.destroy();
      }
    }, 30000);
