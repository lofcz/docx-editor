// Test-only transport uses the EigenPal Pro License. Production SDK imports stay unchanged.
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlNode,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { propertyEditor } from '../../test/support/property-execution';
import { LINK_CONTEXT, linkViews } from '../../test/support/link-review';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
const revision = { author: 'Bot', date: '2026-10-21T21:00:00Z' };
type Kind = 'insert' | 'remove';
type Pair = Awaited<ReturnType<typeof createTwoEditorFixture>>;
const run = (text: string, mark = '') =>
  `<w:r><w:rPr>${mark}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const document_ = (kind: Kind, proofing = false) =>
  docxFixture(
    '<w:p><w:pPr/>' +
      run('See ') +
      (kind === 'remove' ? '<w:hyperlink r:id="shared" w:tooltip="Old">' : '') +
      run('the ', '<w:b/>') +
      (proofing ? '<w:proofErr w:type="spellStart"/>' : '') +
      run('site', '<w:i/>') +
      (kind === 'remove' ? '</w:hyperlink>' : '') +
      run(' or ') +
      '<w:hyperlink r:id="shared">' +
      run('twin') +
      '</w:hyperlink></w:p><w:p><w:pPr/>' +
      run('Tail') +
      '</w:p>',
    {
      relationships:
        '<Relationship Id="shared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/shared" TargetMode="External"/>',
    }
  );
const kids = (node: OoxmlNode) => (node.kind === 'textValue' ? [] : node.children);
function anchors(editor: DocxEditorInstance) {
  const part = editor.surface!.session.part(),
    body = bodyStoryRoot(part)!;
  const p = kids(body).find((node) => node.kind === 'paragraph')!;
  return {
    paragraphId: p.id,
    linkId: kids(p).find((node) => node.kind === 'hyperlink')?.id,
    tail: editor
      .surface!.session.paragraphIds()
      .find((id) => paragraphTextOf(part, id) === 'Tail')!,
  };
}
function mutate(editor: DocxEditorInstance, kind: Kind, tracked = true, existingTarget?: string) {
  const { paragraphId, linkId } = anchors(editor);
  return editor.surface!.applyAutomationOps(
    (relate) => {
      let op: TreeDocOp;
      if (kind === 'insert') {
        const relationshipId = existingTarget ?? relate('https://example.org/new');
        if (!relationshipId) return null;
        op = {
          op: 'insertHyperlink',
          paragraphId,
          start: 4,
          end: 12,
          relationshipId,
          tooltip: 'New',
        };
      } else op = { op: 'removeHyperlink', linkId: linkId! };
      return [{ ...op, ...(tracked ? { revision } : {}) } as TreeDocOp];
    },
    { kind: 'body' },
    [],
    tracked,
    true
  );
}
async function golden(kind: Kind) {
  const opened = readOoxmlPackage(document_(kind));
  if (!opened.ok) throw new Error(opened.reason);
  const editor = propertyEditor(opened.package);
  try {
    expect(mutate(editor, kind, false)).toMatchObject({ committed: true, rejected: false });
    return (await linkViews(editor))[1]!;
  } finally {
    editor.destroy();
  }
}
const options = (kind: Kind) => ({
  document: document_(kind),
  documentId: LINK_CONTEXT.documentId,
  ownedAttributes: true,
  editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
});

for (const kind of ['insert', 'remove'] as const)
  for (const clientIds of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      for (const action of ['accept', 'reject'] as const)
        test(`native ${kind} link retains exact spelling markers ${clientIds}/${receiver}/${action}`, async () => {
          const pair = await createTwoEditorFixture({
            ...options(kind),
            document: document_(kind, true),
            clientIds,
          });
          let direct: DocxEditorInstance | undefined;
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
            const opened = readOoxmlPackage(new Uint8Array(await alice.save()));
            if (!opened.ok) throw Error(opened.reason);
            direct = propertyEditor(opened.package);
            const original = canonicalOoxmlFingerprint(alice.surface!.session.part());
            // Reuse one real relationship: joined rooms reserve different allocation ranges.
            expect(mutate(direct, kind, false, 'shared')).toMatchObject({
              committed: true,
              rejected: false,
            });
            const wanted =
              action === 'accept'
                ? canonicalOoxmlFingerprint(direct.surface!.session.part())
                : original;
            pair.pause();
            expect(mutate(alice, kind, true, 'shared')).toMatchObject({
              committed: true,
              rejected: false,
            });
            const pending = canonicalOoxmlFingerprint(alice.surface!.session.part());
            pair.reconnect(receiver);
            await pair.settle();
            for (const editor of pair.editors)
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
            expect(bob.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors) {
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(wanted);
              const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
              if (!saved.ok) throw Error(saved.reason);
              expect(
                canonicalOoxmlFingerprint(saved.package.parts.get('/word/document.xml')!)
              ).toBe(wanted);
            }
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(pending);
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(wanted);
            pair.ownership.forEach((handle) => handle.assertHealthy());
          } finally {
            direct?.destroy();
            pair.destroy();
          }
        }, 30000);

async function check(pair: Pair, original: string, proposed: string) {
  for (const editor of pair.editors) expect(await linkViews(editor)).toEqual([original, proposed]);
}
async function geometry(pair: Pair, id: string, requestedOffset?: number) {
  const [alice, bob] = pair.editors;
  const offset = requestedOffset ?? paragraphTextOf(bob!.surface!.session.part(), id)!.length;
  bob!.surface!.setSelection({
    anchor: { paragraphId: id, offset },
    head: { paragraphId: id, offset },
  });
  await pair.settle();
  const remote = pair.rooms[0]!.session.remoteSelections().find(
    (selection) => selection.actorId === 'bob'
  );
  expect(remote?.head.nodeId).toBe(id);
  expect(remote?.head.offset).toBe(offset);
  const a = caretAt(alice!.surface!.layout(), { paragraphId: id, offset }),
    b = caretAt(bob!.surface!.layout(), { paragraphId: id, offset });
  expect(a).not.toBeNull();
  expect(b).not.toBeNull();
  expect({ x: a!.x, y: a!.y }).toEqual({ x: b!.x, y: b!.y });
  expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(0);
}
async function saved(pair: Pair, kind: Kind) {
  const expected = await linkViews(pair.editors[0]!);
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  for (const room of [false, true]) {
    const again = await createTwoEditorFixture({
      ...options(kind),
      document: new Uint8Array(await pair.editors[0]!.save()),
      clientIds: room ? [401, 433] : [181, 211],
      ...(room ? { roomSnapshot: Y.encodeStateAsUpdate(pair.documents[0]!) } : {}),
    });
    try {
      for (const editor of again.editors) expect(await linkViews(editor)).toEqual(expected);
    } finally {
      again.destroy();
    }
  }
}
for (const kind of ['insert', 'remove'] as const)
  for (const clientIds of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      for (const action of ['accept', 'reject'] as const)
        test(`native tracked ${kind} link preserves both histories, ${action}, reconnect ${clientIds}/${receiver}`, async () => {
          const pair = await createTwoEditorFixture({ ...options(kind), clientIds });
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
              original = (await linkViews(alice))[0]!,
              expected = await golden(kind),
              { tail } = anchors(alice);
            pair.pause();
            const publications: { origin: string; main: string }[] = [];
            const stop = alice.surface!.session.subscribe((change) =>
              publications.push({
                origin: change.origin,
                main: canonicalOoxmlFingerprint(alice.surface!.session.part()),
              })
            );
            expect(mutate(alice, kind)).toMatchObject({ committed: true, rejected: false });
            stop();
            expect(publications.filter((entry) => !entry.origin.endsWith('.remote'))).toHaveLength(
              1
            );
            expect(new Set(publications.map((entry) => entry.main)).size).toBe(1);
            expect(
              bob.surface!.applyAutomationOps(
                () => [{ op: 'insertText', paragraphId: tail, offset: 4, text: ' Peer' }],
                { kind: 'body' },
                [],
                false,
                true
              )
            ).toMatchObject({ committed: true, rejected: false });
            pair.reconnect(receiver);
            await pair.settle();
            const before = original.replace('Tail', 'Tail Peer'),
              after = expected.replace('Tail', 'Tail Peer');
            await check(pair, before, after);
            await geometry(pair, tail);
            await geometry(pair, anchors(alice).paragraphId, 16);
            await saved(pair, kind);
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, before, before);
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, before, after);
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, original, expected);
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, before, after);
            expect(bob.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
            await pair.settle();
            const decided = action === 'accept' ? after : before;
            await check(pair, decided, decided);
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, before, after);
            expect(bob.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            await check(pair, decided, decided);
            await saved(pair, kind);
            pair.ownership.forEach((handle) => handle.assertHealthy());
          } finally {
            pair.destroy();
          }
        }, 30000);

for (const kind of ['insert', 'remove'] as const)
  for (const clientIds of [
    [41, 73],
    [73, 41],
  ] as const)
    for (const receiver of [0, 1] as const)
      test(`native tracked ${kind} link cannot recreate a peer-deleted paragraph through either history; ${clientIds}/${receiver}`, async () => {
        const pair = await createTwoEditorFixture({ ...options(kind), clientIds });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
            original = (await linkViews(alice))[0]!,
            expected = await golden(kind),
            { paragraphId, tail } = anchors(alice);
          pair.pause();
          expect(mutate(alice, kind)).toMatchObject({ committed: true, rejected: false });
          expect(
            bob.surface!.applyAutomationOps(
              () => [{ op: 'insertText', paragraphId: tail, offset: 4, text: ' Peer' }],
              { kind: 'body' },
              [],
              false,
              true
            )
          ).toMatchObject({ committed: true, rejected: false });
          expect(
            bob.surface!.applyAutomationOps(
              () => [{ op: 'deleteBlock', blockId: paragraphId }],
              { kind: 'body' },
              [],
              false,
              true
            )
          ).toMatchObject({ committed: true, rejected: false });
          pair.reconnect(receiver);
          await pair.settle();
          await check(pair, 'Tail Peer', 'Tail Peer');
          await geometry(pair, tail);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, 'Tail Peer', 'Tail Peer');
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, 'Tail Peer', 'Tail Peer');
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check(
            pair,
            original.replace('Tail', 'Tail Peer'),
            expected.replace('Tail', 'Tail Peer')
          );
          await geometry(pair, paragraphId, 16);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check(pair, 'Tail Peer', 'Tail Peer');
          await saved(pair, kind);
        } finally {
          pair.destroy();
        }
      }, 30000);
