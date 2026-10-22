import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt, paragraphFragmentsOf } from '@docx-editor.dev/core/layout';
import { strFromU8, unzipSync } from 'fflate';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  type OoxmlElement,
  type OoxmlParagraphNode,
  type RevisionAddress,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { bodyPart, docxFixture } from '../../test/support/document';
import { readEditorRevision } from '../projection/revision-cache';
import { readCommittedView } from './committed-view';
import { readWorkingView } from './working-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const date = '2026-10-19T09:00:00Z';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const original = 'Opening\\.\n\nClosing\\.';
const seeds = (xml: string) =>
  (bodyPart(xml).root.children[0] as OoxmlElement).children as readonly OoxmlParagraphNode[];
const fixture = (clientIds: readonly [number, number]) =>
  createTwoEditorFixture({
    document: docxFixture(p('Opening.') + p('Closing.')),
    clientIds,
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
  });
function apply(editor: DocxEditorInstance, ops: readonly TreeDocOp[], review = false) {
  const result = editor.surface!.applyAutomationOps(
    () => ops,
    { kind: 'body' },
    undefined,
    review,
    true
  );
  if (result.rejected || !result.committed) throw new Error(JSON.stringify(result));
}
function insert(editor: DocxEditorInstance, xml: string, beforeId: string | null, author: string) {
  const parent = editor.surface!.session.part().root.children[0]!;
  apply(
    editor,
    [
      {
        op: 'insertParagraphs',
        parentId: parent.id,
        beforeId,
        paragraphs: seeds(xml),
        revision: { author, date },
      },
    ],
    true
  );
}
function address(editor: DocxEditorInstance, author: string): RevisionAddress {
  const node = collectRevisionSites(editor.surface!.session.part()).find((site) =>
    site.node.attributes.some(
      (attribute) => attribute.localName === 'author' && attribute.value === author
    )
  )?.node;
  if (!node) throw new Error('Missing native revision');
  const attr = (name: string) =>
    node.attributes.find((attribute) => attribute.localName === name)?.value;
  return { id: attr('id')!, author: attr('author')!, date: attr('date')! };
}
async function views(editor: DocxEditorInstance) {
  const token = readEditorRevision(editor, 'siblings', 'epoch')!;
  const [committed, working] = await Promise.all([
    readCommittedView(editor.surface!.session, token),
    readWorkingView(editor.surface!.session, token),
  ]);
  if (!committed.success) throw new Error(committed.error);
  if (!working.success) throw new Error(working.error);
  return [committed.view.projection.text, working.view.projection.text];
}
async function saved(pair: Awaited<ReturnType<typeof fixture>>) {
  const digests = await pair.savedDigests();
  expect(digests[0]).toEqual(digests[1]);
  const container = document.createElement('div');
  document.body.append(container);
  const reopened = createDocxEditor({
    container,
    document: new Uint8Array(await pair.editors[0]!.save()),
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
  try {
    expect(await views(reopened)).toEqual(await views(pair.editors[0]!));
  } finally {
    reopened.destroy();
    container.remove();
  }
}

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`rejecting inserted paragraphs preserves concurrent peer content; clients ${clients.join('/')}, receiver ${receiver}`, async () => {
      const pair = await fixture(clients);
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        insert(alice, p('New'), null, 'Alice AI');
        await pair.settle();
        const id = alice.surface!.session.paragraphIds().at(-1)!;
        const revision = address(alice, 'Alice AI');
        alice.surface!.setSelection({
          anchor: { paragraphId: id, offset: 3 },
          head: { paragraphId: id, offset: 3 },
        });
        bob.surface!.setSelection({
          anchor: { paragraphId: id, offset: 3 },
          head: { paragraphId: id, offset: 3 },
        });
        await pair.settle();
        pair.pause();
        apply(alice, [{ op: 'rejectRevision', revision }], true);
        const localSelection = alice.surface!.state().selection;
        expect(localSelection).not.toBeNull();
        expect(localSelection!.head.paragraphId).not.toBe(id);
        expect(caretAt(alice.surface!.layout(), localSelection!.head)).not.toBeNull();
        expect(alice.surface!.session.paragraphIds()).not.toContain(id);
        expect(findNode(alice.surface!.session.part(), id)).not.toBeNull();
        expect(
          alice
            .surface!.layout()
            .pages.flatMap((page) => paragraphFragmentsOf(page))
            .some((paragraph) => paragraph.paragraphId === id)
        ).toBe(false);
        const beforeSave = canonicalOoxmlFingerprint(alice.surface!.session.part());
        const savedBytes = new Uint8Array(await alice.save());
        const xml = strFromU8(unzipSync(savedBytes)['word/document.xml']!);
        expect(xml).not.toContain('urn:docx-editor.dev:review:retired-paragraph:1');
        expect(xml).not.toContain('>New<');
        expect(canonicalOoxmlFingerprint(alice.surface!.session.part())).toBe(beforeSave);
        const staleRevision = alice.surface!.session.packageRevision();
        const stale = alice.surface!.applyAutomationOps(
          () => [{ op: 'rejectRevision', revision }],
          { kind: 'body' },
          undefined,
          true,
          true
        );
        expect(stale.rejected).toBe(true);
        expect(stale.reason).toBe('unknown-revision');
        expect(alice.surface!.session.packageRevision()).toBe(staleRevision);
        expect(canonicalOoxmlFingerprint(alice.surface!.session.part())).toBe(beforeSave);
        apply(bob, [{ op: 'insertText', paragraphId: id, offset: 3, text: 'Peer', bias: 'right' }]);
        bob.surface!.setSelection({
          anchor: { paragraphId: id, offset: 7 },
          head: { paragraphId: id, offset: 7 },
        });
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of pair.editors) {
          const visible = await views(editor);
          expect(visible).toEqual([`${original}\n\nPeer`, `${original}\n\nPeer`]);
          expect(collectRevisionSites(editor.surface!.session.part())).toHaveLength(0);
          expect(
            editor.surface!.session.paragraphIds().filter((paragraphId) => paragraphId === id)
          ).toHaveLength(1);
        }
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        )!;
        expect(remote.head.nodeId).toBe(id);
        expect(remote.head.offset).toBe(4);
        expect(
          caretAt(alice.surface!.layout(), {
            paragraphId: remote.head.nodeId,
            offset: remote.head.offset,
          })
        ).not.toBeNull();
        expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(
          0
        );
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([`${original}\n\nPeer`, `${original}\n\nNewPeer`]);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await views(editor)).toEqual([`${original}\n\nPeer`, `${original}\n\nPeer`]);
        await saved(pair);
      } finally {
        pair.destroy();
      }
    });

for (const receiver of [0, 1] as const)
  test(`explicit owner deletion still wins over an offline edit to a retired paragraph; receiver ${receiver}`, async () => {
    const pair = await fixture([41, 73]);
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      insert(alice, p('New'), null, 'Alice AI');
      await pair.settle();
      const id = alice.surface!.session.paragraphIds().at(-1)!;
      apply(alice, [{ op: 'rejectRevision', revision: address(alice, 'Alice AI') }], true);
      await pair.settle();
      pair.pause();
      apply(alice, [{ op: 'deleteBlock', blockId: id }]);
      apply(bob, [{ op: 'insertText', paragraphId: id, offset: 0, text: 'Peer', bias: 'right' }]);
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(findNode(editor.surface!.session.part(), id)).toBeNull();
        expect(await views(editor)).toEqual([original, original]);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(await views(editor)).toEqual([`${original}\n\nPeer`, `${original}\n\nPeer`]);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(await views(editor)).toEqual([original, original]);
      await saved(pair);
    } finally {
      pair.destroy();
    }
  });

test('normal editing after peer recovery creates a visible new paragraph and preserves peer history', async () => {
  const pair = await fixture([41, 73]);
  try {
    const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
    insert(alice, p('New'), null, 'Alice AI');
    await pair.settle();
    const id = alice.surface!.session.paragraphIds().at(-1)!;
    pair.pause();
    apply(alice, [{ op: 'rejectRevision', revision: address(alice, 'Alice AI') }], true);
    apply(bob, [{ op: 'insertText', paragraphId: id, offset: 3, text: 'Peer', bias: 'right' }]);
    pair.reconnect();
    await pair.settle();
    apply(bob, [{ op: 'splitParagraph', paragraphId: id, offset: 4 }]);
    await pair.settle();
    const tail = bob.surface!.session.paragraphIds().at(-1)!;
    const opening = alice.surface!.session.paragraphIds()[0]!;
    apply(alice, [{ op: 'insertText', paragraphId: opening, offset: 0, text: 'Alice: ' }]);
    alice.surface!.setSelection({
      anchor: { paragraphId: tail, offset: 0 },
      head: { paragraphId: tail, offset: 0 },
    });
    await pair.settle();
    for (const editor of pair.editors) {
      expect(editor.surface!.session.paragraphIds()).toHaveLength(4);
      expect(paragraphTextOf(editor.surface!.session.part(), tail)).toBe('');
      expect(caretAt(editor.surface!.layout(), { paragraphId: tail, offset: 0 })).not.toBeNull();
    }
    const remote = pair.rooms[1]!.session.remoteSelections().find(
      (selection) => selection.actorId === 'alice'
    )!;
    expect(remote.head.nodeId).toBe(tail);
    expect(
      caretAt(bob.surface!.layout(), {
        paragraphId: remote.head.nodeId,
        offset: remote.head.offset,
      })
    ).not.toBeNull();
    expect(bob.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors) {
      expect(editor.surface!.session.paragraphIds()).toHaveLength(3);
      expect(await views(editor)).toEqual([
        `Alice: ${original}\n\nPeer`,
        `Alice: ${original}\n\nPeer`,
      ]);
    }
    expect(bob.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    for (const editor of pair.editors)
      expect(editor.surface!.session.paragraphIds()).toHaveLength(4);
    const bytes = new Uint8Array(await bob.save()),
      loaded = readOoxmlPackage(bytes);
    if (!loaded.ok) throw new Error(loaded.reason);
    const main = loaded.package.parts.get(loaded.package.mainDocumentPart)!;
    const body = main.root.children.find((node) => node.kind === 'body')!;
    expect(storyParagraphs(body)).toHaveLength(4);
    expect(strFromU8(unzipSync(bytes)['word/document.xml']!)).not.toContain(
      'urn:docx-editor.dev:review:retired-paragraph:1'
    );
    await saved(pair);
  } finally {
    pair.destroy();
  }
});
