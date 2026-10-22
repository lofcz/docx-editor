import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  collectRevisionSites,
  paragraphTextOf,
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
const section = '## New section\n\nFirst addition\\.\n\nSecond addition\\.';
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
    for (const decision of ['acceptRevision', 'rejectRevision'] as const)
      test(`sibling insertion ${decision}, clients ${clients.join('/')}, receiver ${receiver}: identities, peer edits, history, geometry, save`, async () => {
        const pair = await fixture(clients);
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          const [opening, closing] = alice.surface!.session.paragraphIds();
          bob.surface!.setSelection({
            anchor: { paragraphId: closing!, offset: 0 },
            head: { paragraphId: closing!, offset: 0 },
          });
          await pair.settle();
          const before = caretAt(alice.surface!.layout(), { paragraphId: closing!, offset: 0 })!;
          pair.pause();
          insert(
            alice,
            '<w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>New section</w:t></w:r></w:p>' +
              p('First addition.') +
              p('Second addition.'),
            closing!,
            'Alice AI'
          );
          insert(bob, p('Bob tail'), null, 'Bob AI');
          apply(bob, [{ op: 'insertText', paragraphId: opening!, offset: 0, text: 'Bob: ' }]);
          bob.surface!.setSelection({
            anchor: { paragraphId: closing!, offset: 0 },
            head: { paragraphId: closing!, offset: 0 },
          });
          pair.reconnect(receiver);
          await pair.settle();
          const pendingOriginal = 'Bob: ' + original;
          const pendingWorking = `Bob: Opening\\.\n\n${section}\n\nClosing\\.\n\nBob tail`;
          const check = async (expected: string[]) => {
            for (const editor of pair.editors) {
              expect(await views(editor)).toEqual(expected);
              expect(paragraphTextOf(editor.surface!.session.part(), opening!)).toBe(
                'Bob: Opening.'
              );
              expect(paragraphTextOf(editor.surface!.session.part(), closing!)).toBe('Closing.');
              expect(editor.surface!.session.paragraphIds()).toContain(opening!);
              expect(editor.surface!.session.paragraphIds()).toContain(closing!);
            }
          };
          await check([pendingOriginal, pendingWorking]);
          const aliceAddress = address(alice, 'Alice AI');
          expect(address(bob, 'Alice AI')).toEqual(aliceAddress);
          expect(address(alice, 'Bob AI').id).not.toBe(aliceAddress.id);
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          )!;
          expect(remote.head.nodeId).toBe(closing!);
          const after = caretAt(alice.surface!.layout(), {
            paragraphId: remote.head.nodeId,
            offset: remote.head.offset,
          })!;
          expect(after).not.toBeNull();
          expect(after.y).toBeGreaterThan(before.y);
          expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(
            0
          );
          await saved(pair);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check([pendingOriginal, `${pendingOriginal}\n\nBob tail`]);
          expect(address(alice, 'Bob AI')).toBeDefined();
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check([pendingOriginal, pendingWorking]);
          apply(bob, [{ op: decision, revision: aliceAddress }], true);
          await pair.settle();
          const decidedOriginal =
            decision === 'acceptRevision'
              ? `Bob: Opening\\.\n\n${section}\n\nClosing\\.`
              : pendingOriginal;
          await check([decidedOriginal, `${decidedOriginal}\n\nBob tail`]);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          await check([pendingOriginal, pendingWorking]);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          await check([decidedOriginal, `${decidedOriginal}\n\nBob tail`]);
          await saved(pair);
        } finally {
          pair.destroy();
        }
      });

for (const decision of ['acceptRevision', 'rejectRevision'] as const)
  test(`terminal insertion ${decision} retains a synchronized peer suffix and decision history`, async () => {
    const pair = await fixture([41, 73]);
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      insert(alice, p('New'), null, 'Alice AI');
      await pair.settle();
      const id = alice.surface!.session.paragraphIds().at(-1)!;
      const revision = address(alice, 'Alice AI');
      apply(bob, [{ op: 'insertText', paragraphId: id, offset: 3, text: 'Peer', bias: 'right' }]);
      await pair.settle();
      for (const editor of pair.editors)
        expect(await views(editor)).toEqual([`${original}\n\nPeer`, `${original}\n\nNewPeer`]);
      apply(alice, [{ op: decision, revision }], true);
      await pair.settle();
      const expected = `${original}\n\n${decision === 'acceptRevision' ? 'NewPeer' : 'Peer'}`;
      for (const editor of pair.editors) expect(await views(editor)).toEqual([expected, expected]);
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(await views(editor)).toEqual([`${original}\n\nPeer`, `${original}\n\nNewPeer`]);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(await views(editor)).toEqual([expected, expected]);
      await saved(pair);
    } finally {
      pair.destroy();
    }
  });

for (const receiver of [0, 1] as const)
  test(`explicit anchor deletion during a peer insertion converges without recreating that anchor; receiver ${receiver}`, async () => {
    const pair = await fixture([41, 73]);
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const closing = alice.surface!.session.paragraphIds()[1]!;
      pair.pause();
      apply(alice, [{ op: 'deleteBlock', blockId: closing }]);
      insert(bob, p('New'), closing, 'Bob AI');
      pair.reconnect(receiver);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(paragraphTextOf(editor.surface!.session.part(), closing)).toBeNull();
        expect(await views(editor)).toEqual(['Opening\\.', 'Opening\\.\n\nNew']);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(paragraphTextOf(editor.surface!.session.part(), closing)).toBe('Closing.');
        expect(await views(editor)).toEqual([original, 'Opening\\.\n\nNew\n\nClosing\\.']);
      }
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(await views(editor)).toEqual(['Opening\\.', 'Opening\\.\n\nNew']);
      await saved(pair);
    } finally {
      pair.destroy();
    }
  });
