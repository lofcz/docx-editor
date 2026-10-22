import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  findNode,
  paragraphTextOf,
  parentNodeOf,
  readOoxmlPackage,
  semanticDigest,
  serializeOoxmlPart,
  readOoxmlPart,
  withPart,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { blockClipboardDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from './block-index';
import { moveBlock, type BlockMove } from './move';
import { removeBlock } from './remove';
import { writeBlockPayload } from './payload-part';

const index = (editor: DocxEditorInstance) =>
  buildBlockIndex(editor.surface!.session.currentPackage());
const digest = (editor: DocxEditorInstance) =>
  semanticDigest(editor.surface!.session.currentPackage().parts.values());
const metadata = (editor: DocxEditorInstance) =>
  semanticDigest(
    [...editor.surface!.session.currentPackage().parts.values()].filter(
      (part) => part.name !== editor.surface!.session.part().name
    )
  );
const subtree = (editor: DocxEditorInstance, id: string) => {
  const part = editor.surface!.session.part(),
    root = findNode(part, id);
  if (!root || root.kind === 'textValue') throw new Error('Missing native element');
  return canonicalOoxmlFingerprint({ ...part, root });
};
function ignoredNamespace(part: Parameters<typeof findNode>[0], id: string): string | undefined {
  const node = findNode(part, id);
  if (!node || node.kind === 'textValue') throw new Error('Missing control');
  const prefix = node.attributes.find(
    (attribute) =>
      attribute.namespaceUri === 'http://schemas.openxmlformats.org/markup-compatibility/2006' &&
      attribute.localName === 'Ignorable'
  )!.value;
  for (let current = node; current; current = parentNodeOf(part, current.id)!) {
    const binding = current.namespaceBindings.find((binding) => binding.prefix === prefix);
    if (binding) return binding.namespaceUri;
  }
  return undefined;
}
const request = (editor: DocxEditorInstance, extra: Partial<BlockMove> = {}): BlockMove => ({
  controlId: index(editor).byId.get('activity-1')!.controlId,
  parentControlId: index(editor).byId.get('phase-2')!.controlId,
  beforeId: null,
  ...extra,
});

test('an activity moves between phases with unchanged subtree, native ordering, metadata, and exact history', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: blockClipboardDocument(),
  });
  try {
    const part = editor.surface!.session.part(),
      r = request(editor);
    const source = findNode(part, r.controlId)!,
      parent = parentNodeOf(part, source.id)!;
    const target = contentControlContentNodeOf(findNode(part, r.parentControlId!)!)!;
    const before = digest(editor),
      data = metadata(editor);
    expect(moveBlock(editor, r)).toMatchObject({ success: true, changed: true });
    expect(index(editor).byId.get('activity-1')!.parentId).toBe('phase-2');
    expect(index(editor).byId.get('note-1')!.parentId).toBe('activity-1');
    expect(findNode(editor.surface!.session.part(), source.id)).toBe(source);
    expect(findNode(editor.surface!.session.part(), parent.id)!.kind).not.toBe('textValue');
    expect(
      (findNode(editor.surface!.session.part(), parent.id) as typeof parent).children.map(
        (node) => node.id
      )
    ).toEqual([parent.children[0]!.id, parent.children[2]!.id]);
    expect(
      (findNode(editor.surface!.session.part(), target.id) as typeof target).children.map(
        (node) => node.id
      )
    ).toEqual([target.children[0]!.id, source.id]);
    expect(metadata(editor)).toEqual(data);
    const after = digest(editor);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(digest(editor)).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(digest(editor)).toEqual(after);
    const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!opened.ok) throw new Error(opened.reason);
    expect(semanticDigest(opened.package.parts.values())).toEqual(after);
  } finally {
    editor.destroy();
  }
});

test('same-parent ordering and no-op anchors preserve history while invalid moves refuse atomically', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: blockClipboardDocument(),
  });
  try {
    const r = request(editor),
      phase = index(editor).byId.get('phase-1')!;
    const parent = contentControlContentNodeOf(
      findNode(editor.surface!.session.part(), phase.controlId)!
    )!;
    const before = digest(editor),
      revision = editor.surface!.session.packageRevision();
    expect(
      moveBlock(editor, { ...r, parentControlId: phase.controlId, beforeId: r.controlId })
    ).toMatchObject({ success: true, changed: false, revision });
    expect(digest(editor)).toEqual(before);
    const reject = (extra: Partial<ReturnType<typeof request>>) => {
      const initial = digest(editor);
      expect(moveBlock(editor, { ...r, ...extra }).success).toBe(false);
      expect(digest(editor)).toEqual(initial);
    };
    reject({ parentControlId: r.controlId });
    reject({ parentControlId: index(editor).byId.get('note-1')!.controlId });
    reject({ beforeId: 'missing' });
    reject({ expectedRevision: -1 });
    reject({ controlId: 'missing' });
    editor.surface!.setEditingMode('suggest');
    reject({});
    editor.surface!.setEditingMode('view');
    reject({});
    editor.surface!.setEditingMode('edit');
    expect(
      moveBlock(editor, {
        ...r,
        parentControlId: phase.controlId,
        beforeId: parent.children[0]!.id,
      }).success
    ).toBe(true);
    expect(
      (findNode(editor.surface!.session.part(), parent.id) as typeof parent).children.map(
        (node) => node.id
      )
    ).toEqual([r.controlId, parent.children[0]!.id, parent.children[2]!.id]);
  } finally {
    editor.destroy();
  }
});

test('native source and destination locks refuse movement without changing metadata or history', () => {
  for (const [owner, lock] of [
    ['activity-1', 'sdtLocked'],
    ['phase-1', 'contentLocked'],
    ['phase-2', 'contentLocked'],
  ] as const) {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: blockClipboardDocument(),
    });
    try {
      const controlId = index(editor).byId.get(owner)!.controlId;
      expect(
        editor.surface!.applyAutomationOps(() => [
          { op: 'setContentControlProperties', controlId, lock },
        ]).rejected
      ).toBe(false);
      const before = digest(editor),
        revision = editor.surface!.session.packageRevision();
      expect(moveBlock(editor, request(editor)).success).toBe(false);
      expect(digest(editor)).toEqual(before);
      expect(editor.surface!.session.packageRevision()).toBe(revision);
    } finally {
      editor.destroy();
    }
  }
});

test('movement to the body inserts before its requested native sibling and returns in one undo', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: blockClipboardDocument(),
  });
  try {
    const before = digest(editor),
      r = request(editor),
      tail = editor.surface!.session.paragraphIds().at(-1)!;
    expect(moveBlock(editor, { ...r, parentControlId: null, beforeId: tail }).success).toBe(true);
    expect(index(editor).byId.get('activity-1')!.parentId).toBeNull();
    const body = parentNodeOf(editor.surface!.session.part(), r.controlId)!;
    expect(body.children.at(-2)!.id).toBe(r.controlId);
    expect(body.children.at(-1)!.id).toBe(tail);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(digest(editor)).toEqual(before);
  } finally {
    editor.destroy();
  }
});

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    test(`movement preserves peer text, payloads, caret identity, and exact history ${clients}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: blockClipboardDocument(),
        clientIds: clients,
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const r = request(alice),
          note = index(bob).byId.get('note-1')!;
        const paragraph = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Note text')!;
        pair.pause();
        expect(moveBlock(alice, r).success).toBe(true);
        const point = { paragraphId: paragraph, offset: 9 };
        bob.surface!.setSelection({ anchor: point, head: point });
        bob.surface!.insertPlainText(' peer');
        bob.surface!.flushPendingInput();
        expect(
          writeBlockPayload(bob, {
            controlId: note.controlId,
            payload: { ...note.payload, attributes: { ...note.payload.attributes, tone: 'peer' } },
            label: 'Peer note',
            fallbackOps: [],
          }).success
        ).toBe(true);
        const peerOnly = digest(bob),
          movedSubtree = subtree(bob, r.controlId);
        pair.reconnect(receiver);
        await pair.settle();
        for (const editor of [alice, bob]) {
          expect(index(editor).byId.get('activity-1')!.parentId).toBe('phase-2');
          expect(subtree(editor, r.controlId)).toBe(movedSubtree);
          expect(paragraphTextOf(editor.surface!.session.part(), paragraph)).toBe('Note text peer');
          expect(index(editor).byId.get('note-1')!.payload.attributes).toEqual({
            activityRef: 'activity-1',
            tone: 'peer',
          });
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          )!;
          expect(remote.head.nodeId).toBe(paragraph);
          expect(
            caretAt(editor.surface!.layout(), {
              paragraphId: paragraph,
              offset: remote.head.offset,
            })
          ).not.toBeNull();
        }
        const after = digest(alice);
        expect(after).toEqual(digest(bob));
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        expect(digest(alice)).toEqual(peerOnly);
        expect(digest(bob)).toEqual(peerOnly);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(digest(alice)).toEqual(after);
        expect(digest(bob)).toEqual(after);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    });

for (const receiver of [0, 1] as const)
  for (const deleted of ['activity-1', 'phase-2'] as const)
    test(`offline movement and peer deletion converge ${deleted}/${receiver}`, async () => {
      const pair = await createTwoEditorFixture({
        document: blockClipboardDocument(),
        clientIds: [73, 41],
      });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const r = request(alice),
          controlId = index(bob).byId.get(deleted)!.controlId;
        pair.pause();
        expect(moveBlock(alice, r).success).toBe(true);
        expect(removeBlock(bob, { controlId, mode: 'delete' }).success).toBe(true);
        pair.reconnect(receiver);
        await pair.settle();
        expect(digest(alice)).toEqual(digest(bob));
        expect(index(alice).byId.has(deleted)).toBe(false);
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
      } finally {
        pair.destroy();
      }
    });

test('opaque inherited namespace meanings survive shared movement, history, and Word reopen', async () => {
  const opened = readOoxmlPackage(blockClipboardDocument());
  if (!opened.ok) throw new Error(opened.reason);
  const part = opened.package.parts.get(opened.package.mainDocumentPart)!;
  const xml = serializeOoxmlPart(part)
    .replace(
      '<w:sdt><w:sdtPr><w:tag w:val="scio:1:phase-1"',
      '<w:sdt xmlns:q="urn:source"><w:sdtPr><w:tag w:val="scio:1:phase-1"'
    )
    .replace(
      '<w:sdt><w:sdtPr><w:tag w:val="scio:1:phase-2"',
      '<w:sdt xmlns:q="urn:target"><w:sdtPr><w:tag w:val="scio:1:phase-2"'
    )
    .replace(
      '<w:sdt><w:sdtPr><w:tag w:val="scio:1:activity-1"',
      '<w:sdt xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="q"><w:sdtPr><w:tag w:val="scio:1:activity-1"'
    );
  expect(xml).toContain('xmlns:q="urn:source"');
  expect(xml).toContain('mc:Ignorable="q"');
  const read = readOoxmlPart(xml, { name: part.name, contentType: part.contentType });
  if (!read.ok) throw new Error(read.reason);
  const pair = await createTwoEditorFixture({
    document: writeOoxmlPackage(withPart(opened.package, read.part)),
  });
  try {
    const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
    const before = digest(alice),
      r = request(alice);
    expect(ignoredNamespace(alice.surface!.session.part(), r.controlId)).toBe('urn:source');
    expect(moveBlock(alice, r).success).toBe(true);
    await pair.settle();
    for (const editor of [alice, bob]) {
      const node = findNode(editor.surface!.session.part(), r.controlId)!;
      if (node.kind === 'textValue') throw new Error('Missing control');
      expect(ignoredNamespace(editor.surface!.session.part(), r.controlId)).toBe('urn:source');
    }
    const after = digest(alice);
    expect(alice.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    expect(digest(alice)).toEqual(before);
    expect(digest(bob)).toEqual(before);
    expect(alice.exec({ type: 'redo' }).ok).toBe(true);
    await pair.settle();
    expect(digest(alice)).toEqual(after);
    expect(digest(bob)).toEqual(after);
    const saved = readOoxmlPackage(new Uint8Array(await alice.save()));
    if (!saved.ok) throw new Error(saved.reason);
    const block = buildBlockIndex(saved.package).byId.get('activity-1')!;
    const node = findNode(saved.package.parts.get(block.partName)!, block.controlId)!;
    if (node.kind === 'textValue') throw new Error('Missing saved control');
    expect(ignoredNamespace(saved.package.parts.get(block.partName)!, block.controlId)).toBe(
      'urn:source'
    );
    expect(await pair.savedDigests()).toEqual([
      semanticDigest(saved.package.parts.values()),
      semanticDigest(saved.package.parts.values()),
    ]);
  } finally {
    pair.destroy();
  }
});
