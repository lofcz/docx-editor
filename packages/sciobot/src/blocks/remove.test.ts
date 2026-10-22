import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  contentControlPropertiesOf,
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  withCustomXmlDataPart,
  withCustomXmlNode,
  withPart,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
  type OoxmlParagraphNode,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import {
  blockClipboardDocument,
  hostControl,
  hostParagraph,
} from '../../test/support/react-block-document';
import { bodyPart } from '../../test/support/document';
import { buildBlockIndex } from './block-index';
import { readPayloadStore, withPayloadRecord } from './payload-opc-codec';
import { removeBlock } from './remove';
import { writeBlockPayload } from './payload-part';

const blocks = (editor: DocxEditorInstance) =>
  buildBlockIndex(editor.surface!.session.currentPackage());
const digest = (editor: DocxEditorInstance) =>
  semanticDigest(editor.surface!.session.currentPackage().parts.values());
const payloadIds = (editor: DocxEditorInstance) => {
  const store = readPayloadStore(editor.surface!.session.currentPackage());
  if (store.status !== 'ready') throw new Error('Missing payload store');
  return [...store.entries.keys()].sort();
};
const texts = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));

function removalDocument(standalone: boolean): Uint8Array {
  if (!standalone) return blockClipboardDocument();
  const opened = readOoxmlPackage(blockClipboardDocument());
  if (!opened.ok) throw new Error(opened.reason);
  return writeOoxmlPackage(
    withPart(
      opened.package,
      bodyPart(
        hostControl(
          'activity-1',
          hostParagraph('Activity text') + hostControl('note-1', hostParagraph('Note text'))
        )
      )
    )
  );
}

function expectedNative(body: string, original: OoxmlPart): string {
  const identities = new Map<string, OoxmlParagraphNode['attributes']>();
  const collect = (node: OoxmlNode) => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'paragraph')
      identities.set(paragraphTextOf(original, node.id)!, node.attributes);
    node.children.forEach(collect);
  };
  collect(original.root);
  const expected = bodyPart(body);
  const preserveIdentity = (node: OoxmlNode): OoxmlNode => {
    if (node.kind === 'textValue') return node;
    const attributes =
      node.kind === 'paragraph'
        ? identities.get(paragraphTextOf(expected, node.id)!)
        : node.attributes;
    if (!attributes) throw new Error('Expected paragraph has no original identity');
    return { ...node, attributes, children: node.children.map(preserveIdentity) } as OoxmlNode;
  };
  return canonicalOoxmlFingerprint({
    ...expected,
    root: preserveIdentity(expected.root) as typeof expected.root,
  });
}

for (const mode of ['delete', 'unwrap'] as const)
  test(`${mode} preserves exact nested history, other payloads, and Word roundtrip`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: blockClipboardDocument(),
    });
    try {
      const originalPart = editor.surface!.session.part();
      const selectedParagraph = editor
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(originalPart, id) === 'Note text')!;
      const selected = { paragraphId: selectedParagraph, offset: 4 };
      editor.surface!.setSelection({ anchor: selected, head: selected });
      const before = digest(editor),
        controlId = blocks(editor).byId.get('activity-1')!.controlId;
      const result = removeBlock(editor, { controlId, mode });
      expect(result).toMatchObject({
        success: true,
        removedPayloadIds: mode === 'delete' ? ['activity-1', 'note-1'] : ['activity-1'],
      });
      expect(payloadIds(editor)).toEqual(
        mode === 'delete' ? ['phase-1', 'phase-2'] : ['note-1', 'phase-1', 'phase-2']
      );
      expect(blocks(editor).byId.has('activity-1')).toBe(false);
      const caret = editor.surface!.state().selection;
      expect(caret).not.toBeNull();
      expect(caretAt(editor.surface!.layout(), caret!.head)).not.toBeNull();
      if (mode === 'unwrap') expect(caret!.head).toEqual(selected);
      expect(texts(editor)).toEqual(
        mode === 'delete'
          ? ['Intro', 'Phase one', 'Phase end', 'Phase two', 'Tail']
          : ['Intro', 'Phase one', 'Activity text', 'Note text', 'Phase end', 'Phase two', 'Tail']
      );
      if (mode === 'unwrap') expect(blocks(editor).byId.get('note-1')!.parentId).toBe('phase-1');
      const expectedBody =
        hostParagraph('Intro') +
        hostControl(
          'phase-1',
          hostParagraph('Phase one') +
            (mode === 'unwrap'
              ? hostParagraph('Activity text') + hostControl('note-1', hostParagraph('Note text'))
              : '') +
            hostParagraph('Phase end')
        ) +
        hostControl('phase-2', hostParagraph('Phase two')) +
        hostParagraph('Tail');
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(
        expectedNative(expectedBody, originalPart)
      );
      const after = digest(editor);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(digest(editor)).toEqual(before);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(digest(editor)).toEqual(after);
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      expect(semanticDigest(reopened.package.parts.values())).toEqual(after);
    } finally {
      editor.destroy();
    }
  });

test('removal preserves unrelated custom XML and orphan payload records exactly', () => {
  const opened = readOoxmlPackage(blockClipboardDocument());
  if (!opened.ok) throw new Error(opened.reason);
  const orphan = withPayloadRecord(
    opened.package,
    { id: 'unused', type: 'note', version: 1, attributes: { keep: true } },
    'Keep orphan'
  );
  if (!orphan.ok) throw new Error(orphan.reason);
  const foreign = withCustomXmlDataPart(
    orphan.pkg,
    orphan.pkg.mainDocumentPart,
    'urn:foreign:preserve',
    'records'
  );
  if (!foreign.part) throw new Error('Missing foreign store');
  const pkg = withCustomXmlNode(foreign.pkg, foreign.part.partName, {
    id: 'keep',
    label: 'Untouched',
    data: '{"foreign":true}',
  });
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(pkg),
  });
  try {
    const foreignBefore = semanticDigest(
      [...editor.surface!.session.currentPackage().parts.values()].filter(
        (part) => part.name !== pkg.mainDocumentPart
      )
    );
    const before = digest(editor);
    const originalStore = readPayloadStore(editor.surface!.session.currentPackage());
    if (originalStore.status !== 'ready') throw new Error('Missing original store');
    const controlId = blocks(editor).byId.get('activity-1')!.controlId;
    expect(removeBlock(editor, { controlId, mode: 'delete' }).success).toBe(true);
    const store = readPayloadStore(editor.surface!.session.currentPackage());
    if (store.status !== 'ready') throw new Error('Missing store');
    expect([...store.entries]).toEqual(
      [...originalStore.entries].filter(([id]) => id !== 'activity-1' && id !== 'note-1')
    );
    expect(store.entries.get('unused')!.decoded).toMatchObject({
      status: 'supported',
      value: { attributes: { keep: true } },
    });
    expect(
      canonicalOoxmlFingerprint(
        editor.surface!.session.currentPackage().parts.get(foreign.part.partName)!
      )
    ).toBe(canonicalOoxmlFingerprint(pkg.parts.get(foreign.part.partName)!));
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(digest(editor)).toEqual(before);
    expect(
      semanticDigest(
        [...editor.surface!.session.currentPackage().parts.values()].filter(
          (part) => part.name !== pkg.mainDocumentPart
        )
      )
    ).toEqual(foreignBefore);
  } finally {
    editor.destroy();
  }
});

test('deleting the only body block leaves an editable native paragraph', async () => {
  const opened = readOoxmlPackage(blockClipboardDocument());
  if (!opened.ok) throw new Error(opened.reason);
  const body = hostControl(
    'activity-1',
    hostParagraph('Activity text') + hostControl('note-1', hostParagraph('Note text'))
  );
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(withPart(opened.package, bodyPart(body))),
  });
  try {
    const before = digest(editor);
    const controlId = blocks(editor).byId.get('activity-1')!.controlId;
    expect(removeBlock(editor, { controlId, mode: 'delete' }).success).toBe(true);
    expect(texts(editor)).toEqual(['']);
    const selection = editor.surface!.state().selection;
    expect(selection).not.toBeNull();
    expect(caretAt(editor.surface!.layout(), selection!.head)).not.toBeNull();
    const after = digest(editor);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(digest(editor)).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(digest(editor)).toEqual(after);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(semanticDigest(reopened.package.parts.values())).toEqual(after);
  } finally {
    editor.destroy();
  }
});

test('a descendant ID reused outside the target refuses removal before any package write', () => {
  const opened = readOoxmlPackage(blockClipboardDocument());
  if (!opened.ok) throw new Error(opened.reason);
  const body =
    hostParagraph('Intro') +
    hostControl(
      'activity-1',
      hostParagraph('Activity text') + hostControl('note-1', hostParagraph('Inside'))
    ) +
    hostControl('note-1', hostParagraph('Outside'));
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: writeOoxmlPackage(withPart(opened.package, bodyPart(body))),
  });
  try {
    const before = digest(editor);
    const controlId = blocks(editor).byId.get('activity-1')!.controlId;
    expect(removeBlock(editor, { controlId, mode: 'delete' })).toMatchObject({
      success: false,
      code: 'ambiguous-target',
    });
    expect(digest(editor)).toEqual(before);
  } finally {
    editor.destroy();
  }
});

test('stale, missing, locked, suggesting, and viewing removals leave the complete package unchanged', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: blockClipboardDocument(),
  });
  try {
    const controlId = blocks(editor).byId.get('activity-1')!.controlId;
    const attempt = (extra = {}) => {
      const before = digest(editor);
      expect(removeBlock(editor, { controlId, mode: 'delete', ...extra }).success).toBe(false);
      expect(digest(editor)).toEqual(before);
    };
    attempt({ expectedRevision: -1 });
    attempt({ controlId: 'missing' });
    editor.surface!.setEditingMode('suggest');
    attempt();
    editor.surface!.setEditingMode('view');
    attempt();
    editor.surface!.setEditingMode('edit');
    expect(
      editor.surface!.applyAutomationOps(() => [
        { op: 'setContentControlProperties', controlId, lock: 'sdtLocked' },
      ]).rejected
    ).toBe(false);
    attempt();
    expect(
      contentControlPropertiesOf(findNode(editor.surface!.session.part(), controlId)!).lock
    ).toBe('sdtLocked');
  } finally {
    editor.destroy();
  }
});

for (const clients of [
  [41, 73],
  [73, 41],
] as const)
  for (const receiver of [0, 1] as const)
    for (const mode of ['delete', 'unwrap'] as const)
      for (const standalone of [false, true])
        test(`${mode} converges through concurrent nested input, history, reconnect, and Word save ${clients}/${receiver}/${standalone}`, async () => {
          const pair = await createTwoEditorFixture({
            document: removalDocument(standalone),
            clientIds: clients,
          });
          try {
            const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
            const target = blocks(alice).byId.get('activity-1')!;
            const note = blocks(bob).byId.get('note-1')!;
            const paragraph = bob
              .surface!.session.paragraphIds()
              .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Note text')!;
            pair.pause();
            expect(removeBlock(alice, { controlId: target.controlId, mode }).success).toBe(true);
            const point = { paragraphId: paragraph, offset: 9 };
            bob.surface!.setSelection({ anchor: point, head: point });
            bob.surface!.insertPlainText(' peer');
            bob.surface!.flushPendingInput();
            expect(
              writeBlockPayload(bob, {
                controlId: note.controlId,
                payload: {
                  ...note.payload,
                  attributes: { ...note.payload.attributes, tone: 'peer' },
                },
                label: 'Peer note',
                fallbackOps: [],
              }).success
            ).toBe(true);
            pair.reconnect(receiver);
            await pair.settle();
            expect(digest(alice)).toEqual(digest(bob));
            for (const editor of [alice, bob]) {
              expect(blocks(editor).byId.has('activity-1')).toBe(false);
              expect(payloadIds(editor).includes('activity-1')).toBe(false);
              expect(blocks(editor).byId.has('note-1')).toBe(mode === 'unwrap');
              if (standalone && mode === 'delete') expect(texts(editor)).toEqual(['']);
              expect(paragraphTextOf(editor.surface!.session.part(), paragraph)).toBe(
                mode === 'unwrap' ? 'Note text peer' : null
              );
            }
            expect(alice.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of [alice, bob]) {
              expect(paragraphTextOf(editor.surface!.session.part(), paragraph)).toBe(
                'Note text peer'
              );
              expect(blocks(editor).byId.get('note-1')!.payload.attributes).toEqual({
                activityRef: 'activity-1',
                tone: 'peer',
              });
              const payloadStore = readPayloadStore(editor.surface!.session.currentPackage());
              if (payloadStore.status !== 'ready') throw new Error('Missing restored store');
              expect(payloadStore.entries.get('note-1')!.label).toBe('Peer note');
              expect(blocks(editor).byId.get('note-1')!.controlId).toBe(note.controlId);
              const remote = pair.rooms[0]!.session.remoteSelections().find(
                (selection) => selection.actorId === 'bob'
              );
              expect(remote).toBeDefined();
              expect(
                caretAt(editor.surface!.layout(), {
                  paragraphId: remote!.head.nodeId,
                  offset: remote!.head.offset,
                })
              ).not.toBeNull();
            }
            expect(alice.exec({ type: 'redo' }).ok).toBe(true);
            await pair.settle();
            expect(digest(alice)).toEqual(digest(bob));
            const saved = await pair.savedDigests();
            expect(saved[0]).toEqual(saved[1]);
          } finally {
            pair.destroy();
          }
        });
