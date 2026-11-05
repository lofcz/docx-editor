import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  contentControlContentNodeOf,
  findNode,
  parentNodeOf,
  paragraphTextOf,
  readOoxmlPackage,
  storyParagraphs,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { buildBlockIndex } from './block-index';
import { readPayloadStore } from './payload-opc-codec';
import { removeBlockSelection } from './remove';
import { duplicateBlockSelection } from './duplicate';
import { moveBlockSelection } from './selection-move';
import { resolveBlockSelectionRange } from './selection-range';

const MARKDOWN = [
  'Before',
  '<activity name="Warm up" duration="5">\n\nActivity body\n\n<callout type="tip">\n\nNested tip\n\n</callout>\n\n</activity>',
  'Between',
  '| First | Second |\n| --- | --- |\n| Cell A | Cell B |',
  'After',
].join('\n\n');

async function lesson() {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed</w:t></w:r></w:p>');
  let id = 0;
  const planned = planTargetDocument(input, MARKDOWN, { allocateBlockId: () => `range-${++id}` });
  if (!planned.success) throw Error(planned.error);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw Error(applied.error);
  return writeOoxmlPackage(applied.pkg);
}

function mount(bytes: Uint8Array) {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: bytes });
  return {
    editor,
    destroy: () => {
      editor.destroy();
      container.remove();
    },
  };
}
const texts = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));
const roots = (editor: DocxEditorInstance) => {
  const body = bodyStoryRoot(editor.surface!.session.part());
  if (!body || body.kind === 'textValue') throw Error('Missing body');
  return body.children.filter((node) =>
    ['paragraph', 'table', 'contentControl'].includes(node.kind)
  );
};
const fingerprint = (editor: DocxEditorInstance) =>
  canonicalOoxmlFingerprint(editor.surface!.session.part().root);

test('moving a mixed native range preserves exact subtree identities and one undo step', async () => {
  const { editor, destroy } = mount(await lesson());
  try {
    const part = editor.surface!.session.part();
    const selected = roots(editor).slice(1, 4);
    const before = fingerprint(editor);
    const request = {
      nodeIds: selected.map((node) => node.id),
      parentId: bodyStoryRoot(part)!.id,
      beforeId: null,
    };
    expect(moveBlockSelection(editor, request)).toMatchObject({ success: true, changed: true });
    expect(roots(editor).map((node) => node.id)).toEqual([
      roots(editor)[0]!.id,
      roots(editor)[1]!.id,
      ...selected.map((node) => node.id),
    ]);
    expect(texts(editor)).toEqual([
      'Before',
      'After',
      'Activity body',
      'Nested tip',
      'Between',
      'First',
      'Second',
      'Cell A',
      'Cell B',
    ]);
    for (const node of selected)
      expect(findNode(editor.surface!.session.part(), node.id)).toBe(node);
    const revision = editor.surface!.session.packageRevision();
    expect(moveBlockSelection(editor, request)).toMatchObject({ success: true, changed: false });
    expect(editor.surface!.session.packageRevision()).toBe(revision);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(editor)).toBe(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const reopened = mount(new Uint8Array(await editor.save()));
    try {
      expect(texts(reopened.editor)).toEqual(texts(editor));
    } finally {
      reopened.destroy();
    }
  } finally {
    destroy();
  }
});

test('movement refuses a selected descendant destination and keeps a paragraph in an emptied cell', async () => {
  const { editor, destroy } = mount(await lesson());
  try {
    const part = editor.surface!.session.part();
    const activity = roots(editor)[1]!;
    const content = contentControlContentNodeOf(activity)!;
    const before = fingerprint(editor);
    expect(
      moveBlockSelection(editor, { nodeIds: [activity.id], parentId: content.id, beforeId: null })
        .success
    ).toBe(false);
    expect(fingerprint(editor)).toBe(before);
    const cellParagraph = editor
      .surface!.session.paragraphIds()
      .find((id) => paragraphTextOf(part, id) === 'Cell A')!;
    const cell = parentNodeOf(part, cellParagraph)!;
    expect(
      moveBlockSelection(editor, {
        nodeIds: [cellParagraph],
        parentId: bodyStoryRoot(part)!.id,
        beforeId: roots(editor).at(-1)!.id,
      }).success
    ).toBe(true);
    const remaining = storyParagraphs(findNode(editor.surface!.session.part(), cell.id)!);
    expect(remaining).toHaveLength(1);
    expect(paragraphTextOf(editor.surface!.session.part(), remaining[0]!.id)).toBe('');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(editor)).toBe(before);
  } finally {
    destroy();
  }
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`mixed range movement retains peer typing, cursors, deletion, and history (${clientIds})`, async () => {
    const pair = await createTwoEditorFixture({ document: await lesson(), clientIds });
    try {
      const [alice, bob] = pair.editors;
      const part = alice!.surface!.session.part();
      const selected = roots(alice!).slice(1, 4);
      const paragraphId = selected[1]!.id;
      const point = { paragraphId, offset: 7 };
      bob!.surface!.setSelection({ anchor: point, head: point });
      await pair.settle();
      pair.pause();
      expect(
        moveBlockSelection(alice!, {
          nodeIds: selected.map((node) => node.id),
          parentId: bodyStoryRoot(part)!.id,
          beforeId: null,
        }).success
      ).toBe(true);
      expect(
        bob!.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId, offset: 7, text: ' peer' },
        ]).committed
      ).toBe(true);
      pair.reconnect(clientIds[0] === 41 ? 0 : 1);
      await pair.settle();
      for (const editor of pair.editors)
        expect(texts(editor).slice(0, 5)).toEqual([
          'Before',
          'After',
          'Activity body',
          'Nested tip',
          'Between peer',
        ]);
      const remote = pair.rooms[0]!.session.remoteSelections().find(
        (selection) => selection.actorId === 'bob'
      );
      expect(remote?.head.nodeId).toBe(paragraphId);
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect(texts(editor).slice(0, 4)).toEqual([
          'Before',
          'Activity body',
          'Nested tip',
          'Between peer',
        ]);
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const removed = removeBlockSelection(bob!, { nodeIds: [paragraphId] });
      if (!removed.success) throw Error(removed.error);
      await pair.settle();
      for (const editor of pair.editors)
        expect(findNode(editor.surface!.session.part(), paragraphId)).toBeNull();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });

test('mixed range duplication keeps complete sibling order and fresh nested payload identities in one undo step', async () => {
  const { editor, destroy } = mount(await lesson());
  try {
    const before = fingerprint(editor);
    const original = roots(editor).slice(1, 4);
    let id = 0;
    const copied = duplicateBlockSelection(editor, {
      nodeIds: original.map((node) => node.id),
      allocateBlockId: (type) => `copy-${type}-${++id}`,
    });
    if (!copied.success) throw Error(copied.error);
    expect(copied.nodeIds).toHaveLength(3);
    expect(new Set(copied.nodeIds).size).toBe(3);
    expect(copied.nodeIds.every((nodeId) => !original.some((node) => node.id === nodeId))).toBe(
      true
    );
    expect(Object.keys(copied.payloadIds)).toHaveLength(2);
    expect(texts(editor)).toEqual([
      'Before',
      'Activity body',
      'Nested tip',
      'Between',
      'First',
      'Second',
      'Cell A',
      'Cell B',
      'Activity body',
      'Nested tip',
      'Between',
      'First',
      'Second',
      'Cell A',
      'Cell B',
      'After',
    ]);
    const part = editor.surface!.session.part();
    expect(copied.nodeIds.map((nodeId) => findNode(part, nodeId)?.kind)).toEqual([
      'contentControl',
      'paragraph',
      'table',
    ]);
    const reopened = mount(new Uint8Array(await editor.save()));
    try {
      expect(texts(reopened.editor)).toEqual(texts(editor));
    } finally {
      reopened.destroy();
    }
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(editor)).toBe(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(texts(editor).filter((text) => text === 'Between')).toHaveLength(2);
  } finally {
    destroy();
  }
});

test('unsupported native copy references and protected owners refuse the complete range without writing', async () => {
  for (const body of [
    '<w:p><w:bookmarkStart w:id="1" w:name="Keep"/><w:r><w:t>Bookmarked</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>',
    '<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Locked child</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p/>',
  ]) {
    const input = await targetDocumentInput(body);
    const { editor, destroy } = mount(writeOoxmlPackage(input.pkg));
    try {
      const before = fingerprint(editor);
      const nodeId = editor.surface!.session.paragraphIds()[0]!;
      const copied = duplicateBlockSelection(editor, {
        nodeIds: [nodeId],
        allocateBlockId: () => 'unused',
      });
      expect(copied.success).toBe(false);
      expect(fingerprint(editor)).toBe(before);
      expect(editor.surface!.session.canUndo()).toBe(false);
    } finally {
      destroy();
    }
  }
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`copied native sibling IDs survive publication, peer deletion, and history (${clientIds})`, async () => {
    const pair = await createTwoEditorFixture({ document: await lesson(), clientIds });
    try {
      const [alice, bob] = pair.editors;
      const source = roots(alice!).slice(1, 4);
      pair.pause();
      let id = 0;
      const copied = duplicateBlockSelection(alice!, {
        nodeIds: source.map((node) => node.id),
        allocateBlockId: (type) => `peer-copy-${type}-${++id}`,
      });
      if (!copied.success) throw Error(copied.error);
      expect(
        copied.nodeIds.every((nodeId) => !!findNode(alice!.surface!.session.part(), nodeId))
      ).toBe(true);
      const copiedParagraph = copied.nodeIds[1]!;
      expect(
        alice!.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: copiedParagraph, offset: 7, text: ' copy' },
        ]).committed
      ).toBe(true);
      const removed = removeBlockSelection(bob!, {
        nodeIds: roots(bob!)
          .slice(1, 4)
          .map((node) => node.id),
      });
      if (!removed.success) throw Error(removed.error);
      pair.reconnect(clientIds[0] === 41 ? 0 : 1);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(texts(editor)).toContain('Between copy');
        expect(texts(editor)).not.toContain('Between');
      }
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) expect(texts(editor)).toEqual(['Before', 'After']);
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const digests = await pair.savedDigests();
      expect(digests[0]).toEqual(digests[1]);
      for (const editor of pair.editors) {
        const reopened = mount(new Uint8Array(await editor.save()));
        try {
          expect(texts(reopened.editor)).toEqual(texts(editor));
        } finally {
          reopened.destroy();
        }
      }
    } finally {
      pair.destroy();
    }
  });

test('native range normalizes nested overlap and order, but never expands partial or disjoint requests', async () => {
  const { editor, destroy } = mount(await lesson());
  try {
    const nodes = roots(editor);
    const part = editor.surface!.session.part();
    const nested = [
      ...buildBlockIndex(editor.surface!.session.currentPackage()).byId.values(),
    ].find((block) => block.payload.type === 'callout')!;
    const range = resolveBlockSelectionRange(part, [
      nodes[2]!.id,
      nested.controlId,
      nodes[1]!.id,
      nodes[1]!.id,
    ]);
    expect(range.success && range.nodes.map((node) => node.id)).toEqual([
      nodes[1]!.id,
      nodes[2]!.id,
    ]);
    expect(resolveBlockSelectionRange(part, [nodes[1]!.id, nodes[3]!.id])).toMatchObject({
      success: false,
    });
    const paragraph = storyParagraphs(
      contentControlContentNodeOf(findNode(part, nodes[1]!.id)!)!
    )[0]!;
    if (paragraph.kind !== 'paragraph') throw Error('Missing paragraph');
    expect(
      resolveBlockSelectionRange(part, [paragraph.children.find((node) => node.kind === 'run')!.id])
    ).toMatchObject({ success: false });
    expect(resolveBlockSelectionRange(part, [paragraph.id, nodes[2]!.id])).toMatchObject({
      success: false,
    });
  } finally {
    destroy();
  }
});

test('a complete paragraph, table, and activity range deletes in one step with exact undo and saved text', async () => {
  const { editor, destroy } = mount(await lesson());
  try {
    const before = fingerprint(editor);
    const removed = removeBlockSelection(editor, {
      nodeIds: roots(editor)
        .slice(1, 4)
        .map((node) => node.id),
    });
    if (!removed.success) throw Error(removed.error);
    expect(removed.removedPayloadIds).toHaveLength(2);
    expect(texts(editor)).toEqual(['Before', 'After']);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw Error(saved.reason);
    const store = readPayloadStore(saved.package);
    expect(store.status === 'ready' && store.entries.size).toBe(0);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(editor)).toBe(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(texts(editor)).toEqual(['Before', 'After']);
  } finally {
    destroy();
  }
});

test('whole-document deletion keeps an empty native paragraph without requiring an application payload store', async () => {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Only paragraph</w:t></w:r></w:p>');
  const { editor, destroy } = mount(writeOoxmlPackage(input.pkg));
  try {
    const before = fingerprint(editor);
    const removed = removeBlockSelection(editor, { nodeIds: roots(editor).map((node) => node.id) });
    if (!removed.success) throw Error(removed.error);
    expect(texts(editor)).toEqual(['']);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(editor)).toBe(before);
  } finally {
    destroy();
  }
});

test('section marks and user locks refuse a mixed range without partial changes', async () => {
  for (const body of [
    '<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>Section</w:t></w:r></w:p><w:p/>',
    '<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:sdt><w:sdtPr><w:lock w:val="sdtLocked"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Locked</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p/>',
  ]) {
    const input = await targetDocumentInput(body);
    const { editor, destroy } = mount(writeOoxmlPackage(input.pkg));
    try {
      const before = fingerprint(editor);
      const history = editor.surface!.session.canUndo();
      expect(
        removeBlockSelection(editor, {
          nodeIds: roots(editor)
            .slice(0, 2)
            .map((node) => node.id),
        }).success
      ).toBe(false);
      expect(fingerprint(editor)).toBe(before);
      expect(editor.surface!.session.canUndo()).toBe(history);
    } finally {
      destroy();
    }
  }
});

for (const clientIds of [
  [41, 73],
  [73, 41],
] as const)
  test(`native range deletion converges with peer text, history, and saved documents (${clientIds})`, async () => {
    const pair = await createTwoEditorFixture({ document: await lesson(), clientIds });
    try {
      const [alice, bob] = pair.editors;
      pair.pause();
      const between = bob!
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob!.surface!.session.part(), id) === 'Between')!;
      const after = bob!.surface!.session.paragraphIds().at(-1)!;
      expect(
        bob!.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: between, offset: 7, text: ' peer' },
          { op: 'insertText', paragraphId: after, offset: 5, text: ' peer' },
        ]).committed
      ).toBe(true);
      const removed = removeBlockSelection(alice!, {
        nodeIds: roots(alice!)
          .slice(1, 4)
          .map((node) => node.id),
      });
      if (!removed.success) throw Error(removed.error);
      pair.reconnect(clientIds[0] === 41 ? 0 : 1);
      await pair.settle();
      for (const editor of pair.editors) expect(texts(editor)).toEqual(['Before', 'After peer']);
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(texts(editor)).toContain('After peer');
        expect(texts(editor)).toContain('Between peer');
      }
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
      for (const editor of pair.editors) {
        const reopened = mount(new Uint8Array(await editor.save()));
        try {
          expect(texts(reopened.editor)).toEqual(['Before', 'After peer']);
        } finally {
          reopened.destroy();
        }
      }
    } finally {
      pair.destroy();
    }
  });
