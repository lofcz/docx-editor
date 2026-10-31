import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { buildBlockIndex } from './block-index';
import { duplicateBlock } from './duplicate';
import { readPayloadStore } from './payload-opc-codec';
import { removeBlocks } from './remove';

const MARKDOWN = [
  'Intro',
  '<activity name="Warm up" duration="5">\n\nFirst body\n\n<callout type="tip">\n\nNested tip\n\n</callout>\n\n</activity>',
  '<activity name="Practice" duration="10">\n\nSecond body\n\n</activity>',
  'After',
].join('\n\n');

let ids = 0;
async function lesson() {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, MARKDOWN, { allocateBlockId: () => `b-${++ids}` });
  if (!planned.success) throw new Error(planned.error);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw new Error(applied.error);
  return writeOoxmlPackage(applied.pkg);
}

const texts = (editor: DocxEditorInstance) =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextOf(editor.surface!.session.part(), id));
const activities = (pkg: OoxmlPackage) =>
  [...buildBlockIndex(pkg).byId.values()].filter(
    (block) => block.payload.type === 'lesson_activity'
  );
const records = (pkg: OoxmlPackage) => {
  const store = readPayloadStore(pkg);
  return store.status === 'ready' ? [...store.entries.keys()].sort() : [];
};

function mount(bytes: Uint8Array) {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: bytes });
  return {
    editor,
    destroy() {
      editor.destroy();
      container.remove();
    },
  };
}

test('two adjacent activities delete in one step; undo restores both with their payloads', async () => {
  const { editor, destroy } = mount(await lesson());
  try {
    const pkg = () => editor.surface!.session.currentPackage();
    const before = { texts: texts(editor), records: records(pkg()) };
    const controls = activities(pkg()).map((block) => block.controlId);
    expect(controls).toHaveLength(2);
    const removed = removeBlocks(editor, { controlIds: controls, mode: 'delete' });
    if (!removed.success) throw new Error(removed.error);
    // The nested callout's payload goes with its activity.
    expect(removed.removedPayloadIds).toHaveLength(3);
    expect(texts(editor)).toEqual(['Intro', 'After']);
    expect(records(pkg())).toEqual([]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect({ texts: texts(editor), records: records(pkg()) }).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(texts(editor)).toEqual(['Intro', 'After']);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(activities(reopened.package)).toHaveLength(0);
    // A block inside another selected block is refused: the outer one already covers it.
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const nested = [...buildBlockIndex(pkg()).byId.values()].find(
      (block) => block.payload.type === 'callout'
    )!;
    expect(
      removeBlocks(editor, { controlIds: [controls[0]!, nested.controlId], mode: 'delete' })
    ).toMatchObject({ success: false, code: 'invalid-input' });
  } finally {
    destroy();
  }
});

test('a duplicate gets fresh payload IDs for itself and its nested blocks', async () => {
  const { editor, destroy } = mount(await lesson());
  try {
    const pkg = () => editor.surface!.session.currentPackage();
    const [first] = activities(pkg());
    let next = 0;
    const copy = duplicateBlock(editor, {
      controlId: first!.controlId,
      allocateBlockId: (type) => `copy-${type}-${++next}`,
    });
    if (!copy.success) throw new Error(copy.error);
    expect(Object.values(copy.payloadIds)).toEqual(['copy-lesson_activity-1', 'copy-callout-2']);
    expect(texts(editor)).toEqual([
      'Intro',
      'First body',
      'Nested tip',
      'First body',
      'Nested tip',
      'Second body',
      'After',
    ]);
    const index = buildBlockIndex(pkg());
    expect(index.byId.get('copy-lesson_activity-1')?.payload.attributes).toEqual(
      first!.payload.attributes
    );
    expect(index.byId.get(first!.id)?.controlId).toBe(first!.controlId);
    // Editing the copy leaves the original paragraph alone.
    const copied = editor.surface!.session.paragraphIds()[3]!;
    expect(
      editor.surface!.applyAutomationOps(() => [
        { op: 'insertText', paragraphId: copied, offset: 0, text: 'Copy ' },
      ]).committed
    ).toBe(true);
    expect(texts(editor).slice(1, 4)).toEqual(['First body', 'Nested tip', 'Copy First body']);
    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(records(reopened.package)).toEqual([...records(pkg())].sort());
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(editor)).toEqual(['Intro', 'First body', 'Nested tip', 'Second body', 'After']);
    expect(records(pkg())).not.toContain('copy-lesson_activity-1');
    // Refusals: an ID in use, suggest mode.
    expect(
      duplicateBlock(editor, { controlId: first!.controlId, allocateBlockId: () => first!.id })
    ).toMatchObject({ success: false, code: 'invalid-input' });
    editor.surface!.setEditingMode('suggest');
    expect(
      duplicateBlock(editor, { controlId: first!.controlId, allocateBlockId: () => 'x-1' })
    ).toMatchObject({ success: false, code: 'unsupported-operation' });
  } finally {
    destroy();
  }
});

for (const first of [0, 1] as const)
  test(`block deletion and duplication converge beside peer typing; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({ document: await lesson() });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      const pkg = (editor: DocxEditorInstance) => editor.surface!.session.currentPackage();
      pair.pause();
      const [one, two] = activities(pkg(alice));
      const copy = duplicateBlock(alice, {
        controlId: one!.controlId,
        allocateBlockId: (type) => `dup-${type}`,
      });
      if (!copy.success) throw new Error(copy.error);
      expect(removeBlocks(alice, { controlIds: [two!.controlId], mode: 'delete' }).success).toBe(
        true
      );
      const after = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'After')!;
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: after, offset: 5, text: ' Bob' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      for (const editor of pair.editors)
        expect(texts(editor)).toEqual([
          'Intro',
          'First body',
          'Nested tip',
          'First body',
          'Nested tip',
          'After Bob',
        ]);
      expect(records(pkg(alice))).toEqual(records(pkg(bob)));
      expect(records(pkg(bob))).toContain('dup-lesson_activity');
      // Alice undoes the deletion; Bob's text stays.
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors) {
        expect(texts(editor)).toContain('Second body');
        expect(texts(editor)).toContain('After Bob');
      }
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
