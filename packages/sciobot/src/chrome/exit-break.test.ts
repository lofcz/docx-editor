import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  paragraphTextOf,
  parentNodeOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { previewDirectPlan } from '../ai/apply-direct';
import { planTargetDocument } from '../diff/plan-target';
import { exitBlock } from './exit-break';

let ids = 0;
async function documentWith(markdown: string) {
  const input = await targetDocumentInput('<w:p><w:r><w:t>Seed text</w:t></w:r></w:p>');
  const planned = planTargetDocument(input, `${markdown}\n\nAfter text`, {
    allocateBlockId: () => `block-${++ids}`,
  });
  if (!planned.success) throw new Error(planned.error);
  const applied = previewDirectPlan(input.pkg, planned.plan);
  if (!applied.success) throw new Error(applied.error);
  return writeOoxmlPackage(applied.pkg);
}

/** Body children in order: `P:<text>` for paragraphs, the element name otherwise. */
function bodyShape(editor: DocxEditorInstance) {
  const part = editor.surface!.session.part();
  const body = bodyStoryRoot(part)!;
  if (body.kind === 'textValue') return [];
  return body.children
    .filter((child) => child.kind !== 'textValue' && child.localName !== 'sectPr')
    .map((child) =>
      child.kind === 'paragraph'
        ? `P:${paragraphTextOf(part, child.id) ?? ''}`
        : child.kind === 'textValue'
          ? ''
          : child.localName
    );
}

function caretIn(editor: DocxEditorInstance, text: string) {
  const s = editor.surface!;
  const id = s.session.paragraphIds().find((p) => paragraphTextOf(s.session.part(), p) === text)!;
  const point = { paragraphId: id, offset: 1 };
  s.setSelection({ anchor: point, head: point });
}

for (const [name, markdown, inside, structure] of [
  ['table', '| Cell A | Cell B |\n| --- | --- |\n| One | Two |', 'One', 'tbl'],
  ['code block', '```python\nprint(1)\n```', 'print(1)', 'sdt'],
  ['callout', '<callout>\n\nInside callout\n\n</callout>', 'Inside callout', 'sdt'],
  [
    'activity',
    '<activity name="Warm up" duration="5">\n\nActivity body\n\n</activity>',
    'Activity body',
    'sdt',
  ],
] as const)
  test(`exit break leaves a ${name} after and before, with history and reopen`, async () => {
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({ container, document: await documentWith(markdown) });
    try {
      const initial = bodyShape(editor);
      const at = initial.indexOf(structure);
      expect(at).toBeGreaterThan(-1);
      caretIn(editor, inside);
      const after = exitBlock(editor, 'after');
      if (!after.success) throw new Error(after.error);
      const shape = bodyShape(editor);
      expect(shape[at + 1]).toBe('P:');
      expect(shape.length).toBe(initial.length + 1);
      // The caret is in the new body paragraph, outside the structure.
      const head = editor.surface!.state().selection.head;
      expect(head.paragraphId).toBe(after.paragraphId);
      const part = editor.surface!.session.part();
      expect(parentNodeOf(part, head.paragraphId)?.id).toBe(bodyStoryRoot(part)!.id);
      caretIn(editor, inside);
      const before = exitBlock(editor, 'before');
      if (!before.success) throw new Error(before.error);
      expect(bodyShape(editor)[at]).toBe('P:');
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(bodyShape(editor)).toEqual(initial);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(bodyShape(editor)[at + 1]).toBe('P:');
    } finally {
      editor.destroy();
      container.remove();
    }
  });

test('exit break refuses outside editing mode', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: await documentWith('Plain line') });
  try {
    editor.surface!.setEditingMode('suggest');
    caretIn(editor, 'Plain line');
    expect(exitBlock(editor, 'after')).toMatchObject({ success: false });
  } finally {
    editor.destroy();
    container.remove();
  }
});

for (const first of [0, 1] as const)
  test(`two sessions converge on an exit break beside peer typing; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({
      document: await documentWith('<callout>\n\nInside callout\n\n</callout>'),
    });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      pair.pause();
      caretIn(alice, 'Inside callout');
      expect(exitBlock(alice, 'after').success).toBe(true);
      const tail = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'After text')!;
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: tail, offset: 10, text: ' Bob' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      expect(bodyShape(alice)).toEqual(bodyShape(bob));
      expect(bodyShape(alice)).toContain('P:After text Bob');
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      expect(bodyShape(alice)).toEqual(bodyShape(bob));
      expect(bodyShape(alice)).toContain('P:After text Bob');
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
