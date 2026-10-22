import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  directParagraphProperties,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { listDocument, listParagraph } from '../../test/support/list-document';
import { targetParagraph as p } from '../../test/support/target-document';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from '../ai/apply-direct';
import type { CanonicalEdit } from './plan';
import { planTargetDocument } from './plan-target';

const context = { documentId: 'barriers', epoch: 'mount' };
const OPAQUE = '<x:block xmlns:x="urn:opaque" x:keep="block"><x:data>opaque</x:data></x:block>';
const SECTION =
  '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:bottom="720" w:left="720" w:right="720"/></w:sectPr></w:pPr><w:r><w:t>Section.</w:t></w:r></w:p>';
const TAIL = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
function mount(document_: Uint8Array) {
  const container = document.createElement('div');
  document.body.append(container);
  return createDocxEditor({ container, document: document_ });
}
async function source(editor: DocxEditorInstance) {
  const token = readEditorRevision(editor, context.documentId, context.epoch)!,
    view = await readWorkingView(editor.surface!.session, token);
  if (!view.success) throw new Error(view.error);
  return {
    revision: token,
    pkg: editor.surface!.session.currentPackage(),
    projection: view.view.projection,
  };
}
const children = (editor: DocxEditorInstance) => {
  const root = bodyStoryRoot(editor.surface!.session.part())!;
  if (root.kind === 'textValue') throw new Error('Invalid body');
  return root.children;
};
const label = (editor: DocxEditorInstance, node: OoxmlNode) =>
  node.kind === 'paragraph'
    ? paragraphTextOf(editor.surface!.session.part(), node.id)
    : node.kind === 'textValue'
      ? ''
      : node.localName;
/** Plan, apply directly, replan empty, keep the barrier, undo in one step, and reopen. */
async function land(
  editor: DocxEditorInstance,
  target: string,
  kinds: readonly CanonicalEdit['kind'][],
  barrierOf: (editor: DocxEditorInstance) => OoxmlNode
) {
  const barrier = barrierOf(editor),
    print = canonicalOoxmlFingerprint(barrier),
    planned = planTargetDocument(await source(editor), target);
  if (!planned.success) throw new Error(`${planned.code}: ${planned.error}`);
  expect(planned.plan.operations.map((edit) => edit.kind)).toEqual([...kinds]);
  const applied = applyDirectPlan(editor, planned.plan, context);
  if (!applied.success) throw new Error(`${applied.code}: ${applied.error}`);
  const replanned = planTargetDocument(await source(editor), target);
  expect(replanned.success && replanned.plan.operations).toEqual([]);
  expect(findNode(editor.surface!.session.part(), barrier.id)).toBe(barrier);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(editor.surface!.session.canUndo()).toBe(false);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  const saved = new Uint8Array(await editor.save()),
    reopened = readOoxmlPackage(saved);
  if (!reopened.ok) throw new Error(reopened.reason);
  const again = mount(saved);
  try {
    expect(canonicalOoxmlFingerprint(barrierOf(again))).toBe(print);
    const check = planTargetDocument(await source(again), target);
    expect(check.success && check.plan.operations).toEqual([]);
  } finally {
    again.destroy();
  }
}

test('targets around an opaque block plan each physical window and keep the barrier', async () => {
  const editor = mount(docxFixture(p('Alpha.') + p('Beta.') + OPAQUE + p('Gamma.') + TAIL));
  try {
    const opaque = (value: DocxEditorInstance) =>
      children(value).find((node) => node.kind !== 'textValue' && node.localName === 'block')!;
    await land(
      editor,
      'Alpha changed.\n\nBeta.\n\nBefore barrier.\n\nGamma.\n\nAfter barrier.',
      ['text', 'siblings', 'siblings'],
      opaque
    );
    expect(children(editor).map((node) => label(editor, node))).toEqual([
      'Alpha changed.',
      'Beta.',
      'Before barrier.',
      'block',
      'Gamma.',
      'After barrier.',
      'sectPr',
    ]);
  } finally {
    editor.destroy();
  }
}, 60_000);

test('targets around a mid-body section break keep the section mark', async () => {
  const editor = mount(docxFixture(p('Alpha.') + SECTION + p('Beta.') + TAIL));
  try {
    const section = (value: DocxEditorInstance) =>
      children(value).find(
        (node) => node.kind === 'paragraph' && label(value, node) === 'Section.'
      )!;
    await land(
      editor,
      'Alpha.\n\nSection.\n\nBeta changed.\n\nNew.',
      ['text', 'siblings'],
      section
    );
  } finally {
    editor.destroy();
  }
}, 60_000);

test('a list continued across a plain paragraph keeps its Word instance for new items', async () => {
  const editor = mount(
    listDocument(
      listParagraph('One', '1') + listParagraph('Middle') + listParagraph('Two', '1') + TAIL
    )
  );
  try {
    const middle = (value: DocxEditorInstance) =>
      children(value).find((node) => node.kind === 'paragraph' && label(value, node) === 'Middle')!;
    await land(editor, '1. One\n\nMiddle\n\n2. Two!\n\n3. Three', ['text', 'siblings'], middle);
    const part = editor.surface!.session.part(),
      three = children(editor).find(
        (node) => node.kind === 'paragraph' && label(editor, node) === 'Three'
      )!;
    const numPr = directParagraphProperties(part, three.id).find(
      (item) => item.localName === 'numPr'
    );
    expect(JSON.stringify(numPr ?? null)).not.toBe('null');
  } finally {
    editor.destroy();
  }
}, 60_000);
