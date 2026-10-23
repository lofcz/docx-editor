import { MDX_BLOCK } from '../projection/mdx';
import { insertBlock } from './insert';
import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { mathDocument, importedMath } from '../../test/support/math-document';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { planTargetDocument } from '../diff/plan-target';
import { decodeEditPlan, serializeEditPlan } from '../diff/plan';
import { applyDirectPlan } from '../ai/apply-direct';
import { targetAstKey } from '../diff/target-key';
import { createMathPayload } from './math-codec';
import { insertMath } from './math-insert';
import { buildBlockIndex } from './block-index';
const context = { documentId: 'math', epoch: 'projection' };
const mathValues = (node: TargetAstNode): unknown[] =>
  node.type === 'math' || node.type === 'inlineMath'
    ? [node.value]
    : ((node.children as TargetAstNode[] | undefined)?.flatMap(mathValues) ?? []);
async function source(editor: DocxEditorInstance) {
  const revision = readEditorRevision(editor, context.documentId, context.epoch)!;
  const view = await readWorkingView(editor.surface!.session, revision);
  if (!view.success) throw Error(view.error);
  return {
    revision,
    pkg: editor.surface!.session.currentPackage(),
    projection: view.view.projection,
  };
}
const paragraphs = (node: OoxmlNode): OoxmlNode[] =>
  node.kind === 'textValue'
    ? []
    : node.kind === 'paragraph'
      ? [node]
      : node.children.flatMap(paragraphs);

for (const expression of ['', 'x^2', '<br>', '  a\r\nb | 𝛼\n'])
  test(`native inline/block math projection preserves source, metadata, untouched OMML and empty replanning: ${JSON.stringify(expression)}`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: mathDocument(),
    });
    try {
      const first = editor.surface!.session.paragraphIds()[0]!;
      const inline = createMathPayload('inline', 'inline', expression),
        block = createMathPayload('block', 'block', expression);
      expect(
        insertMath(editor, {
          payload: inline,
          position: { kind: 'inline', paragraphId: first, offset: 6 },
        }).success
      ).toBe(true);
      expect(
        insertMath(editor, {
          payload: block,
          position: { kind: 'block', parentControlId: null, beforeId: null },
        }).success
      ).toBe(true);
      const before = editor.surface!.session.currentPackage(),
        fingerprint = canonicalOoxmlFingerprint(editor.surface!.session.part()),
        omml = importedMath(before);
      const input = await source(editor);
      expect(editor.surface!.session.currentPackage()).toBe(before);
      expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(fingerprint);
      const planned = planTargetDocument(input, input.projection.text);
      if (!planned.success) throw Error(planned.error);
      expect(planned.plan.operations).toEqual([]);
      for (const id of ['inline', 'block']) {
        expect(
          input.projection.attributes.filter((p) => p.blockId === id).map((p) => p.path)
        ).toEqual([['texExpression']]);
        const span = input.projection.blocks.find((block) => block.id === id)!;
        const parsed = parseTargetMarkdown(input.projection.text.slice(span.start, span.end));
        if (!parsed.success) throw Error(parsed.error);
        expect(mathValues(parsed.target.normalizedAst)).toEqual([expression]);
      }
      const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!loaded.ok) throw Error(loaded.reason);
      expect(importedMath(loaded.package)).toEqual(omml);
      expect(buildBlockIndex(loaded.package).byId.get('inline')?.payload).toEqual(inline);
      expect(buildBlockIndex(loaded.package).byId.get('block')?.payload).toEqual(block);
    } finally {
      editor.destroy();
    }
  });

test('registered lossless display syntax materializes original source and readable multiline Word fallback in one history unit', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
  });
  try {
    const initial = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const target = 'Keep.\n\n<scio-equation texExpression="1&#13;&#10;2" />';
    let admissions = 0;
    const planned = planTargetDocument(await source(editor), target, {
      allocateBlockId: () => 'new-equation',
      newBlockPayload: () => {
        admissions++;
        return null;
      },
    });
    expect(admissions).toBe(0);
    if (!planned.success) throw Error(planned.error);
    expect(decodeEditPlan(serializeEditPlan(planned.plan))).toEqual({
      success: true,
      plan: planned.plan,
    });
    const applied = applyDirectPlan(editor, planned.plan, context);
    if (!applied.success) throw Error(applied.error);
    const read = await source(editor),
      block = buildBlockIndex(read.pkg).byId.get('new-equation')!;
    expect(block.payload.attributes.texExpression).toBe('1\r\n2');
    expect(
      paragraphs(findNode(editor.surface!.session.part(), block.controlId)!).map((p) =>
        paragraphTextOf(editor.surface!.session.part(), p.id)
      )
    ).toEqual(['1', '2']);
    const parsed = parseTargetMarkdown(target),
      actual = parseTargetMarkdown(read.projection.text);
    if (!parsed.success || !actual.success) throw Error('Expected target syntax');
    expect(targetAstKey(actual.target.normalizedAst)).toBe(
      targetAstKey(parsed.target.normalizedAst)
    );
    const noop = planTargetDocument(read, target);
    if (!noop.success) throw Error(noop.error);
    expect(noop.plan.operations).toEqual([]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(initial);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!loaded.ok) throw Error(loaded.reason);
    expect(buildBlockIndex(loaded.package).byId.get('new-equation')?.payload).toEqual(
      block.payload
    );
  } finally {
    editor.destroy();
  }
});

for (const tag of ['equation', 'inline_equation'])
  test(`opaque React ${tag} blocks retain their complete AST, payload and unchanged native planning`, async () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture('<w:p><w:r><w:t>Keep.</w:t></w:r></w:p>'),
    });
    try {
      const payload = {
        id: 'opaque',
        type: MDX_BLOCK,
        version: 1 as const,
        attributes: { private: { keep: true } },
        mdast: {
          type: 'mdxJsxFlowElement',
          name: tag,
          attributes: [{ type: 'mdxJsxAttribute', name: 'texExpression', value: 'x' }],
          children: [],
        },
      };
      const inserted = insertBlock(editor, {
        parentControlId: null,
        beforeId: null,
        payload,
        label: 'React equation',
        paragraphs: ['React fallback'],
      });
      if (!inserted.success) throw Error(inserted.error);
      const input = await source(editor);
      const parsed = parseTargetMarkdown(input.projection.text);
      if (!parsed.success) throw Error(parsed.error);
      expect(parsed.target.normalizedAst.children[1]).toMatchObject(payload.mdast);
      const noop = planTargetDocument(input, input.projection.text);
      if (!noop.success) throw Error(noop.error);
      expect(noop.plan.operations).toEqual([]);
      const loaded = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!loaded.ok) throw Error(loaded.reason);
      expect(buildBlockIndex(loaded.package).byId.get(payload.id)?.payload).toEqual(payload);
    } finally {
      editor.destroy();
    }
  });
