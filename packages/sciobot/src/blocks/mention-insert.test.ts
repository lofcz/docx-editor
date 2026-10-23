import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  contentControlPropertiesOf,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { mathDocument, importedMath } from '../../test/support/math-document';
import { docxFixture } from '../../test/support/document';
import { readWorkingView } from '../review/working-view';
import { readEditorRevision } from '../projection/revision-cache';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { buildBlockIndex } from './block-index';
import { createMentionPayload, mentionWordFallback } from './mention-codec';
import { insertMention } from './mention-insert';
import { serializeMention } from '../projection/mention';
import { blockAttributeAt } from '../projection/blocks';
const person = createMentionPayload('person', 'user-7', 'Ada @ada & [notes]', {
  private: { keep: [7, true] },
});
const page = createMentionPayload('page', '/docs/lesson-1', 'Lesson\r\n 𝛼 | [one]', {
  icon: '📕',
  coverImage: 'asset:cover',
});
const mentions = (node: TargetAstNode): TargetAstNode[] =>
  node.type === 'mention'
    ? [node]
    : ((node.children as TargetAstNode[] | undefined)?.flatMap(mentions) ?? []);

test('person/document controls preserve exact metadata, Word fallback, Markdown, history and adjacent imported math', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: mathDocument(),
  });
  try {
    const surface = editor.surface!,
      before = canonicalOoxmlFingerprint(surface.session.part());
    const imported = importedMath(surface.session.currentPackage());
    const first = surface.session.paragraphIds()[0]!,
      neighbor = surface.session.paragraphIds().at(-1)!;
    const one = insertMention(editor, { payload: person, paragraphId: first, offset: 6 });
    if (!one.success) throw Error(one.error);
    const middle = canonicalOoxmlFingerprint(surface.session.part());
    const two = insertMention(editor, { payload: page, paragraphId: neighbor, offset: 0 });
    if (!two.success) throw Error(two.error);
    const both = canonicalOoxmlFingerprint(surface.session.part());
    const check = async () => {
      const bytes = new Uint8Array(await editor.save()),
        read = readOoxmlPackage(bytes);
      if (!read.ok) throw Error(read.reason);
      for (const pkg of [surface.session.currentPackage(), read.package]) {
        expect(importedMath(pkg)).toEqual(imported);
        const index = buildBlockIndex(pkg);
        expect(index.diagnostics).toEqual([]);
        for (const payload of [person, page]) {
          const block = index.byId.get(payload.id)!;
          expect(block.payload).toEqual(payload);
          expect(
            contentControlPropertiesOf(findNode(pkg.parts.get(block.partName)!, block.controlId)!)
              .lock
          ).toBe('sdtContentLocked');
        }
      }
      const pkg = surface.session.currentPackage(),
        revision = surface.session.packageRevision();
      const view = await readWorkingView(
        surface.session,
        readEditorRevision(editor, 'mentions', 'one')!
      );
      if (!view.success) throw Error(view.error);
      expect(surface.session.currentPackage()).toBe(pkg);
      expect(surface.session.packageRevision()).toBe(revision);
      const projected = view.view.projection;
      for (const payload of [person, page]) {
        expect(projected.text).toContain(serializeMention(payload).text);
        const properties = projected.attributes.filter(
          (property) => property.blockId === payload.id
        );
        expect(properties.map((p) => p.path)).toEqual([['key'], ['value']]);
        expect(properties.every((p) => p.valueStart === null && p.valueEnd === null)).toBe(true);
        const span = projected.blocks.find((block) => block.id === payload.id)!;
        expect(blockAttributeAt(projected, span.start + 1)?.property).toBe('value');
        expect(blockAttributeAt(projected, span.end - 2)?.property).toBe('key');
      }
      const parsed = parseTargetMarkdown(projected.text);
      if (!parsed.success) throw Error(parsed.error);
      expect(
        mentions(parsed.target.normalizedAst).map((node) => ({
          key: node.username,
          value: node.displayText,
        }))
      ).toEqual(
        [person, page].map((payload) => ({
          key: payload.attributes.key,
          value: payload.attributes.value,
        }))
      );
    };
    expect(paragraphTextOf(surface.session.part(), first)).toBe(
      'Before' + mentionWordFallback(person).text + '\uFFFCAfter'
    );
    expect(paragraphTextOf(surface.session.part(), neighbor)).toBe(
      mentionWordFallback(page).text + 'Neighbor'
    );
    await check();
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(middle);
    expect(buildBlockIndex(surface.session.currentPackage()).byId.has('page')).toBe(false);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(before);
    expect(buildBlockIndex(surface.session.currentPackage()).byId.size).toBe(0);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(both);
    surface.setSelection({
      anchor: { paragraphId: first, offset: 0 },
      head: { paragraphId: first, offset: 0 },
    });
    surface.insertPlainText('Lead ');
    surface.flushPendingInput();
    expect(paragraphTextOf(surface.session.part(), first)).toStartWith('Lead Before');
    await check();
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(both);
  } finally {
    editor.destroy();
  }
});

test('mention serialization preserves GFM table cells and exact identifiers without HTML promotion', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>'
    ),
  });
  try {
    const paragraphId = editor
      .surface!.session.paragraphIds()
      .find((id) => paragraphTextOf(editor.surface!.session.part(), id) === 'Cell')!;
    expect(insertMention(editor, { payload: page, paragraphId, offset: 0 }).success).toBe(true);
    const view = await readWorkingView(
      editor.surface!.session,
      readEditorRevision(editor, 'mentions', 'one')!,
      { tableSyntax: 'gfm' }
    );
    if (!view.success) throw Error(view.error);
    expect(view.view.projection.text).not.toContain('<table');
    const parsed = parseTargetMarkdown(view.view.projection.text);
    if (!parsed.success) throw Error(parsed.error);
    expect(mentions(parsed.target.normalizedAst)).toHaveLength(1);
    expect(mentions(parsed.target.normalizedAst)[0]).toMatchObject({
      username: page.attributes.key,
      displayText: page.attributes.value,
    });
  } finally {
    editor.destroy();
  }
});

test('invalid, stale, reused and suggesting insertions refuse before package writes', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: mathDocument(),
  });
  try {
    const surface = editor.surface!,
      paragraphId = surface.session.paragraphIds()[0]!;
    for (const extra of [{ offset: -1 }, { offset: 0, expectedRevision: -1 }]) {
      const before = surface.session.currentPackage();
      expect(insertMention(editor, { payload: person, paragraphId, ...extra }).success).toBe(false);
      expect(surface.session.currentPackage()).toBe(before);
    }
    expect(insertMention(editor, { payload: person, paragraphId, offset: 0 }).success).toBe(true);
    const before = surface.session.currentPackage();
    expect(insertMention(editor, { payload: person, paragraphId, offset: 0 }).success).toBe(false);
    surface.setEditingMode('suggest');
    expect(insertMention(editor, { payload: page, paragraphId, offset: 0 })).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
    expect(surface.session.currentPackage()).toBe(before);
  } finally {
    editor.destroy();
  }
});

for (const side of ['before', 'after'] as const)
  test(`native typing immediately ${side} a mention preserves atomic metadata and independent history`, () => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture('<w:p><w:r><w:t>Before After</w:t></w:r></w:p>'),
    });
    try {
      const surface = editor.surface!,
        paragraphId = surface.session.paragraphIds()[0]!;
      const payload = createMentionPayload('boundary', 'user', 'Ada');
      expect(insertMention(editor, { payload, paragraphId, offset: 7 }).success).toBe(true);
      const beforeTyping = canonicalOoxmlFingerprint(surface.session.part());
      const offset = side === 'before' ? 7 : 7 + mentionWordFallback(payload).text.length;
      surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
      surface.insertPlainText(' typed ');
      surface.flushPendingInput();
      expect(paragraphTextOf(surface.session.part(), paragraphId)).toBe(
        side === 'before' ? 'Before  typed AdaAfter' : 'Before Ada typed After'
      );
      expect(
        buildBlockIndex(surface.session.currentPackage()).byId.get(payload.id)?.payload
      ).toEqual(payload);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(surface.session.part())).toBe(beforeTyping);
      expect(
        buildBlockIndex(surface.session.currentPackage()).byId.get(payload.id)?.payload
      ).toEqual(payload);
    } finally {
      editor.destroy();
    }
  });
