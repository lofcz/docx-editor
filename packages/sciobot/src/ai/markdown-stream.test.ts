import { runWithTransactionActor } from '../../../core/src/store/package/actor-scoped-ids';
import { removeExternalHyperlinkRelationship } from '../../../core/src/store/package/hyperlink-part';
import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  bodyStoryRoot,
  type OoxmlPart,
  canonicalOoxmlFingerprint,
  withPart,
  ensureHyperlinkRelationship,
  ensureListDefinition,
  writeOoxmlPackage,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  serializeOoxmlPart,
  storyParagraphs,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { blockClipboardDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from '../blocks/block-index';
import { beginMarkdownInsertion } from './markdown-stream';
const context = { documentId: 'markdown-stream', epoch: 'mount' };
const activity = (editor: DocxEditorInstance) =>
  buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('activity-1')!.controlId;
const texts = (editor: DocxEditorInstance) =>
  storyParagraphs(bodyStoryRoot(editor.surface!.session.part())!).map((node) =>
    paragraphTextOf(editor.surface!.session.part(), node.id)
  );
const emptyNumberingPart = (part: OoxmlPart): OoxmlPart => {
  if (part.root.kind !== 'generic') throw new Error('Expected a generic numbering root.');
  return { ...part, root: { ...part.root, children: [] } };
};
const mounted = () => {
  const container = document.createElement('div');
  document.body.append(container);
  return { container, editor: createDocxEditor({ container, document: blockClipboardDocument() }) };
};

test('split MDX, rich inline content and a nested list flush without a terminal newline', async () => {
  const { editor, container } = mounted();
  try {
    const before = new Uint8Array(await editor.save());
    let id = 0;
    const started = beginMarkdownInsertion(editor, context, {
      selectedIds: [activity(editor)],
      allocateBlockId: () => `generated-${++id}`,
    });
    if (!started.success) throw new Error(started.error);
    const markdown =
      'First **bold** with [reference](https://example.com).\n\n<activity name="Discussion" duration="5">\nTalk *together*.\n\n- One\n- Two\n</activity>\n\nFinal 😀';
    for (let at = 0; at < markdown.length; at += 3) {
      const result = started.stream.push(markdown.slice(at, at + 3));
      if (!result.success) throw new Error(result.error);
    }
    expect(texts(editor)).toContain('First bold with reference.');
    expect(texts(editor)).not.toContain('Final 😀');
    const result = started.stream.finish();
    if (!result.success) throw new Error(result.error);
    expect(started.stream.finish()).toEqual(result);
    expect(started.stream.state).toBe('ended');
    expect(started.stream.pendingMarkdown).toBe('');
    expect(texts(editor)).toEqual([
      'Intro',
      'Phase one',
      'Activity text',
      'Note text',
      'First bold with reference.',
      'Talk together.',
      'One',
      'Two',
      'Final 😀',
      'Phase end',
      'Phase two',
      'Tail',
    ]);
    expect(
      buildBlockIndex(editor.surface!.session.currentPackage()).byId.get('generated-1')!.payload
    ).toMatchObject({ type: 'lesson_activity', attributes: { name: 'Discussion', duration: 5 } });
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw new Error(saved.reason);
    const xml = serializeOoxmlPart(saved.package.parts.get('/word/document.xml')!);
    expect(xml).toContain('<w:b');
    expect(xml).toContain('<w:i');
    expect(xml).toContain('<w:numPr');
    expect([...saved.package.relationships.values()].flat()).toContainEqual(
      expect.objectContaining({ rawTarget: 'https://example.com' })
    );
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    // Numbering shell allocations are monotonic across native undo. Build the exact
    // allowed footprint independently: one hyperlink slot, one bullet list, then no link.
    const original = readOoxmlPackage(before);
    if (!original.ok) throw new Error(original.reason);
    const linked = ensureHyperlinkRelationship(original.package, 'https://example.com')!;
    const numbered = ensureListDefinition(linked.pkg, 'bullet')!;
    const resourceOnly = removeExternalHyperlinkRelationship(
      numbered.pkg,
      numbered.pkg.mainDocumentPart,
      linked.relationshipId
    )!;
    const expectedUndo = readOoxmlPackage(writeOoxmlPackage(resourceOnly));
    const undone = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!expectedUndo.ok || !undone.ok)
      throw new Error('Expected resource shell failed to reopen.');
    expect(semanticDigest(undone.package.parts.values())).toEqual(
      semanticDigest(expectedUndo.package.parts.values())
    );
    expect(undone.package.contentTypes).toEqual(expectedUndo.package.contentTypes);
    expect(undone.package.externalTargets).toEqual(expectedUndo.package.externalTargets);

    expect(editor.surface!.session.canUndo()).toBe(false);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    const redone = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!redone.ok) throw new Error(redone.reason);
    expect(semanticDigest(redone.package.parts.values())).toEqual(
      semanticDigest(saved.package.parts.values())
    );
  } finally {
    editor.destroy();
    container.remove();
  }
});

test('cancel flushes already received text before releasing the anchor', async () => {
  const { editor, container } = mounted();
  try {
    const before = new Uint8Array(await editor.save());
    const started = beginMarkdownInsertion(editor, context, { selectedIds: [activity(editor)] });
    if (!started.success) throw new Error(started.error);
    expect(started.stream.push('First\n\nIncomplete')).toMatchObject({
      success: true,
      committed: true,
    });
    expect(started.stream.pendingMarkdown).toBe('Incomplete');
    expect(started.stream.cancel().success).toBe(true);
    expect(started.stream.cancel().success).toBe(true);
    expect(started.stream.pendingMarkdown).toBe('');
    expect(started.stream.state).toBe('ended');
    expect(started.stream.push('late').success).toBe(false);
    expect(texts(editor)).toContain('First');
    expect(texts(editor)).toContain('Incomplete');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(new Uint8Array(await editor.save())).toEqual(before);
  } finally {
    editor.destroy();
    container.remove();
  }
});

test('invalid final syntax releases the anchor without a serialized placeholder', async () => {
  const { editor, container } = mounted();
  try {
    const before = new Uint8Array(await editor.save());
    const started = beginMarkdownInsertion(editor, context, { selectedIds: [activity(editor)] });
    if (!started.success) throw new Error(started.error);
    expect(started.stream.push('<activity name="unfinished').success).toBe(true);
    expect(started.stream.finish().success).toBe(false);
    expect(started.stream.state).toBe('ended');
    expect(new Uint8Array(await editor.save())).toEqual(before);
    expect(editor.surface!.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
    container.remove();
  }
});

test('split surrogate pairs flush exactly and oversized deltas refuse before writes', async () => {
  const { editor, container } = mounted();
  try {
    const before = new Uint8Array(await editor.save());
    const started = beginMarkdownInsertion(editor, context, { selectedIds: [activity(editor)] });
    if (!started.success) throw new Error(started.error);
    expect(started.stream.push('x'.repeat(2 * 1024 * 1024 + 1))).toMatchObject({
      success: false,
      code: 'invalid-input',
    });
    expect(started.stream.pendingMarkdown).toBe('');
    expect(new Uint8Array(await editor.save())).toEqual(before);
    expect(started.stream.push('Hello \uD83D').success).toBe(true);
    expect(started.stream.push('\uDE00').success).toBe(true);
    expect(started.stream.finish().success).toBe(true);
    expect(texts(editor)).toContain('Hello 😀');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(new Uint8Array(await editor.save())).toEqual(before);
  } finally {
    editor.destroy();
    container.remove();
  }
});

for (const rich of [false, true])
  for (const receiver of [0, 1] as const)
    test(`raw deltas preserve offline peer edits and final history; rich=${rich}, receiver ${receiver}`, async () => {
      const pair = await createTwoEditorFixture({ document: blockClipboardDocument() });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const started = beginMarkdownInsertion(alice, context, {
          selectedIds: [activity(alice)],
          allocateBlockId: () => `streamed-${receiver}-${rich}`,
        });
        if (!started.success) throw new Error(started.error);
        pair.pause();
        const first = rich
          ? '<activity name="Generated" duration="5">\nFirst **bold**.\n\n- One\n- Two\n</activity>'
          : 'First **bold**.';
        expect(started.stream.push(first + '\n\nNext').success).toBe(true);
        const paragraphId = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'Activity text')!;
        expect(
          bob.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId, offset: 13, text: ' Bob' },
          ]).rejected
        ).toBe(false);
        const peerOnly = await pair.savedDigests();
        const originalPeer = readOoxmlPackage(new Uint8Array(await bob.save()));
        if (!originalPeer.ok) throw new Error(originalPeer.reason);
        const numbered = rich
          ? runWithTransactionActor('alice', () =>
              ensureListDefinition(originalPeer.package, 'bullet')
            )
          : null;
        const numbering = numbered?.pkg.parts.get('/word/numbering.xml');
        const emptyNumbering = numbering && emptyNumberingPart(numbering);
        const expectedUndo =
          numbered && emptyNumbering
            ? readOoxmlPackage(writeOoxmlPackage(withPart(numbered.pkg, emptyNumbering)))
            : originalPeer;
        if (!expectedUndo.ok) throw new Error(expectedUndo.reason);

        pair.reconnect(receiver);
        await pair.settle();
        expect(started.stream.push(' *continued*.').success).toBe(true);
        expect(started.stream.finish().success).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(texts(editor)).toEqual([
            'Intro',
            'Phase one',
            'Activity text Bob',
            'Note text',
            'First bold.',
            ...(rich ? ['One', 'Two'] : []),
            'Next continued.',
            'Phase end',
            'Phase two',
            'Tail',
          ]);
        const landed = await pair.savedDigests();
        expect(landed[0]).toEqual(landed[1]);
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        const expectedDigest = rich
          ? semanticDigest(expectedUndo.package.parts.values())
          : peerOnly[1];
        if (emptyNumbering)
          for (const editor of pair.editors) {
            const opened = readOoxmlPackage(new Uint8Array(await editor.save()));
            if (!opened.ok) throw new Error(opened.reason);
            // Shared undo keeps the directory/relationship, but removes this author's definitions.
            const restoredNumbering = opened.package.parts.get('/word/numbering.xml');
            expect(restoredNumbering).toBeDefined();
            expect(canonicalOoxmlFingerprint(restoredNumbering!)).toBe(
              canonicalOoxmlFingerprint(emptyNumbering)
            );
            expect(opened.package.contentTypes).toEqual(expectedUndo.package.contentTypes);
          }

        expect(await pair.savedDigests()).toEqual([expectedDigest, expectedDigest]);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        expect(await pair.savedDigests()).toEqual(landed);
      } finally {
        pair.destroy();
      }
    });
